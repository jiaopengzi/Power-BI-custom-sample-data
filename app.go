// FilePath    : app.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : Wails 后端 API, 绑定生成/增量/目录/语言等方法给前端.

package main

import (
	"context"
	"os"
	"os/exec"
	stdruntime "runtime"
	"time"

	wruntime "github.com/wailsapp/wails/v2/pkg/runtime"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/generator"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// App 承载 Wails 后端上下文与绑定给前端的 API 方法.
type App struct {
	ctx context.Context
	ds  *data.Dataset
}

// NewApp 创建 App 实例.
// 返回值 *App, 应用实例.
func NewApp() *App {
	return &App{}
}

// startup 由 Wails 在启动时调用, 保存上下文并预加载基础数据.
//   - ctx, 应用上下文.
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
	if ds, err := data.Load(); err == nil {
		a.ds = ds
	}
}

// Params 前端提交的生成参数, 日期以 YYYY-MM-DD 字符串传递.
type Params struct {
	OutputDir      string `json:"outputDir"`
	Locale         string `json:"locale"`
	ProductCount   int    `json:"productCount"`
	StoreCount     int    `json:"storeCount"`
	InventoryCycle int    `json:"inventoryCycle"`
	StartDate      string `json:"startDate"`
	EndDate        string `json:"endDate"`
}

// Response 统一返回结构.
//   - Code, 结果码: ok / no_base_data / date_conflict / error.
//   - Message, 附加消息 (如冲突区间).
//   - Result, 行数统计.
type Response struct {
	Code    string           `json:"code"`
	Message string           `json:"message"`
	Result  generator.Result `json:"result"`
}

// 响应结果码.
const (
	codeOK           = "ok"
	codeError        = "error"
	codeNoBaseData   = "no_base_data"
	codeDateConflict = "date_conflict"
)

// toConfig 将前端参数转换为内部配置.
//   - p, 前端参数.
//
// 返回值 *config.Config, 配置; error, 校验失败时非 nil.
func (a *App) toConfig(p Params) (*config.Config, error) {
	start, err := time.Parse(config.DateLayout, p.StartDate)
	if err != nil {
		return nil, err
	}
	end, err := time.Parse(config.DateLayout, p.EndDate)
	if err != nil {
		return nil, err
	}
	locale := config.LocaleZhCN
	if config.Locale(p.Locale) == config.LocaleEnUS {
		locale = config.LocaleEnUS
	}
	cfg := &config.Config{
		OutputDir:      p.OutputDir,
		Locale:         locale,
		ProductCount:   p.ProductCount,
		StoreCount:     p.StoreCount,
		InventoryCycle: p.InventoryCycle,
		StartDate:      start,
		EndDate:        end,
	}
	return cfg, cfg.Validate()
}

// emitProgress 生成进度回调, 向前端发送 generate:progress 事件.
// 返回值 generator.ProgressFunc, 进度回调.
func (a *App) emitProgress() generator.ProgressFunc {
	return func(percent float64, stage string) {
		if a.ctx == nil {
			return
		}
		wruntime.EventsEmit(a.ctx, "generate:progress", map[string]any{
			"percent": percent,
			"stage":   stage,
		})
	}
}

// GenerateSample 全量生成示例数据.
//   - p, 生成参数.
//
// 返回值 Response, 结果; error, 参数非法或生成失败时非 nil.
func (a *App) GenerateSample(p Params) (Response, error) {
	cfg, err := a.toConfig(p)
	if err != nil {
		return Response{Code: codeError, Message: err.Error()}, nil
	}
	res, err := generator.New(cfg, a.ds, a.emitProgress()).GenerateAll()
	if err != nil {
		return Response{Code: codeError, Message: err.Error()}, nil
	}
	return Response{Code: codeOK, Result: res}, nil
}

// IncrementalUpdate 增量更新.
//   - p, 生成参数 (日期区间为增量窗口).
//
// 返回值 Response, 结果; error, 系统级错误时非 nil.
func (a *App) IncrementalUpdate(p Params) (Response, error) {
	cfg, err := a.toConfig(p)
	if err != nil {
		return Response{Code: codeError, Message: err.Error()}, nil
	}
	res, err := generator.IncrementalUpdate(cfg, a.ds, a.emitProgress())
	switch e := err.(type) {
	case nil:
		return Response{Code: codeOK, Result: res}, nil
	case *generator.DateConflictError:
		return Response{Code: codeDateConflict, Message: e.Error()}, nil
	default:
		if err == generator.ErrNoBaseData {
			return Response{Code: codeNoBaseData, Message: err.Error()}, nil
		}
		return Response{Code: codeError, Message: err.Error()}, nil
	}
}

// SelectOutputDir 弹出目录选择对话框.
// 返回值 string, 选中的目录; error, 出错时非 nil.
func (a *App) SelectOutputDir() (string, error) {
	if a.ctx == nil {
		return "", nil
	}
	return wruntime.OpenDirectoryDialog(a.ctx, wruntime.OpenDialogOptions{
		Title: "选择数据存放目录",
	})
}

// OpenOutputDir 在系统文件管理器中打开指定目录.
//   - dir, 目录路径.
//
// 返回值 error, 出错时非 nil.
func (a *App) OpenOutputDir(dir string) error {
	if dir == "" {
		return nil
	}
	var cmd *exec.Cmd
	switch stdruntime.GOOS {
	case "windows":
		cmd = exec.Command("explorer", dir) // #nosec G204 路径来自用户选择的目录
	case "darwin":
		cmd = exec.Command("open", dir) // #nosec G204
	default:
		cmd = exec.Command("xdg-open", dir) // #nosec G204
	}
	return cmd.Start()
}

// HasData 判断目录中是否已存在基础示例数据.
//   - dir, 目录路径.
//
// 返回值 bool, 存在返回 true.
func (a *App) HasData(dir string) bool {
	if dir == "" {
		return false
	}
	for _, f := range []string{model.FileProduct, model.FileStore, model.FileCustomer, model.FileOrder} {
		if _, err := os.Stat(dir + string(os.PathSeparator) + f); err != nil {
			return false
		}
	}
	return true
}

// DefaultOutputDir 返回默认数据存放目录 (用户主目录下的子目录).
// 返回值 string, 默认目录.
func (a *App) DefaultOutputDir() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return "."
	}
	return home + string(os.PathSeparator) + "PowerBISampleData"
}
