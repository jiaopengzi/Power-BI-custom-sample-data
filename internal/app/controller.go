// FilePath    : internal/app/controller.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 编排控制器, 承接界面请求并调用 generator 完成生成/增量.

package app

import (
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/generator"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// Controller 承载基础数据集, 对外提供与界面交互的编排方法.
type Controller struct {
	ds *data.Dataset
}

// New 创建控制器并预加载内嵌基础数据 (加载失败时 ds 置 nil, 与原逻辑一致).
// 返回值 *Controller, 控制器实例.
func New() *Controller {
	c := &Controller{}
	if ds, err := data.Load(); err == nil {
		c.ds = ds
	}
	return c
}

// toConfig 将界面参数转换为内部配置并校验.
//   - p, 界面参数.
//
// 返回值 *config.Config, 配置; error, 日期解析或校验失败时非 nil.
func (c *Controller) toConfig(p Params) (*config.Config, error) {
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

// GenerateSample 全量生成示例数据.
//   - p, 生成参数.
//   - prog, 进度回调, 可为 nil.
//
// 返回值 Response, 结果 (错误经 Code 归一, 不再返回 error).
func (c *Controller) GenerateSample(p Params, prog generator.ProgressFunc) Response {
	cfg, err := c.toConfig(p)
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	res, err := generator.New(cfg, c.ds, prog).GenerateAll()
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	return Response{Code: CodeOK, Result: res}
}

// IncrementalUpdate 增量更新.
//   - p, 生成参数 (日期区间为增量窗口).
//   - prog, 进度回调, 可为 nil.
//
// 返回值 Response, 结果 (无基础数据/日期冲突分别归一为对应 Code).
func (c *Controller) IncrementalUpdate(p Params, prog generator.ProgressFunc) Response {
	cfg, err := c.toConfig(p)
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	res, err := generator.IncrementalUpdate(cfg, c.ds, prog)
	switch e := err.(type) {
	case nil:
		return Response{Code: CodeOK, Result: res}
	case *generator.DateConflictError:
		return Response{Code: CodeDateConflict, Message: e.Error()}
	default:
		if err == generator.ErrNoBaseData {
			return Response{Code: CodeNoBaseData, Message: err.Error()}
		}
		return Response{Code: CodeError, Message: err.Error()}
	}
}

// HasData 判断目录中是否已存在基础示例数据 (T00/T01/T02/T04 均存在).
//   - dir, 目录路径.
//
// 返回值 bool, 存在返回 true.
func (c *Controller) HasData(dir string) bool {
	if dir == "" {
		return false
	}
	for _, f := range []string{model.FileProduct, model.FileStore, model.FileCustomer, model.FileOrder} {
		if _, err := os.Stat(filepath.Join(dir, f)); err != nil {
			return false
		}
	}
	return true
}

// DefaultOutputDir 返回默认数据存放目录 (用户主目录下的 PowerBISampleData).
// 返回值 string, 默认目录, 获取主目录失败时返回当前目录.
func (c *Controller) DefaultOutputDir() string {
	home, err := os.UserHomeDir()
	if err != nil {
		return "."
	}
	return filepath.Join(home, "PowerBISampleData")
}

// OpenOutputDir 在系统文件管理器中打开指定目录.
//   - dir, 目录路径.
//
// 返回值 error, 启动失败时非 nil.
func (c *Controller) OpenOutputDir(dir string) error {
	if dir == "" {
		return nil
	}
	var cmd *exec.Cmd
	switch runtime.GOOS {
	case "windows":
		cmd = exec.Command("explorer", dir) // #nosec G204 路径来自用户选择的目录
	case "darwin":
		cmd = exec.Command("open", dir) // #nosec G204
	default:
		cmd = exec.Command("xdg-open", dir) // #nosec G204
	}
	return cmd.Start()
}
