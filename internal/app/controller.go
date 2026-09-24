// FilePath    : internal/app/controller.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 编排控制器, 承接界面请求并调用 generator 完成生成/增量.

package app

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/generator"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// 指定目录下的两个产物子目录: data 存放 CSV 产物, pbip 存放释放的 PBIP 模板.
const (
	dataDirName = "data"
	pbipDirName = "pbip"
)

// 进度编排: 生成器内部进度 [0,100] 压缩映射到 [0,88], 预留收尾阶段 90 (PBIP 释放/日历扩张)
// 与 95 (结果统计); 100% 不在后台任务中上报, 而由界面在结果表格渲染完成后设置,
// 保证 "100% 生成完成" 与结果表格同帧出现, 避免数据量大时长时间停在 100% 造成卡住观感.
const (
	// genProgressScale 生成器内部进度的整体缩放系数 (压到 88%).
	genProgressScale = 0.88
	// pbipStagePct PBIP 释放 (全量) / 日历扩张 (增量) 阶段的进度百分比.
	pbipStagePct = 90
	// summaryStagePct 结果统计阶段 (逐表统计行数与大小, 需读完全部产物 CSV) 的进度百分比.
	summaryStagePct = 95
)

// dataDir 返回指定目录下的 CSV 数据子目录 (<dir>/data).
//   - dir, 界面指定的存放目录.
//
// 返回值 string, 数据子目录路径.
func dataDir(dir string) string {
	return filepath.Join(dir, dataDirName)
}

// pbipDir 返回指定目录下的 PBIP 模板子目录 (<dir>/pbip).
//   - dir, 界面指定的存放目录.
//
// 返回值 string, PBIP 子目录路径.
func pbipDir(dir string) string {
	return filepath.Join(dir, pbipDirName)
}

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
// CSV 产物统一落在 <指定目录>/data 子目录 (PBIP 模板的 Path 参数指向该子目录).
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
	outDir := strings.TrimSpace(p.OutputDir)
	if outDir == "" {
		return nil, fmt.Errorf("output directory is empty")
	}
	locale := config.LocaleZhCN
	if config.Locale(p.Locale) == config.LocaleEnUS {
		locale = config.LocaleEnUS
	}
	cfg := &config.Config{
		OutputDir:      dataDir(outDir),
		Locale:         locale,
		ProductCount:   p.ProductCount,
		StoreCount:     p.StoreCount,
		InventoryCycle: p.InventoryCycle,
		StartDate:      start,
		EndDate:        end,
	}
	return cfg, cfg.Validate()
}

// wrapProgress 将生成器内部进度缩放到 [0,88], 并滤除其内部完成态 stageDone,
// 使 "完成" 时机统一由收尾阶段 (PBIP 释放/结果统计) 与界面渲染接管.
//   - prog, 原始进度回调, 可为 nil.
//
// 返回值 generator.ProgressFunc, 包装后的进度回调.
func wrapProgress(prog generator.ProgressFunc) generator.ProgressFunc {
	return func(pct float64, stage string) {
		if prog == nil || stage == "stageDone" {
			return
		}
		prog(pct*genProgressScale, stage)
	}
}

// reportProgress nil 安全地上报一次阶段进度.
//   - prog, 进度回调, 可为 nil; pct, 百分比; stage, 阶段标识.
func reportProgress(prog generator.ProgressFunc, pct float64, stage string) {
	if prog == nil {
		return
	}
	prog(pct, stage)
}

// GenerateSample 全量生成示例数据: CSV 产物写入 <指定目录>/data,
// 并将内嵌 PBIP 模板释放到 <指定目录>/pbip (Path 参数改指 <指定目录>/data,
// 日历表 date_start/date_end 年份改写为生成窗口起止年份).
// 进度编排: 生成器内部进度压至 88% 以内, 收尾的 PBIP 释放 (90%) 与结果统计 (95%) 由本方法上报,
// 100% 由界面在结果表格渲染完成后设置.
// 返回值 Response, 结果 (错误经 Code 归一, 不再返回 error).
func (c *Controller) GenerateSample(p Params, prog generator.ProgressFunc) Response {
	cfg, err := c.toConfig(p)
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	res, err := generator.New(cfg, c.ds, wrapProgress(prog)).GenerateAll()
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	reportProgress(prog, pbipStagePct, "stagePbip")
	err = data.InstallPbip(pbipDir(strings.TrimSpace(p.OutputDir)), data.PbipOptions{
		DataDir:      cfg.OutputDir,
		CalStartYear: cfg.StartDate.Year(),
		CalEndYear:   cfg.EndDate.Year(),
	})
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	reportProgress(prog, summaryStagePct, "stageSummary")
	return Response{Code: CodeOK, Result: res, Tables: collectTables(cfg.OutputDir)}
}

// IncrementalUpdate 增量更新: 向 <指定目录>/data 追加数据, 其余 pbip 内容不变;
// 仅将 <指定目录>/pbip 日历表年份范围按增量窗口扩张 (起年取较小值, 止年取较大值).
//   - p, 生成参数 (日期区间为增量窗口).
//   - prog, 进度回调, 可为 nil.
//
// 进度编排与全量一致: 生成器内部进度压至 88% 以内, 日历扩张 (90%) 与结果统计 (95%) 由本方法上报,
// 100% 由界面在结果表格渲染完成后设置.
// 返回值 Response, 结果 (无基础数据/日期冲突分别归一为对应 Code).
func (c *Controller) IncrementalUpdate(p Params, prog generator.ProgressFunc) Response {
	cfg, err := c.toConfig(p)
	if err != nil {
		return Response{Code: CodeError, Message: err.Error()}
	}
	res, err := generator.IncrementalUpdate(cfg, c.ds, wrapProgress(prog))
	switch e := err.(type) {
	case nil:
		reportProgress(prog, pbipStagePct, "stagePbip")
		if cerr := data.ExpandPbipCalendar(
			pbipDir(strings.TrimSpace(p.OutputDir)), cfg.StartDate.Year(), cfg.EndDate.Year(),
		); cerr != nil {
			return Response{Code: CodeError, Message: cerr.Error()}
		}
		reportProgress(prog, summaryStagePct, "stageSummary")
		return Response{Code: CodeOK, Result: res, Tables: collectTables(cfg.OutputDir)}
	case *generator.DateConflictError:
		return Response{Code: CodeDateConflict, Message: e.Range()}
	default:
		if err == generator.ErrNoBaseData {
			return Response{Code: CodeNoBaseData, Message: err.Error()}
		}
		return Response{Code: CodeError, Message: err.Error()}
	}
}

// HasData 判断指定目录中是否已存在基础示例数据 (<dir>/data 下 T00/T01/T02/T04 均存在).
//   - dir, 目录路径.
//
// 返回值 bool, 存在返回 true.
func (c *Controller) HasData(dir string) bool {
	if dir == "" {
		return false
	}
	for _, f := range []string{model.FileProduct, model.FileStore, model.FileCustomer, model.FileOrder} {
		if _, err := os.Stat(filepath.Join(dataDir(dir), f)); err != nil {
			return false
		}
	}
	return true
}

// LastFactDate 返回事实表 (订单主表) 的最晚下单日期, 用于界面显示截止日期与增量起始校验.
//   - dir, 目录路径.
//
// 返回值 time.Time, 最晚日期; bool, 是否存在数据.
func (c *Controller) LastFactDate(dir string) (time.Time, bool) {
	if dir == "" {
		return time.Time{}, false
	}
	d, ok, err := generator.LastOrderDate(dataDir(dir))
	if err != nil {
		return time.Time{}, false
	}
	return d, ok
}

// ExistingTables 返回指定目录中已存在产物表的统计 (无基础数据时返回 nil), 供软件加载时展示历史结果.
//   - dir, 目录路径.
//
// 返回值 []TableStat, 表统计; 无数据返回 nil.
func (c *Controller) ExistingTables(dir string) []TableStat {
	if !c.HasData(dir) {
		return nil
	}
	return collectTables(dataDir(dir))
}

// ClearData 清空指定目录中已生成的全部产物: <dir>/data 下的 CSV 与 <dir>/pbip 模板目录
// (指定目录本身保留; data 子目录清空后移除, 内含其他文件时保留, pbip 目录整体移除).
//   - dir, 目录路径.
//
// 返回值 error, 删除失败时非 nil (文件不存在不视为错误).
func (c *Controller) ClearData(dir string) error {
	if dir == "" {
		return nil
	}
	// data 子目录: 删除全部产物 CSV, 已空则连同目录移除 (有其他文件时保留).
	dd := dataDir(dir)
	for _, f := range model.AllFiles {
		if err := os.Remove(filepath.Join(dd, f)); err != nil && !os.IsNotExist(err) {
			return err
		}
	}
	if isEmptyDir(dd) {
		if err := os.Remove(dd); err != nil && !os.IsNotExist(err) {
			return err
		}
	}
	// pbip 子目录: 由模板整体释放生成, 直接整体移除 (不存在时不报错).
	return os.RemoveAll(pbipDir(dir))
}

// isEmptyDir 判断目录是否为空 (目录不存在同样视为空, 便于随后移除).
//   - dir, 目录路径.
//
// 返回值 bool, 空或不存在返回 true.
func isEmptyDir(dir string) bool {
	ents, err := os.ReadDir(dir)
	if err != nil {
		return true
	}
	return len(ents) == 0
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

// collectTables 扫描输出目录下全部产物表, 汇总各表名称/行数/字节大小 (缺失的表跳过).
//   - dir, 输出目录.
//
// 返回值 []TableStat, 按 model.AllFiles 顺序排列的表统计.
func collectTables(dir string) []TableStat {
	stats := make([]TableStat, 0, len(model.AllFiles))
	for _, f := range model.AllFiles {
		path := filepath.Join(dir, f)
		fi, err := os.Stat(path)
		if err != nil {
			continue
		}
		stats = append(stats, TableStat{
			Name: strings.TrimSuffix(f, filepath.Ext(f)),
			Rows: countDataRows(path),
			Size: fi.Size(),
		})
	}
	return stats
}

// countDataRows 统计 CSV 数据行数 (总行数减去表头); 每行均以换行结尾.
//   - path, 文件路径.
//
// 返回值 int, 数据行数 (失败或空表返回 0).
func countDataRows(path string) int {
	f, err := os.Open(path) // #nosec G304 路径来自已知产物文件名
	if err != nil {
		return 0
	}

	buf := make([]byte, 256*1024)
	var lines int
	for {
		n, err := f.Read(buf)
		for _, b := range buf[:n] {
			if b == '\n' {
				lines++
			}
		}
		if err != nil {
			break
		}
	}
	if err := f.Close(); err != nil {
		return 0
	}
	if lines <= 1 {
		return 0
	}
	return lines - 1
}
