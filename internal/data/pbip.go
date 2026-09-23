// FilePath    : internal/data/pbip.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 内嵌 PBIP 模板 (demo.pbip / demo.Report / demo.SemanticModel) 的释放安装.

package data

import (
	"embed"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// pbipAssets 内嵌 PBIP 模板, 供生成示例数据时原样释放到指定目录.
// 刻意逐条列出匹配而不是 all:assets/pbip 整树: .gitignore 忽略的 **/.pbi/localSettings.json
// 与 **/.pbi/cache.abf (本地缓存可达上百 MB) 不能进入打包产物;
// 而 .platform / .pbi/editorSettings.json 等以点开头的隐藏文件又是模板必需,
// 普通匹配会整体排除点开头文件, 故涉及的子树使用 all: 前缀 (这些子树内不存在被忽略文件).
// 清单与磁盘模板的对齐由 TestPbipEmbedComplete 守护: 除 .gitignore 忽略的文件外必须全部内嵌.
//
//go:embed assets/pbip/demo.pbip assets/pbip/demo.Report/.platform assets/pbip/demo.Report/definition.pbir all:assets/pbip/demo.Report/definition all:assets/pbip/demo.Report/StaticResources assets/pbip/demo.SemanticModel/.platform assets/pbip/demo.SemanticModel/.pbi/editorSettings.json assets/pbip/demo.SemanticModel/definition.pbism assets/pbip/demo.SemanticModel/diagramLayout.json all:assets/pbip/demo.SemanticModel/definition
var pbipAssets embed.FS

// pbipRoot 内嵌模板在 embed.FS 中的根路径.
const pbipRoot = "assets/pbip"

// pbipPathExprRel Path 参数表达式所在文件 (相对模板根).
// 该表达式被各表 M 分区引用 (Path&"\T00_产品表.csv"), 指向 CSV 数据目录.
const pbipPathExprRel = "demo.SemanticModel/definition/expressions.tmdl"

// pbipCalendarRel 日历表 M 分区所在文件 (相对模板根),
// 其 date_start/date_end 两个 #date() 的年份需随生成窗口年份改写.
const pbipCalendarRel = "demo.SemanticModel/definition/tables/01_Calendar.tmdl"

// pbipPathRe 匹配 Path 参数表达式, 捕获组为引号前的前缀, 引号内为待替换的旧路径.
var pbipPathRe = regexp.MustCompile(`(expression\s+Path\s*=\s*)"[^"]*"`)

// pbipCalStartRe / pbipCalEndRe 匹配日历表 date_start/date_end 的年份, 捕获组为 4 位年份.
var (
	pbipCalStartRe = regexp.MustCompile(`date_start\s*=\s*#date\(\s*(\d{4})`)
	pbipCalEndRe   = regexp.MustCompile(`date_end\s*=\s*#date\(\s*(\d{4})`)
)

// PbipOptions PBIP 模板释放 (InstallPbip) 时的动态替换参数.
type PbipOptions struct {
	// DataDir expressions.tmdl 中 Path 参数的取值 (<指定目录>/data).
	DataDir string
	// CalStartYear, CalEndYear 01_Calendar.tmdl 中 date_start/date_end 的年份, 取生成窗口起止年份.
	CalStartYear int
	CalEndYear   int
}

// InstallPbip 将内嵌 PBIP 模板原封不动地释放到 dir (覆盖同名旧条目), 并完成两处动态替换:
// expressions.tmdl 的 Path 参数改指 DataDir; 01_Calendar.tmdl 的日期起止年份改为生成窗口年份.
//   - dir, PBIP 释放目录 (<指定目录>/pbip).
//   - opt, 动态替换参数.
//
// 返回值 error, 出错时非 nil.
func InstallPbip(dir string, opt PbipOptions) error {
	if err := os.MkdirAll(dir, 0o750); err != nil {
		return err
	}
	return fs.WalkDir(pbipAssets, pbipRoot, func(p string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if p == pbipRoot {
			return nil
		}
		rel := strings.TrimPrefix(p, pbipRoot+"/")
		target := filepath.Join(dir, filepath.FromSlash(rel))
		if d.IsDir() {
			return os.MkdirAll(target, 0o750)
		}
		content, rerr := pbipFileContent(rel, opt)
		if rerr != nil {
			return rerr
		}
		return os.WriteFile(target, content, 0o600) // #nosec G306 模板文件, 无执行需求
	})
}

// pbipFileContent 读取内嵌模板文件内容, 并按需完成动态替换 (Path 参数与日历表年份).
//   - rel, 相对模板根的文件路径.
//   - opt, 动态替换参数.
//
// 返回值 []byte, 处理后的内容; error, 出错时非 nil.
func pbipFileContent(rel string, opt PbipOptions) ([]byte, error) {
	content, err := pbipAssets.ReadFile(pbipRoot + "/" + rel)
	if err != nil {
		return nil, err
	}
	switch rel {
	case pbipPathExprRel:
		return replacePbipPath(content, opt.DataDir)
	case pbipCalendarRel:
		return replacePbipCalendarYears(content, opt.CalStartYear, opt.CalEndYear)
	}
	return content, nil
}

// ExpandPbipCalendar 增量更新后扩张已释放日历表的年份范围: 起年取与 startYear 的较小值,
// 止年取与 endYear 的较大值 (增量窗口必在现有数据区间之外, 只扩张不缩小, 避免砍掉已有数据的日期).
// pbip 目录尚未释放 (如被用户删除) 时静默跳过.
//   - dir, PBIP 释放目录 (<指定目录>/pbip).
//   - startYear, endYear, 增量窗口起止年份.
//
// 返回值 error, 出错时非 nil.
func ExpandPbipCalendar(dir string, startYear, endYear int) error {
	path := filepath.Join(dir, filepath.FromSlash(pbipCalendarRel))
	content, err := os.ReadFile(path) // #nosec G304 路径来自受控释放目录
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		return err
	}
	curStart, curEnd, err := readPbipCalendarYears(content)
	if err != nil {
		return err
	}
	newStart, newEnd := min(curStart, startYear), max(curEnd, endYear)
	if newStart == curStart && newEnd == curEnd {
		return nil
	}
	if content, err = replacePbipCalendarYears(content, newStart, newEnd); err != nil {
		return err
	}
	// #nosec G306,G703 模板文件无执行需求; 路径由受控释放目录拼常量模板子路径而来
	return os.WriteFile(path, content, 0o600)
}

// replacePbipPath 将 expressions.tmdl 内容中 Path 参数表达式的取值替换为 dataDir,
// 其余字节 (含换行符与编码) 保持原样.
//   - content, 原文件内容.
//   - dataDir, 数据目录路径.
//
// 返回值 []byte, 替换后的内容; error, 未找到 Path 表达式时非 nil.
func replacePbipPath(content []byte, dataDir string) ([]byte, error) {
	loc := pbipPathRe.FindSubmatchIndex(content)
	if loc == nil {
		return nil, fmt.Errorf("path expression not found in %s", pbipPathExprRel)
	}
	// loc[2]/loc[3] 为前缀捕获组, 整体匹配止于 loc[1], 故 [loc[3], loc[1]) 即含引号的旧取值.
	out := make([]byte, 0, len(content)+len(dataDir))
	out = append(out, content[:loc[3]]...)
	out = append(out, '"')
	out = append(out, dataDir...)
	out = append(out, '"')
	return append(out, content[loc[1]:]...), nil
}

// replacePbipCalendarYears 将日历表内容中 date_start/date_end 的年份分别替换为 startYear/endYear,
// 其余字节 (含换行符/注释与月日取值) 保持原样.
//   - content, 原文件内容.
//   - startYear, endYear, 起止年份.
//
// 返回值 []byte, 替换后的内容; error, 任一年份表达式未找到时非 nil.
func replacePbipCalendarYears(content []byte, startYear, endYear int) ([]byte, error) {
	content, err := replacePbipYear(content, pbipCalStartRe, startYear, "date_start")
	if err != nil {
		return nil, err
	}
	return replacePbipYear(content, pbipCalEndRe, endYear, "date_end")
}

// readPbipCalendarYears 读取日历表内容中 date_start/date_end 的当前年份.
//   - content, 文件内容.
//
// 返回值 int, 起始年份; int, 结束年份; error, 任一年份表达式未找到或非法时非 nil.
func readPbipCalendarYears(content []byte) (int, int, error) {
	startYear, err := readPbipYear(content, pbipCalStartRe, "date_start")
	if err != nil {
		return 0, 0, err
	}
	endYear, err := readPbipYear(content, pbipCalEndRe, "date_end")
	if err != nil {
		return 0, 0, err
	}
	return startYear, endYear, nil
}

// replacePbipYear 将 content 中首个匹配 re 的年份捕获组替换为指定年份.
//   - content, 文件内容; re, 年份正则; year, 新年份; name, 出错时定位用的表达式名.
//
// 返回值 []byte, 替换后的内容; error, 未匹配到年份时非 nil.
func replacePbipYear(content []byte, re *regexp.Regexp, year int, name string) ([]byte, error) {
	loc := re.FindSubmatchIndex(content)
	if loc == nil {
		return nil, fmt.Errorf("%s year not found in %s", name, pbipCalendarRel)
	}
	out := make([]byte, 0, len(content))
	out = append(out, content[:loc[2]]...)
	out = strconv.AppendInt(out, int64(year), 10)
	return append(out, content[loc[3]:]...), nil
}

// readPbipYear 读取 content 中首个匹配 re 的年份捕获组.
//   - content, 文件内容; re, 年份正则; name, 出错时定位用的表达式名.
//
// 返回值 int, 年份; error, 未匹配或解析失败时非 nil.
func readPbipYear(content []byte, re *regexp.Regexp, name string) (int, error) {
	loc := re.FindSubmatchIndex(content)
	if loc == nil {
		return 0, fmt.Errorf("%s year not found in %s", name, pbipCalendarRel)
	}
	year, err := strconv.Atoi(string(content[loc[2]:loc[3]]))
	if err != nil {
		return 0, fmt.Errorf("parse %s year in %s: %w", name, pbipCalendarRel, err)
	}
	return year, nil
}
