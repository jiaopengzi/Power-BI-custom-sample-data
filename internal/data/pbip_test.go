// FilePath    : internal/data/pbip_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : PBIP 模板内嵌与释放安装测试.

package data

import (
	"bytes"
	"io/fs"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// pbipIgnoredOnDisk 判断模板目录下的相对路径是否被 .gitignore 忽略,
// 当前涉及: .mimosa 目录, 以及 **/.pbi/localSettings.json 与 **/.pbi/cache.abf;
// .gitignore 调整时需同步本函数.
//   - rel, 相对模板根的斜杠分隔路径.
//
// 返回值 bool, 被忽略返回 true.
func pbipIgnoredOnDisk(rel string) bool {
	rel = filepath.ToSlash(rel)
	for _, seg := range strings.Split(rel, "/") {
		if seg == ".mimosa" {
			return true
		}
	}
	dir, base := path.Split(rel)
	if base != "localSettings.json" && base != "cache.abf" {
		return false
	}
	return strings.HasPrefix(dir, ".pbi/") || strings.Contains(dir, "/.pbi/")
}

// TestPbipEmbedComplete 校验内嵌清单与磁盘模板完全对齐: 除 .gitignore 忽略的文件外,
// 磁盘上的模板文件必须全部内嵌 (防止 go:embed 清单漏项, 如 definition.pbir),
// 已忽略文件不得内嵌, 内嵌文件必须仍存在于磁盘.
func TestPbipEmbedComplete(t *testing.T) {
	disk := make(map[string]bool)
	err := filepath.WalkDir(pbipRoot, func(p string, d fs.DirEntry, werr error) error {
		if werr != nil {
			return werr
		}
		if d.IsDir() {
			return nil
		}
		disk[strings.TrimPrefix(filepath.ToSlash(p), pbipRoot+"/")] = true
		return nil
	})
	if err != nil {
		t.Fatalf("walk disk template: %v", err)
	}

	embedded := make(map[string]bool)
	err = fs.WalkDir(pbipAssets, pbipRoot, func(p string, d fs.DirEntry, werr error) error {
		if werr != nil {
			return werr
		}
		if p == pbipRoot || d.IsDir() {
			return nil
		}
		embedded[strings.TrimPrefix(p, pbipRoot+"/")] = true
		return nil
	})
	if err != nil {
		t.Fatalf("walk embedded template: %v", err)
	}

	if len(disk) == 0 || len(embedded) == 0 {
		t.Fatalf("unexpected empty template: disk = %d files, embedded = %d files", len(disk), len(embedded))
	}
	for rel := range disk {
		if pbipIgnoredOnDisk(rel) {
			if embedded[rel] {
				t.Errorf("gitignored file %q must not be embedded", rel)
			}
			continue
		}
		if !embedded[rel] {
			t.Errorf("template file %q missing from the go:embed patterns in pbip.go", rel)
		}
	}
	for rel := range embedded {
		if !disk[rel] {
			t.Errorf("embedded file %q no longer exists on disk", rel)
		}
	}
}

// TestInstallPbip 校验模板释放: 必需文件齐全, .gitignore 忽略文件不落盘,
// Path 参数与日历表年份替换正确且其余内容原样.
func TestInstallPbip(t *testing.T) {
	root := t.TempDir()
	dir := filepath.Join(root, "pbip")
	dataDir := filepath.Join(root, "data")
	if err := InstallPbip(dir, PbipOptions{DataDir: dataDir, CalStartYear: 2022, CalEndYear: 2025}); err != nil {
		t.Fatalf("install pbip: %v", err)
	}

	for _, p := range []string{
		"demo.pbip",
		"demo.Report/.platform",
		"demo.Report/definition.pbir",
		"demo.Report/definition/report.json",
		"demo.Report/StaticResources/SharedResources/BaseThemes/CY21SU07.json",
		"demo.SemanticModel/.platform",
		"demo.SemanticModel/.pbi/editorSettings.json",
		"demo.SemanticModel/definition.pbism",
		"demo.SemanticModel/diagramLayout.json",
		"demo.SemanticModel/definition/database.tmdl",
		"demo.SemanticModel/definition/tables/T00_产品表.tmdl",
	} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(p))); err != nil {
			t.Errorf("missing %s: %v", p, err)
		}
	}

	// .gitignore 忽略的本机缓存/配置不得进入打包与释放产物.
	for _, p := range []string{
		"demo.Report/.pbi/localSettings.json",
		"demo.SemanticModel/.pbi/localSettings.json",
		"demo.SemanticModel/.pbi/cache.abf",
	} {
		if _, err := os.Stat(filepath.Join(dir, filepath.FromSlash(p))); !os.IsNotExist(err) {
			t.Errorf("%s should not be installed, stat err = %v", p, err)
		}
	}

	// Path 参数替换为数据目录, 行内其余内容与 CRLF 换行保持原样.
	b, err := os.ReadFile(filepath.Join(dir, "demo.SemanticModel", "definition", "expressions.tmdl"))
	if err != nil {
		t.Fatalf("read expressions.tmdl: %v", err)
	}
	s := string(b)
	if !strings.Contains(s, `expression Path = "`+dataDir+`"`) {
		t.Errorf("path expression not replaced with %q:\n%s", dataDir, s)
	}
	if strings.Contains(s, `C:\Users\jiaopengzi\PowerBISampleData`) {
		t.Errorf("old dev path still present:\n%s", s)
	}
	if !strings.Contains(s, "meta [IsParameterQuery=true") {
		t.Errorf("path expression tail lost:\n%s", s)
	}
	if !strings.Contains(s, "\r\n") {
		t.Errorf("CRLF line endings not preserved:\n%q", s)
	}

	// 日历表 date_start/date_end 年份替换为生成窗口年份, 其余内容 (月日/注释/缩进) 原样.
	cal, err := os.ReadFile(filepath.Join(dir, "demo.SemanticModel", "definition", "tables", "01_Calendar.tmdl"))
	if err != nil {
		t.Fatalf("read 01_Calendar.tmdl: %v", err)
	}
	cs := string(cal)
	for _, want := range []string{
		"date_start=#date(2022, 1, 1),//开始日期",
		"date_end=#date(2025, 12, 31),//结束日期",
	} {
		if !strings.Contains(cs, want) {
			t.Errorf("calendar want %q, got:\n%s", want, cs)
		}
	}
	for _, stale := range []string{"#date(2017", "#date(2022, 12, 31)"} {
		if strings.Contains(cs, stale) {
			t.Errorf("calendar stale year %q still present:\n%s", stale, cs)
		}
	}
}

// TestExpandPbipCalendar 校验增量场景的日历年份扩张: 起年取较小值, 止年取较大值, 不越界回缩.
func TestExpandPbipCalendar(t *testing.T) {
	root := t.TempDir()
	dir := filepath.Join(root, "pbip")
	if err := InstallPbip(dir, PbipOptions{DataDir: filepath.Join(root, "data"), CalStartYear: 2022, CalEndYear: 2025}); err != nil {
		t.Fatalf("install pbip: %v", err)
	}

	// 向后追加跨年: 止年扩张到 2026, 起年保持 2022.
	if err := ExpandPbipCalendar(dir, 2025, 2026); err != nil {
		t.Fatalf("expand forward: %v", err)
	}
	assertCalendarYears(t, dir, 2022, 2026)

	// 向前补历史数据: 起年回扩到 2020, 止年保持 2026.
	if err := ExpandPbipCalendar(dir, 2020, 2024); err != nil {
		t.Fatalf("expand backward: %v", err)
	}
	assertCalendarYears(t, dir, 2020, 2026)

	// 区间落在已覆盖范围内: 年份不变, 文件内容保持不变.
	calPath := filepath.Join(dir, "demo.SemanticModel", "definition", "tables", "01_Calendar.tmdl")
	before, err := os.ReadFile(calPath)
	if err != nil {
		t.Fatalf("read calendar: %v", err)
	}
	if err = ExpandPbipCalendar(dir, 2023, 2024); err != nil {
		t.Fatalf("expand noop: %v", err)
	}
	after, err := os.ReadFile(calPath)
	if err != nil {
		t.Fatalf("re-read calendar: %v", err)
	}
	if !bytes.Equal(before, after) {
		t.Error("calendar file should be unchanged when years already covered")
	}

	// pbip 目录不存在时静默跳过.
	if err := ExpandPbipCalendar(filepath.Join(root, "absent"), 2022, 2025); err != nil {
		t.Errorf("expand on absent dir should be no-op, got %v", err)
	}
}

// assertCalendarYears 断言释放目录中日历表的起止年份.
//   - t, 测试对象; dir, PBIP 释放目录; startYear, endYear, 期望起止年份.
func assertCalendarYears(t *testing.T, dir string, startYear, endYear int) {
	t.Helper()
	b, err := os.ReadFile(filepath.Join(dir, "demo.SemanticModel", "definition", "tables", "01_Calendar.tmdl"))
	if err != nil {
		t.Fatalf("read calendar: %v", err)
	}
	s := string(b)
	if !strings.Contains(s, "date_start=#date("+strconv.Itoa(startYear)+", 1, 1)") {
		t.Errorf("date_start year = %d, got:\n%s", startYear, s)
	}
	if !strings.Contains(s, "date_end=#date("+strconv.Itoa(endYear)+", 12, 31)") {
		t.Errorf("date_end year = %d, got:\n%s", endYear, s)
	}
}
