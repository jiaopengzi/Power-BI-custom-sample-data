// FilePath    : internal/app/controller_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 编排层目录布局测试: data 子目录, PBIP 释放与清空.

package app

import (
	"bytes"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// genParams 构造一份小规模生成参数.
//   - dir, 指定目录; start, end, 生成时间窗口.
//
// 返回值 Params, 生成参数.
func genParams(dir string, start, end time.Time) Params {
	return Params{
		OutputDir:      dir,
		Locale:         string(config.LocaleZhCN),
		ProductCount:   8,
		StoreCount:     9,
		InventoryCycle: 14,
		StartDate:      start.Format(config.DateLayout),
		EndDate:        end.Format(config.DateLayout),
	}
}

// generateFixture 全量生成一套示例数据到临时指定目录.
//   - t, 测试对象.
//
// 返回值 *Controller, 控制器; string, 指定目录; time.Time, 生成窗口结束日.
func generateFixture(t *testing.T) (*Controller, string, time.Time) {
	t.Helper()
	ctl := New()
	dir := t.TempDir()
	end := time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC)
	resp := ctl.GenerateSample(genParams(dir, end.AddDate(0, 0, -1600), end), nil)
	if resp.Code != CodeOK {
		t.Fatalf("generate: code = %s, msg = %s", resp.Code, resp.Message)
	}
	return ctl, dir, end
}

// TestGenerateSampleLayout 全量生成后: CSV 落在 <dir>/data, PBIP 模板释放到 <dir>/pbip,
// expressions.tmdl 的 Path 参数指向 <dir>/data, 日历表年份改写为生成窗口起止年份.
func TestGenerateSampleLayout(t *testing.T) {
	ctl, dir, end := generateFixture(t)
	startYear := end.AddDate(0, 0, -1600).Year()

	for _, f := range model.AllFiles {
		if _, err := os.Stat(filepath.Join(dir, "data", f)); err != nil {
			t.Errorf("missing data csv %s: %v", f, err)
		}
	}
	if !ctl.HasData(dir) {
		t.Error("HasData should be true after generation")
	}
	for _, name := range []string{"demo.pbip", "demo.Report", "demo.SemanticModel"} {
		if _, err := os.Stat(filepath.Join(dir, "pbip", name)); err != nil {
			t.Errorf("missing pbip entry %s: %v", name, err)
		}
	}

	b, err := os.ReadFile(filepath.Join(dir, "pbip", "demo.SemanticModel", "definition", "expressions.tmdl"))
	if err != nil {
		t.Fatalf("read expressions.tmdl: %v", err)
	}
	want := `expression Path = "` + filepath.Join(dir, "data") + `"`
	if !strings.Contains(string(b), want) {
		t.Errorf("path expression want %s, got:\n%s", want, b)
	}

	cal, err := os.ReadFile(filepath.Join(dir, "pbip", "demo.SemanticModel", "definition", "tables", "01_Calendar.tmdl"))
	if err != nil {
		t.Fatalf("read 01_Calendar.tmdl: %v", err)
	}
	for _, want := range []string{
		"date_start=#date(" + strconv.Itoa(startYear) + ", 1, 1)",
		"date_end=#date(" + strconv.Itoa(end.Year()) + ", 12, 31)",
	} {
		if !strings.Contains(string(cal), want) {
			t.Errorf("calendar want %q, got:\n%s", want, cal)
		}
	}
}

// TestIncrementalKeepsPbip 增量更新仅向 <dir>/data 追加数据; pbip 中除日历表年份按窗口扩张外,
// 其余内容 (如 expressions.tmdl) 保持不变.
func TestIncrementalKeepsPbip(t *testing.T) {
	ctl, dir, end := generateFixture(t)
	startYear := end.AddDate(0, 0, -1600).Year()

	exprPath := filepath.Join(dir, "pbip", "demo.SemanticModel", "definition", "expressions.tmdl")
	before, err := os.ReadFile(exprPath)
	if err != nil {
		t.Fatalf("read expressions.tmdl: %v", err)
	}
	orderPath := filepath.Join(dir, "data", model.FileOrder)
	rowsBefore := countDataRows(orderPath)

	cutoff, ok := ctl.LastFactDate(dir)
	if !ok {
		t.Fatal("no fact cutoff after generation")
	}
	// 增量窗口拉长到跨年, 验证日历止年随之扩张.
	incStart := cutoff.AddDate(0, 0, 1)
	incEnd := incStart.AddDate(0, 0, 400)
	resp := ctl.IncrementalUpdate(genParams(dir, incStart, incEnd), nil)
	if resp.Code != CodeOK {
		t.Fatalf("incremental: code = %s, msg = %s", resp.Code, resp.Message)
	}

	if rowsAfter := countDataRows(orderPath); rowsAfter <= rowsBefore {
		t.Errorf("orders after incremental = %d, want > %d", rowsAfter, rowsBefore)
	}
	after, err := os.ReadFile(exprPath)
	if err != nil {
		t.Fatalf("re-read expressions.tmdl: %v", err)
	}
	if !bytes.Equal(before, after) {
		t.Error("pbip expressions.tmdl changed during incremental update")
	}

	cal, err := os.ReadFile(filepath.Join(dir, "pbip", "demo.SemanticModel", "definition", "tables", "01_Calendar.tmdl"))
	if err != nil {
		t.Fatalf("read 01_Calendar.tmdl: %v", err)
	}
	for _, want := range []string{
		"date_start=#date(" + strconv.Itoa(startYear) + ", 1, 1)",
		"date_end=#date(" + strconv.Itoa(incEnd.Year()) + ", 12, 31)",
	} {
		if !strings.Contains(string(cal), want) {
			t.Errorf("calendar want %q, got:\n%s", want, cal)
		}
	}
}

// TestClearData 清空后 <dir>/data 与 <dir>/pbip 均被移除, 指定目录本身保留.
func TestClearData(t *testing.T) {
	ctl, dir, _ := generateFixture(t)

	if err := ctl.ClearData(dir); err != nil {
		t.Fatalf("clear: %v", err)
	}
	for _, sub := range []string{"data", "pbip"} {
		if _, err := os.Stat(filepath.Join(dir, sub)); !os.IsNotExist(err) {
			t.Errorf("%s dir should be removed, stat err = %v", sub, err)
		}
	}
	if _, err := os.Stat(dir); err != nil {
		t.Errorf("output dir should be kept: %v", err)
	}
	if ctl.HasData(dir) {
		t.Error("HasData should be false after clear")
	}
}
