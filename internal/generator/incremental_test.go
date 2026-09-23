// FilePath    : internal/generator/incremental_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 增量更新, 日期冲突与自增主键续号测试.

package generator

import (
	"errors"
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// TestIncremental 验证增量更新: 无冲突窗口可追加, 重叠窗口返回冲突错误.
func TestIncremental(t *testing.T) {
	ds, _ := data.Load()
	dir := t.TempDir()
	end := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	base := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -1600), EndDate: end,
	}
	if _, err := New(base, ds, nil).GenerateAll(); err != nil {
		t.Fatalf("base generate: %v", err)
	}

	// 无冲突窗口: 现有数据之后.
	inc := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, 1), EndDate: end.AddDate(0, 0, 120),
	}
	if _, err := IncrementalUpdate(inc, ds, nil); err != nil {
		t.Fatalf("incremental (no conflict): %v", err)
	}

	// 冲突窗口: 与现有区间重叠.
	conflict := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -100), EndDate: end.AddDate(0, 0, 50),
	}
	_, err := IncrementalUpdate(conflict, ds, nil)
	if _, ok := errors.AsType[*DateConflictError](err); !ok {
		t.Fatalf("expected DateConflictError, got %v", err)
	}
}

// TestIncrementalSaleTargets 回归测试: 增量更新后 T06 随事实区间重建 —
// 目标延伸覆盖增量窗口月份 (原计划月份转为以实绩为锚), 未超过末年年末,
// 且未变化月份的目标值保持不变 (目标系数为 (省, 年月) 的确定性函数).
func TestIncrementalSaleTargets(t *testing.T) {
	ds, _ := data.Load()
	dir := t.TempDir()
	end := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	base := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -1600), EndDate: end,
	}
	if _, err := New(base, ds, nil).GenerateAll(); err != nil {
		t.Fatalf("base generate: %v", err)
	}

	type key struct {
		pid int
		ym  int
	}
	before := make(map[key]float64)
	var baseMaxYM int
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileSaleTarget))[1:] {
		d := parseDate(rec[3])
		before[key{atoiSafe(rec[1]), d.Year()*100 + int(d.Month())}] = atofSafe(rec[4])
	}
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrder))[1:] {
		if ym := parseDate(rec[3]); ym.Year()*100+int(ym.Month()) > baseMaxYM {
			baseMaxYM = ym.Year()*100 + int(ym.Month())
		}
	}

	inc := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, 1), EndDate: end.AddDate(0, 0, 120),
	}
	if _, err := IncrementalUpdate(inc, ds, nil); err != nil {
		t.Fatalf("incremental: %v", err)
	}

	after := make(map[key]float64)
	var incWindowTarget float64
	recs := readCSVWithBOM(t, filepath.Join(dir, model.FileSaleTarget))
	for i, rec := range recs[1:] {
		d := parseDate(rec[3])
		ym := d.Year()*100 + int(d.Month())
		if ym > 202412 {
			t.Errorf("T06 行 %d: 目标月份 %d 超过末年年末 202412", i+1, ym)
		}
		k := key{atoiSafe(rec[1]), ym}
		after[k] = atofSafe(rec[4])
		if ym > baseMaxYM && ym <= 202404 {
			incWindowTarget += atofSafe(rec[4])
		}
		if want := strconv.Itoa(i + 1); rec[0] != want {
			t.Errorf("T06 行 %d: 首列 F_00_自动编号 = %q, want %q", i+1, rec[0], want)
		}
	}
	if incWindowTarget <= 0 {
		t.Errorf("增量窗口 (202402-202404) 目标合计 = %v, 期望大于 0", incWindowTarget)
	}
	for k, v := range before {
		if k.ym > baseMaxYM {
			continue // 原计划月份重建后转为实绩锚定, 允许变化.
		}
		av, ok := after[k]
		if !ok {
			t.Errorf("T06 重建后缺少省 %d 月份 %d 的目标行", k.pid, k.ym)
			continue
		}
		if av != v {
			t.Errorf("省 %d 月份 %d: 目标由 %v 变为 %v, 未变化月份的目标应保持不变", k.pid, k.ym, v, av)
		}
	}
}

// TestIncrementalAutoIDContinuity 回归测试: 增量追加后 T03/T04/T05 行字段数须与表头一致,
// 且首列 F_00_自动编号 跨全量与增量连续 (1..N) 不重复, 订单编号 OC_ 序号同样续接不重复.
func TestIncrementalAutoIDContinuity(t *testing.T) {
	ds, _ := data.Load()
	dir := t.TempDir()
	end := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	base := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -1600), EndDate: end,
	}
	if _, err := New(base, ds, nil).GenerateAll(); err != nil {
		t.Fatalf("base generate: %v", err)
	}
	inc := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, 1), EndDate: end.AddDate(0, 0, 120),
	}
	if _, err := IncrementalUpdate(inc, ds, nil); err != nil {
		t.Fatalf("incremental: %v", err)
	}

	cases := []struct {
		file   string
		header []string
	}{
		{model.FileInventory, model.HeaderInventory},
		{model.FileOrder, model.HeaderOrder},
		{model.FileOrderItem, model.HeaderOrderItem},
	}
	for _, c := range cases {
		recs := readCSVWithBOM(t, filepath.Join(dir, c.file))
		ocSeen := make(map[string]struct{})
		for i, rec := range recs[1:] {
			if len(rec) != len(c.header) {
				t.Errorf("%s 行 %d: 字段数 = %d, 表头字段数 = %d", c.file, i+1, len(rec), len(c.header))
			}
			if want := strconv.Itoa(i + 1); rec[0] != want {
				t.Errorf("%s 行 %d: 首列 F_00_自动编号 = %q, want %q", c.file, i+1, rec[0], want)
			}
			if c.file == model.FileOrder {
				if _, dup := ocSeen[rec[1]]; dup {
					t.Errorf("%s 行 %d: 订单编号 %q 重复", c.file, i+1, rec[1])
				}
				ocSeen[rec[1]] = struct{}{}
			}
		}
	}
}
