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
