// FilePath    : internal/generator/incremental_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 增量更新与日期冲突测试.

package generator

import (
	"errors"
	"testing"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
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
	var dc *DateConflictError
	if !errors.As(err, &dc) {
		t.Fatalf("expected DateConflictError, got %v", err)
	}
}
