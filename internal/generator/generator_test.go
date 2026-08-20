// FilePath    : internal/generator/generator_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 全量生成端到端冒烟测试.

package generator

import (
	"os"
	"path/filepath"
	"testing"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// TestGenerateAll 端到端冒烟测试: 默认参数生成全部 CSV 并校验行数与文件存在性.
func TestGenerateAll(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	dir := t.TempDir()
	end := time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC)
	cfg := &config.Config{
		OutputDir:      dir,
		Locale:         config.LocaleZhCN,
		ProductCount:   8,
		StoreCount:     9,
		InventoryCycle: 14,
		StartDate:      end.AddDate(0, 0, -1600),
		EndDate:        end,
	}
	if err = cfg.Validate(); err != nil {
		t.Fatalf("validate: %v", err)
	}

	res, err := New(cfg, ds, nil).GenerateAll()
	if err != nil {
		t.Fatalf("generate: %v", err)
	}
	if res.Products != 8 {
		t.Errorf("products = %d, want 8", res.Products)
	}
	if res.Stores != 9 {
		t.Errorf("stores = %d, want 9", res.Stores)
	}
	if res.Customers == 0 || res.Orders == 0 || res.OrderItem == 0 {
		t.Errorf("unexpected zero counts: %+v", res)
	}
	for _, f := range model.AllFiles {
		if _, err := os.Stat(filepath.Join(dir, f)); err != nil {
			t.Errorf("missing file %s: %v", f, err)
		}
	}
}
