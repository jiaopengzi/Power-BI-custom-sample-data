// FilePath    : internal/generator/generator_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 全量生成端到端冒烟测试与自增主键回归测试.

package generator

import (
	"bufio"
	"encoding/csv"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
)

// generateFixture 以默认小规模参数生成一套完整示例数据.
//   - t, 测试对象.
//   - end, 生成窗口结束日期.
//
// 返回值 string, 产物目录 (t.TempDir); Result, 各表行数统计.
func generateFixture(t *testing.T, end time.Time) (string, Result) {
	t.Helper()
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	dir := t.TempDir()
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
	return dir, res
}

// readCSVWithBOM 读取带 UTF-8 BOM 的 CSV 并返回全部记录 (含表头).
// 默认按表头字段数校验每行, 行字段数与表头不一致时报错.
//   - t, 测试对象.
//   - path, 文件路径.
//
// 返回值 [][]string, 全部记录.
func readCSVWithBOM(t *testing.T, path string) [][]string {
	t.Helper()
	f, err := os.Open(path) // #nosec G305 测试临时目录
	if err != nil {
		t.Fatalf("open %s: %v", path, err)
	}
	defer f.Close() //nolint:errcheck // 只读句柄, 测试结束由系统回收
	br := bufio.NewReader(f)
	if bs, perr := br.Peek(3); perr == nil && len(bs) == 3 && bs[0] == 0xEF && bs[1] == 0xBB && bs[2] == 0xBF {
		if _, derr := br.Discard(3); derr != nil {
			t.Fatalf("discard BOM in %s: %v", path, derr)
		}
	}
	recs, err := csv.NewReader(br).ReadAll()
	if err != nil {
		t.Fatalf("read %s: %v", path, err)
	}
	return recs
}

// TestGenerateAll 端到端冒烟测试: 默认参数生成全部 CSV 并校验行数与文件存在性.
func TestGenerateAll(t *testing.T) {
	dir, res := generateFixture(t, time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC))
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

// TestAutoIDColumns 回归测试: 全部产物 CSV 的行字段数须与表头一致, 且首列 F_00_自动编号 为 1..N 连续唯一.
func TestAutoIDColumns(t *testing.T) {
	dir, _ := generateFixture(t, time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC))
	cases := []struct {
		file   string
		header []string
	}{
		{model.FileProduct, model.HeaderProduct},
		{model.FileStore, model.HeaderStore},
		{model.FileCustomer, model.HeaderCustomer},
		{model.FileInventory, model.HeaderInventory},
		{model.FileOrder, model.HeaderOrder},
		{model.FileOrderItem, model.HeaderOrderItem},
		{model.FileSaleTarget, model.HeaderSaleTarget},
		{model.FileRegion, model.HeaderRegion},
		{model.FileProvince, model.HeaderProvince},
		{model.FileCity, model.HeaderCity},
		{model.FileDistrict, model.HeaderDistrict},
	}
	for _, c := range cases {
		recs := readCSVWithBOM(t, filepath.Join(dir, c.file))
		if len(recs) < 2 {
			t.Errorf("%s: no data rows", c.file)
			continue
		}
		header := recs[0]
		for i, rec := range recs[1:] {
			if len(rec) != len(header) {
				t.Errorf("%s 行 %d: 字段数 = %d, 表头字段数 = %d", c.file, i+1, len(rec), len(header))
			}
			if want := strconv.Itoa(i + 1); rec[0] != want {
				t.Errorf("%s 行 %d: 首列 F_00_自动编号 = %q, want %q", c.file, i+1, rec[0], want)
			}
		}
	}
}

// TestStoreCityColumns 回归测试: T01 门店表地理字段对齐 VBA 版本使用城市,
// 表头须为 F_05_城市ID/F_06_城市, 且每行城市 ID 均存在于 D02 城市数据源中并与其名称一致.
func TestStoreCityColumns(t *testing.T) {
	dir, _ := generateFixture(t, time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC))
	recs := readCSVWithBOM(t, filepath.Join(dir, model.FileStore))
	header := recs[0]
	if header[5] != "F_05_城市ID" || header[6] != "F_06_城市" {
		t.Fatalf("T01 表头地理字段 = %q/%q, want F_05_城市ID/F_06_城市", header[5], header[6])
	}
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	cityName := make(map[int]string, len(ds.Cities))
	for _, c := range ds.Cities {
		cityName[c.CityID] = c.Name
	}
	for i, rec := range recs[1:] {
		id, err := strconv.Atoi(rec[5])
		if err != nil {
			t.Errorf("T01 行 %d: 城市ID %q 解析失败: %v", i+1, rec[5], err)
			continue
		}
		name, ok := cityName[id]
		if !ok {
			t.Errorf("T01 行 %d: 城市ID %d 不存在于 D02 城市数据源", i+1, id)
			continue
		}
		if rec[6] != name {
			t.Errorf("T01 行 %d: 城市名 = %q, want %q", i+1, rec[6], name)
		}
	}
}
