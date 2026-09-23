// FilePath    : internal/generator/generator_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 全量生成端到端冒烟测试与自增主键回归测试.

package generator

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
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

// TestProductNameFormat 回归测试: 产品名称为 产品 + 3 位字母数字混合组合 (如 产品A7X),
// 组合须同时含字母与数字且全局唯一; 产品编号保持 SKU_ + 6 位数字 (如 SKU_000122).
func TestProductNameFormat(t *testing.T) {
	dir, _ := generateFixture(t, time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC))
	recs := readCSVWithBOM(t, filepath.Join(dir, model.FileProduct))
	nameRe := regexp.MustCompile(`^产品[A-Z0-9]{3}$`)
	codeRe := regexp.MustCompile(`^SKU_\d{6,}$`)
	isDigit := func(r rune) bool { return r >= '0' && r <= '9' }
	isLetter := func(r rune) bool { return r >= 'A' && r <= 'Z' }
	seen := make(map[string]struct{}, len(recs))
	for i, rec := range recs[1:] {
		name := rec[3]
		if !nameRe.MatchString(name) {
			t.Errorf("T00 行 %d: 产品名称 = %q, want 产品 + 3 位字母数字组合", i+1, name)
			continue
		}
		suffix := strings.TrimPrefix(name, "产品")
		if !strings.ContainsFunc(suffix, isDigit) || !strings.ContainsFunc(suffix, isLetter) {
			t.Errorf("T00 行 %d: 产品名称 = %q, 组合须同时含字母与数字", i+1, name)
		}
		if _, dup := seen[name]; dup {
			t.Errorf("T00 行 %d: 产品名称 = %q 重复", i+1, name)
		}
		seen[name] = struct{}{}
		if !codeRe.MatchString(rec[1]) {
			t.Errorf("T00 行 %d: 产品编号 = %q, want SKU_ + 至少 6 位数字", i+1, rec[1])
		}
	}
}

// TestProductPriceLognormal 回归测试: 产品售价服从对数正态分布并截断在 [1000, 30000],
// 均值大于中位数 (右偏, 多数偏低少数高价), 成本价不超过售价, 分类基本覆盖 A-J.
func TestProductPriceLognormal(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	end := time.Date(2025, 3, 1, 0, 0, 0, 0, time.UTC)
	cfg := &config.Config{
		OutputDir: t.TempDir(), Locale: config.LocaleZhCN,
		ProductCount: 1000, StoreCount: 1, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -60), EndDate: end,
	}
	if err = cfg.Validate(); err != nil {
		t.Fatalf("validate: %v", err)
	}
	g := New(cfg, ds, nil)
	g.genProducts()

	cats := make(map[string]struct{})
	prices := make([]float64, 0, len(g.products))
	for _, p := range g.products {
		if p.salePrice < 1000 || p.salePrice > 30000 {
			t.Errorf("产品 %s: 售价 = %v, 超出 [1000, 30000]", p.code, p.salePrice)
		}
		if p.costPrice > p.salePrice {
			t.Errorf("产品 %s: 成本价 %v 高于售价 %v", p.code, p.costPrice, p.salePrice)
		}
		cats[p.category] = struct{}{}
		prices = append(prices, p.salePrice)
	}
	sort.Float64s(prices)
	mean := 0.0
	for _, v := range prices {
		mean += v
	}
	mean /= float64(len(prices))
	median := prices[len(prices)/2]
	if mean < median*1.2 {
		t.Errorf("售价均值 = %.2f, 中位数 = %.2f, 对数正态右偏应满足均值明显大于中位数", mean, median)
	}
	if len(cats) < 8 {
		t.Errorf("产品分类数 = %d, 期望基本覆盖 A-J", len(cats))
	}
}

// TestSaleTargets 回归测试: T06 销售目标须覆盖 首个事实年度的下一年 1 月 至 最后事实年度的 12 月
// (如事实 [2022-08, 2026-08] 对应目标 [2023-01, 2026-12], 事实首年不出目标, 末年为完整年度),
// 事实期内目标以实绩为锚 (完成率落在 [90%, 120%]), 事实期之后的计划月份有外推目标.
func TestSaleTargets(t *testing.T) {
	end := time.Date(2025, 6, 1, 0, 0, 0, 0, time.UTC)
	dir, _ := generateFixture(t, end)
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}

	// 门店编号 -> 省 ID.
	storeProvince := make(map[string]int)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileStore))[1:] {
		storeProvince[rec[1]] = ds.CityToProvince[atoiSafe(rec[5])]
	}

	// 订单编号 -> (省 ID, 月份 YYYYMM), 并统计事实年度与事实月份区间.
	ocProvinceYM := make(map[string]int)
	factYears := make(map[int]struct{})
	minYM, maxYM := 0, 0
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrder))[1:] {
		d := parseDate(rec[3])
		ym := d.Year()*100 + int(d.Month())
		ocProvinceYM[rec[1]] = ym
		factYears[d.Year()] = struct{}{}
		if minYM == 0 || ym < minYM {
			minYM = ym
		}
		if ym > maxYM {
			maxYM = ym
		}
	}
	if len(factYears) < 3 {
		t.Fatalf("测试窗口事实年度数 = %d, 期望至少 3 年以验证全年度覆盖", len(factYears))
	}

	// 目标区间: 首个事实年度的下一年 1 月 至 最后事实年度的 12 月.
	wantStartYM := (minYM/100+1)*100 + 1
	wantEndYM := maxYM/100*100 + 12

	// 实绩: (省 ID, 月份 YYYYMM) -> 销售金额. 订单主表的门店编号经 T01 换算省 ID.
	// 分省与整体汇总仅统计目标区间内 (次年 1 月起) 的实绩, 首个事实年度无目标不参与完成率.
	orderStore := make(map[string]string)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrder))[1:] {
		orderStore[rec[1]] = rec[2]
	}
	type ymKey struct {
		pid int
		ym  int
	}
	actual := make(map[ymKey]float64)
	provActual := make(map[int]float64)
	var totalActual float64
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrderItem))[1:] {
		pid := storeProvince[orderStore[rec[1]]]
		ym := ocProvinceYM[rec[1]]
		amount := atofSafe(rec[6])
		actual[ymKey{pid: pid, ym: ym}] += amount
		if ym >= wantStartYM {
			provActual[pid] += amount
			totalActual += amount
		}
	}

	// 目标: (省 ID, 月份 YYYYMM) -> 销售目标, 并校验月份区间与年度覆盖.
	target := make(map[ymKey]float64)
	provTarget := make(map[int]float64) // 仅事实期内 (锚定月份) 的目标合计
	var totalTarget float64             // 仅事实期内 (锚定月份) 的目标合计
	var planTarget float64              // 事实期之后计划月份的目标合计
	targetYears := make(map[int]struct{})
	for i, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileSaleTarget))[1:] {
		pid := atoiSafe(rec[1])
		d := parseDate(rec[3])
		ym := d.Year()*100 + int(d.Month())
		if ym < wantStartYM || ym > wantEndYM {
			t.Errorf("T06 行 %d: 目标月份 %d 超出预期区间 [%d, %d]", i+1, ym, wantStartYM, wantEndYM)
		}
		targetYears[d.Year()] = struct{}{}
		v := atofSafe(rec[4])
		target[ymKey{pid: pid, ym: ym}] += v
		switch {
		case ym <= maxYM:
			provTarget[pid] += v
			totalTarget += v
		default:
			planTarget += v
		}
	}
	if _, ok := targetYears[wantStartYM/100]; !ok {
		t.Errorf("T06 缺少起始年度 %d 的目标", wantStartYM/100)
	}
	if _, ok := targetYears[wantEndYM/100]; !ok {
		t.Errorf("T06 缺少结束年度 %d 的目标", wantEndYM/100)
	}
	if _, ok := targetYears[minYM/100]; ok {
		t.Errorf("T06 不应包含首个事实年度 %d 的目标", minYM/100)
	}
	if planTarget <= 0 {
		t.Errorf("事实期之后计划月份的目标合计 = %v, 期望大于 0", planTarget)
	}

	// 完成率校验: 事实期内逐省逐月, 逐省汇总与整体均须落在 [90%, 120%].
	assertRatio := func(name string, actualSum, targetSum float64) {
		if targetSum <= 0 {
			t.Errorf("%s: 目标合计 = %v, 期望大于 0", name, targetSum)
			return
		}
		ratio := actualSum / targetSum
		if ratio < 0.9 || ratio > 1.2 {
			t.Errorf("%s: 完成率 = %.4f, 超出 [0.9, 1.2]", name, ratio)
		}
	}
	for k, tv := range target {
		if k.ym > maxYM {
			continue // 计划月份无实绩, 不参与完成率校验.
		}
		if tv > 0 && actual[k] == 0 {
			t.Errorf("省 %d 月份 %d: 目标 = %v 但实绩为 0", k.pid, k.ym, tv)
			continue
		}
		if tv > 0 {
			assertRatio(fmt.Sprintf("省 %d 月份 %d", k.pid, k.ym), actual[k], tv)
		}
	}
	for pid, av := range provActual {
		assertRatio(fmt.Sprintf("省 %d 汇总", pid), av, provTarget[pid])
	}
	assertRatio("整体汇总", totalActual, totalTarget)
}

// TestSaleTargetsSingleYear 回归测试: 事实仅覆盖单一年度时, 目标区间退化为覆盖该年全年,
// T06 不为空且月份不超出该年 12 月.
func TestSaleTargetsSingleYear(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	end := time.Date(2025, 5, 15, 0, 0, 0, 0, time.UTC)
	dir := t.TempDir()
	cfg := &config.Config{
		OutputDir: dir, Locale: config.LocaleZhCN,
		ProductCount: 8, StoreCount: 9, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -130), EndDate: end,
	}
	if err = cfg.Validate(); err != nil {
		t.Fatalf("validate: %v", err)
	}
	if _, err = New(cfg, ds, nil).GenerateAll(); err != nil {
		t.Fatalf("generate: %v", err)
	}

	recs := readCSVWithBOM(t, filepath.Join(dir, model.FileSaleTarget))
	if len(recs) <= 1 {
		t.Fatalf("T06 无数据行, 单一年度窗口下目标区间应退化为该年全年")
	}
	for i, rec := range recs[1:] {
		d := parseDate(rec[3])
		ym := d.Year()*100 + int(d.Month())
		if d.Year() != 2025 || ym > 202512 {
			t.Errorf("T06 行 %d: 目标月份 %d 超出单一年度 2025 全年", i+1, ym)
		}
	}
}
