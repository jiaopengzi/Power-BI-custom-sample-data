// FilePath    : internal/generator/seasonality_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 季节性交互因子的数学性质, 维度索引解析与生成链路接线回归测试.

package generator

import (
	"math"
	"path/filepath"
	"testing"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// TestSeasonFactorProperties 校验季节性因子的基本数学性质:
// 谐波年均值严格为 1 且恒为正; 年度漂移落在给定幅度内; 各表幅度与漂移的最坏叠加有界;
// 同一维度的任意两档参数在 12 个月中存在明显差异 (淡旺季形态互不相同).
func TestSeasonFactorProperties(t *testing.T) {
	drifts := map[string]float64{
		"产品分类": seasonDriftCategory,
		"大区":   seasonDriftRegion,
		"客户行业": seasonDriftIndustry,
	}
	tables := map[string][][2]float64{
		"产品分类": categorySeasonParam[:],
		"大区":   regionSeasonParam[:],
		"客户行业": industrySeasonParam[:],
	}
	for name, params := range tables {
		drift := drifts[name]
		for i, p := range params {
			sum := 0.0
			for m := 1; m <= 12; m++ {
				f := seasonHarmonic(p[0], p[1], m)
				if f <= 0 {
					t.Fatalf("%s 第 %d 档: 谐波系数 = %v, 须恒为正", name, i, f)
				}
				sum += f
			}
			if math.Abs(sum/12-1) > 1e-9 {
				t.Errorf("%s 第 %d 档: 谐波年均值 = %.12f, 期望严格为 1 (不改变整体量级)", name, i, sum/12)
			}
			if lo, hi := (1-p[0])*(1-drift), (1+p[0])*(1+drift); lo <= 0.2 || hi >= 2.1 {
				t.Errorf("%s 第 %d 档: 幅度 %.2f x 漂移 %.2f 最坏叠加 [%.2f, %.2f] 超出 (0.2, 2.1)", name, i, p[0], drift, lo, hi)
			}
			for j := i + 1; j < len(params); j++ {
				maxDiff := 0.0
				for m := 1; m <= 12; m++ {
					d := math.Abs(seasonHarmonic(p[0], p[1], m) - seasonHarmonic(params[j][0], params[j][1], m))
					maxDiff = max(maxDiff, d)
				}
				if maxDiff <= 0.02 {
					t.Errorf("%s 第 %d/%d 档: 谐波最大月度差异 = %.4f, 淡旺季形态不应几乎相同", name, i, j, maxDiff)
				}
			}
		}
	}
	for i, amp := range marginRegionAmp {
		if amp <= 0 || amp >= 0.3 {
			t.Errorf("大区 %d 毛利率季节幅度 = %v, 超出 (0, 0.3)", i, amp)
		}
	}

	for dim := range 300 {
		for year := 2020; year <= 2030; year++ {
			drift := seasonYearDrift(dim, year, 0.05)
			if drift < 0.95 || drift > 1.05 {
				t.Fatalf("维度 %d 年份 %d: 年度漂移 = %v, 超出 [0.95, 1.05]", dim, year, drift)
			}
		}
	}
}

// TestSeasonIndexHelpers 校验维度索引解析: 产品分类 (A类-J类), 城市所属大区与客户行业,
// 均与 T00/T01/T02 产物中的字段格式兼容 (全量生成与增量载入共用同一解析).
func TestSeasonIndexHelpers(t *testing.T) {
	if got := categoryIndex("A类"); got != 0 {
		t.Errorf(`categoryIndex("A类") = %d, want 0`, got)
	}
	if got := categoryIndex("J类"); got != 9 {
		t.Errorf(`categoryIndex("J类") = %d, want 9`, got)
	}
	for _, c := range []string{"", "K类", "产品", "a类"} {
		if got := categoryIndex(c); got != -1 {
			t.Errorf("categoryIndex(%q) = %d, want -1", c, got)
		}
	}

	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	g := New(&config.Config{OutputDir: t.TempDir()}, ds, nil)
	// 北京 110000 属中区 (大区 ID 5), 广州 440100 属南区 (大区 ID 3).
	if got := g.regionIndex(110000); got != 4 {
		t.Errorf("regionIndex(110000) = %d, want 4 (中区)", got)
	}
	if got := g.regionIndex(440100); got != 2 {
		t.Errorf("regionIndex(440100) = %d, want 2 (南区)", got)
	}
	if got := g.regionIndex(999999); got != -1 {
		t.Errorf("regionIndex(999999) = %d, want -1", got)
	}

	for i, name := range industries {
		if got := industryIndexMap[name]; got != i {
			t.Errorf("industryIndexMap[%q] = %d, want %d", name, got, i)
		}
	}
	if _, ok := industryIndexMap["金融"]; ok {
		t.Error(`industryIndexMap["金融"] 不应存在`)
	}
}

// dimKind 月度聚合的维度类型.
type dimKind int

const (
	dimRegion dimKind = iota
	dimCategory
	dimIndustry
)

// fixtureDims 载入产物目录中的维度索引与订单关联信息, 供月度聚合断言使用.
type fixtureDims struct {
	storeRegion      map[string]int     // 门店编号 -> 大区索引
	customerIndustry map[string]int     // 客户编号 -> 行业索引
	productCategory  map[string]int     // 产品编号 -> 分类索引
	productCost      map[string]float64 // 产品编号 -> 成本价
	orderRegion      map[string]int     // 订单编号 -> 大区索引
	orderIndustry    map[string]int     // 订单编号 -> 行业索引
	orderYM          map[string]int     // 订单编号 -> YYYYMM
}

// loadFixtureDims 解析产物目录中的 T00/T01/T02/T04, 建立维度关联索引.
//   - t, 测试对象.
//   - dir, 产物目录.
//   - ds, 基础数据集 (城市->省->大区映射).
//
// 返回值 *fixtureDims, 维度关联索引.
func loadFixtureDims(t *testing.T, dir string, ds *data.Dataset) *fixtureDims {
	t.Helper()
	g := &Generator{ds: ds}
	f := &fixtureDims{
		storeRegion:      make(map[string]int),
		customerIndustry: make(map[string]int),
		productCategory:  make(map[string]int),
		productCost:      make(map[string]float64),
		orderRegion:      make(map[string]int),
		orderIndustry:    make(map[string]int),
		orderYM:          make(map[string]int),
	}
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileStore))[1:] {
		f.storeRegion[rec[1]] = g.regionIndex(atoiSafe(rec[5]))
	}
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileProduct))[1:] {
		f.productCategory[rec[1]] = categoryIndex(rec[2])
		f.productCost[rec[1]] = atofSafe(rec[5])
	}
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileCustomer))[1:] {
		idx, ok := industryIndexMap[rec[6]]
		if !ok {
			idx = -1
		}
		f.customerIndustry[rec[1]] = idx
	}
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrder))[1:] {
		d := parseDate(rec[3])
		f.orderRegion[rec[1]] = f.storeRegion[rec[2]]
		f.orderIndustry[rec[1]] = f.customerIndustry[rec[5]]
		f.orderYM[rec[1]] = d.Year()*100 + int(d.Month())
	}
	return f
}

// monthlyRevenue 返回按维度分组的月度销售金额: 维度值 -> YYYYMM -> 金额.
//   - t, 测试对象.
//   - dir, 产物目录.
//   - kind, 聚合维度.
//
// 返回值 map[int]map[int]float64, 月度销售金额.
func (f *fixtureDims) monthlyRevenue(t *testing.T, dir string, kind dimKind) map[int]map[int]float64 {
	t.Helper()
	out := make(map[int]map[int]float64)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrderItem))[1:] {
		idx := -1
		switch kind {
		case dimRegion:
			idx = f.orderRegion[rec[1]]
		case dimCategory:
			idx = f.productCategory[rec[2]]
		case dimIndustry:
			idx = f.orderIndustry[rec[1]]
		}
		if idx < 0 {
			continue
		}
		if out[idx] == nil {
			out[idx] = make(map[int]float64)
		}
		out[idx][f.orderYM[rec[1]]] += atofSafe(rec[6])
	}
	return out
}

// monthlyMarginByRegion 返回各大区的月度毛利率 (口径同 PBIP 度量 006_毛利率%):
// (销售金额 - 成本金额) / 销售金额.
//   - t, 测试对象.
//   - dir, 产物目录.
//
// 返回值 map[int]map[int]float64, 大区索引 -> YYYYMM -> 毛利率.
func (f *fixtureDims) monthlyMarginByRegion(t *testing.T, dir string) map[int]map[int]float64 {
	t.Helper()
	rev := make(map[int]map[int]float64)
	cost := make(map[int]map[int]float64)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrderItem))[1:] {
		r := f.orderRegion[rec[1]]
		if r < 0 {
			continue
		}
		ym := f.orderYM[rec[1]]
		if rev[r] == nil {
			rev[r] = make(map[int]float64)
			cost[r] = make(map[int]float64)
		}
		rev[r][ym] += atofSafe(rec[6])
		cost[r][ym] += f.productCost[rec[2]] * atofSafe(rec[5])
	}
	margin := make(map[int]map[int]float64)
	for r, revs := range rev {
		for ym, v := range revs {
			if v > 0 {
				if margin[r] == nil {
					margin[r] = make(map[int]float64)
				}
				margin[r][ym] = (v - cost[r][ym]) / v
			}
		}
	}
	return margin
}

// genWithSeed 以固定种子生成一套小规模数据, 供参数突变前后对比 (两次仅季节参数不同时产物可比).
//   - t, 测试对象.
//   - ds, 基础数据集.
//   - seed, 随机种子.
//   - base, 生成配置 (OutputDir 被替换为独立临时目录).
//
// 返回值 string, 产物目录; *Generator, 承载当次生成状态的生成器 (含省份权重体系).
func genWithSeed(t *testing.T, ds *data.Dataset, seed int64, base *config.Config) (string, *Generator) {
	t.Helper()
	cfg := *base
	cfg.OutputDir = t.TempDir()
	g := New(&cfg, ds, nil)
	g.rnd = util.NewRandSeed(seed)
	if _, err := g.GenerateAll(); err != nil {
		t.Fatalf("generate: %v", err)
	}
	return cfg.OutputDir, g
}

// assertProfileShift 断言同一随机种子下季节参数调整前后, 目标维度的月度形态发生实质变化:
// 逐月比值的标准差须超过阈值 (交互因子未接入生成链路时两次产物完全一致, 比值恒为 1).
//   - t, 测试对象.
//   - name, 断言对象描述.
//   - a, b, 调整前后的月度序列 (YYYYMM -> 数值).
//   - threshold, 比值标准差下限.
func assertProfileShift(t *testing.T, name string, a, b map[int]float64, threshold float64) {
	t.Helper()
	if len(a) == 0 {
		t.Fatalf("%s: 月度序列为空, 样本未覆盖该维度", name)
	}
	var ratios []float64
	for ym, va := range a {
		vb, ok := b[ym]
		if !ok || va <= 0 || vb <= 0 {
			continue
		}
		ratios = append(ratios, va/vb)
	}
	if len(ratios) < 6 {
		t.Fatalf("%s: 可比月份 = %d, 期望至少 6 个月", name, len(ratios))
	}
	var mean float64
	for _, r := range ratios {
		mean += r
	}
	mean /= float64(len(ratios))
	var variance float64
	for _, r := range ratios {
		variance += (r - mean) * (r - mean)
	}
	if std := math.Sqrt(variance / float64(len(ratios))); std < threshold {
		t.Errorf("%s: 参数调整前后月度比值标准差 = %.4f, 期望大于 %.2f (交互因子应作用于生成链路)", name, std, threshold)
	}
}

// TestSeasonalityWired 回归测试: 交互因子必须接入生成链路.
// 以同一随机种子生成两次, 仅将目标维度的季节幅度调大; 接入时该维度的月度金额/毛利率
// 形态发生实质变化, 未接入时两次产物完全一致. 目标维度取样本量最大的一档, 避免小样本随机缺失.
func TestSeasonalityWired(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	end := time.Date(2026, 8, 31, 0, 0, 0, 0, time.UTC)
	base := config.Config{
		Locale:       config.LocaleZhCN,
		ProductCount: 60, StoreCount: 16, InventoryCycle: 14,
		StartDate: end.AddDate(0, 0, -420), EndDate: end,
	}
	const seed = 20260924

	t.Run("大区季节作用于订单量", func(t *testing.T) {
		// 让利路径与数量路径共用参数表, 置零让利以隔离, 使本子测试只验证订单量链路.
		oldKappa := discountSeasonKappa
		discountSeasonKappa = 0
		defer func() { discountSeasonKappa = oldKappa }()
		baseDir, _ := genWithSeed(t, ds, seed, &base)
		fb := loadFixtureDims(t, baseDir, ds)
		// 取门店数最多的大区, 避免小样本随机缺失.
		regionStores := make(map[int]int)
		for _, r := range fb.storeRegion {
			if r >= 0 {
				regionStores[r]++
			}
		}
		target, count := -1, 0
		for r, n := range regionStores {
			if n > count {
				target, count = r, n
			}
		}
		if target < 0 {
			t.Fatal("样本未覆盖任何大区")
		}
		old := regionSeasonParam[target]
		mutated := old
		mutated[0] += 0.35                 // 幅度大幅调大, 保证两次产物可比差异充分
		mutated[0] = min(mutated[0], 0.95) // 上限 0.95, 保证系数恒为正
		regionSeasonParam[target] = mutated
		defer func() { regionSeasonParam[target] = old }()
		mutDir, _ := genWithSeed(t, ds, seed, &base)
		fm := loadFixtureDims(t, mutDir, ds)
		assertProfileShift(t, "大区月度金额", fb.monthlyRevenue(t, baseDir, dimRegion)[target],
			fm.monthlyRevenue(t, mutDir, dimRegion)[target], 0.03)
	})

	t.Run("分类季节作用于销量", func(t *testing.T) {
		// 让利路径与数量路径共用参数表, 置零让利以隔离, 使本子测试只验证销量链路.
		oldKappa := discountSeasonKappa
		discountSeasonKappa = 0
		defer func() { discountSeasonKappa = oldKappa }()
		baseDir, _ := genWithSeed(t, ds, seed, &base)
		fb := loadFixtureDims(t, baseDir, ds)
		// 取产品数最多的分类.
		categoryProducts := make(map[int]int)
		for _, c := range fb.productCategory {
			if c >= 0 {
				categoryProducts[c]++
			}
		}
		target, count := -1, 0
		for c, n := range categoryProducts {
			if n > count {
				target, count = c, n
			}
		}
		if target < 0 {
			t.Fatal("样本未覆盖任何产品分类")
		}
		old := categorySeasonParam[target]
		mutated := old
		mutated[0] += 0.35                 // 幅度大幅调大, 保证两次产物可比差异充分
		mutated[0] = min(mutated[0], 0.95) // 上限 0.95, 保证系数恒为正
		categorySeasonParam[target] = mutated
		defer func() { categorySeasonParam[target] = old }()
		mutDir, _ := genWithSeed(t, ds, seed, &base)
		fm := loadFixtureDims(t, mutDir, ds)
		assertProfileShift(t, "分类月度金额", fb.monthlyRevenue(t, baseDir, dimCategory)[target],
			fm.monthlyRevenue(t, mutDir, dimCategory)[target], 0.03)
	})

	t.Run("行业季节作用于购买量", func(t *testing.T) {
		baseDir, _ := genWithSeed(t, ds, seed, &base)
		fb := loadFixtureDims(t, baseDir, ds)
		industryCustomers := make(map[int]int)
		for _, i := range fb.customerIndustry {
			if i >= 0 {
				industryCustomers[i]++
			}
		}
		target, count := -1, 0
		for i, n := range industryCustomers {
			if n > count {
				target, count = i, n
			}
		}
		if target < 0 {
			t.Fatal("样本未覆盖任何客户行业")
		}
		old := industrySeasonParam[target]
		mutated := old
		mutated[0] += 0.35                 // 幅度大幅调大, 保证两次产物可比差异充分
		mutated[0] = min(mutated[0], 0.95) // 上限 0.95, 保证系数恒为正
		industrySeasonParam[target] = mutated
		defer func() { industrySeasonParam[target] = old }()
		mutDir, _ := genWithSeed(t, ds, seed, &base)
		fm := loadFixtureDims(t, mutDir, ds)
		assertProfileShift(t, "行业月度金额", fb.monthlyRevenue(t, baseDir, dimIndustry)[target],
			fm.monthlyRevenue(t, mutDir, dimIndustry)[target], 0.03)
	})

	t.Run("旺季让利作用于毛利率", func(t *testing.T) {
		// 大区让利强度归零后再恢复默认各生成一次: 接入时各大区月度毛利率发生实质变化.
		oldKappa := marginRegionKappa
		marginRegionKappa = 0
		baseDir, _ := genWithSeed(t, ds, seed, &base)
		marginRegionKappa = oldKappa
		defer func() { marginRegionKappa = oldKappa }()
		mutDir, _ := genWithSeed(t, ds, seed, &base)
		fb := loadFixtureDims(t, baseDir, ds)
		fm := loadFixtureDims(t, mutDir, ds)
		baseMargin := fb.monthlyMarginByRegion(t, baseDir)
		mutMargin := fm.monthlyMarginByRegion(t, mutDir)
		shift := 0.0
		for r, sa := range baseMargin {
			for ym, va := range sa {
				if vb, ok := mutMargin[r][ym]; ok {
					shift = max(shift, math.Abs(va-vb))
				}
			}
		}
		if shift < 0.005 {
			t.Errorf("让利系数调整前后各大区月度毛利率最大差异 = %.4f, 期望大于 0.005 (让利因子应作用于折扣链路)", shift)
		}
	})
}

// TestProvinceSalesTiers (省级五梯队销售额规律) 见 province_test.go.
