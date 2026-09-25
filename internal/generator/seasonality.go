// FilePath    : internal/generator/seasonality.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 维度 x 时间交互的季节性随机因子, 打破各分组趋势形状几乎一致的问题.

package generator

import "math"

// 背景: 订单量与金额的时间因子 (monthTrend/monthNoise/yearTrend/discountMonth) 对所有
// 产品分类/客户行业/大区完全相同, 而实体因子 (productPopularity/regionOrderFactor/门店客流系数)
// 又与时间无关, 期望值构成可分离模型 E[金额(分组, 月份)] = C(分组) x f(月份),
// 表现为任意分组的时间序列形状几乎一致 (仅相差一个常数倍), 毛利率亦只随全局折扣月份系数变化.
// 本文件在既有全局因子之上叠加 "维度 x 时间" 的交互因子:
//  1. 产品分类/客户行业/大区各自拥有相位错开的月度季节谐波, 分别作用于销量与订单量;
//  2. 各维度再叠加确定性年度漂移, 使按年 (NF) 聚合的形态也分化;
//  3. 折扣比例在分类或大区的销售旺季反向让利, 使毛利率在 月份x大区 与 分类x大区 上产生差异.
//
// 数学性质: 谐波的年均值严格为 1 (余弦整周期求和为零), 年度漂移均值为 1, 整体量级保持不变;
// 全部因子按 (维度索引, 年, 月) 确定性导出, 同一实体日期在全量与增量两次生成中取值一致.

// 年度漂移的维度命名空间偏移, 避免不同维度类型共享同一漂移序列.
const (
	seasonDimCategory = 0
	seasonDimRegion   = 100
	seasonDimIndustry = 200
	seasonDimMarginLv = 300 // 大区毛利率水平
	seasonDimInteract = 400 // 分类 x 大区交互
)

// 各维度的年度漂移幅度: 分类 (产品生命周期兴衰) 最强, 行业 (景气周期) 次之, 大区 (宏观趋同) 最弱.
// 漂移均值为 1, 不改变整体量级; 时间序列被门店陆续开业的增长斜坡主导时,
// 只有大幅度的年度漂移才能让各分组的份额随年份明显演化, 打破 "趋势几乎一致".
// 分类的漂移幅度额外补偿销量公式 Round(5F x f)+1 的 +1 与取整带来的约三成幅度衰减.
const (
	seasonDriftCategory = 0.40
	seasonDriftRegion   = 0.10
	seasonDriftIndustry = 0.20
)

// discountSeasonKappa 分类让利强度: 折扣比例随分类季节谐波 (滞后相位) 反向浮动的比例.
var discountSeasonKappa = 0.4

// marginRegionKappa 大区让利强度: 折扣比例随大区毛利率季节谐波反向浮动的比例,
// 声明为变量供测试临时调整以验证接线.
var marginRegionKappa = 0.55

// discountSeasonLag 让利滞后月数: 折扣让利高峰滞后于销量高峰一个季度
// (旺季冲量在先, 季末清仓让利在后); 恰为 3 个月时让利波与销量波正交 (相位差 90°),
// 毛利率获得差异化的一阶效应, 而金额端的季节差异不会被让利反向抵消.
const discountSeasonLag = 3

// marginLevelDrift 大区毛利率水平的年度漂移幅度 (作用于折扣):
// 各大区的毛利率基线逐年上下漂移, 折线图上各大区曲线明显分层且层间排序逐年变化.
const marginLevelDrift = 0.06

// interactDrift 分类 x 大区交互幅度 (作用于折扣, 按年确定性散列):
// 打破 毛利率(分类, 大区) = f(分类) x g(大区) 的可分离性, 使分类 x 大区矩阵逐年不可比例缩放.
const interactDrift = 0.05

// categorySeasonParam 产品分类季节性参数 {幅度, 旺月}, 按分类 A-J 索引.
// 价格档位越低季节性越强 (入门款受大促驱动, 奢侈款全年平稳), 相位按主力销售窗口错开,
// 且上半年主力 (C/D/F/H/I) 与下半年主力 (A/B/E/G/J) 两 camp 重心交错, 季度/半年聚合后形态互异.
var categorySeasonParam = [10][2]float64{
	{0.24, 9.5},  // A 入门款: 双 11 前后冲量 (下半年)
	{0.22, 10.5}, // B: 国庆至双 11 窗口 (下半年)
	{0.20, 1},    // C: 年末乔迁婚庆跨年下单 (上半年)
	{0.19, 3},    // D: 春季家装开工 (上半年)
	{0.18, 7},    // E: 618 之后的暑期换新 (下半年)
	{0.17, 5},    // F: 五一婚房 (上半年)
	{0.15, 8.5},  // G: 夏末秋初开学季 (下半年)
	{0.13, 4},    // H: 春季 (上半年)
	{0.10, 2.5},  // I: 年初平稳 (上半年)
	{0.07, 11},   // J 奢侈款: 高端年末 (下半年)
}

// regionSeasonParam 大区季节性参数 {幅度, 旺月}, 按大区 ID-1 索引 (东区/西区/南区/北区/中区/港澳台).
// 幅度与全局淡旺季 monthTrend 同量级 (0.45-0.60), 各大区用自己的淡旺季形态主导月度走势;
// 旺月近似等距铺满全年 (相邻间隔 >= 2 个月), 任意两区的月度折线不保持同形态.
var regionSeasonParam = [6][2]float64{
	{0.55, 9},   // 东区: 电商促销季
	{0.50, 4},   // 西区: 五一前春装
	{0.45, 11},  // 南区: 年末大促
	{0.60, 7.5}, // 北区: 冬季深淡, 盛夏旺
	{0.50, 2.5}, // 中区: 春节后返工开工
	{0.55, 1},   // 港澳台: 春节旅游
}

// industrySeasonParam 客户行业季节性参数 {幅度, 旺月}, 按 industries 数组顺序索引.
// 幅度 0.45-0.70: 因子同时作用于每单种类数与单件数量 (同源相乘), 有效幅度约为标注幅度的两倍;
// 旺月近似等距铺满全年 (相邻间隔 >= 1.6 个月), 任意两个行业的月度折线不保持同形态.
var industrySeasonParam = [7][2]float64{
	{0.60, 4.4},  // 建筑业: 金四银五开工旺季
	{0.45, 7.9},  // 制造业: 夏末赶产供年末交付
	{0.55, 6.1},  // 互联网: 618 大促
	{0.50, 9.6},  // 农业: 秋收
	{0.70, 1},    // 餐饮: 春节聚餐
	{0.60, 11.3}, // 物流: 双 11 之后至年末
	{0.50, 2.7},  // 汽车: 春季购车
}

// marginRegionAmp 大区毛利率季节幅度, 按大区 ID-1 索引, 作用于折扣端.
// 独立于销量端的 regionSeasonParam 幅度: 毛利率对折扣变化更敏感, 幅度过大会使
// 部分月份毛利率整体转负, 0.15-0.20 配合 marginRegionKappa 得到约 ±7 个百分点的合理摆动.
// (大区级的销售规模排序权重已升级为省级五梯队, 见 province.go.)
var marginRegionAmp = [6]float64{0.18, 0.16, 0.15, 0.20, 0.16, 0.18}

// marginRegionNorm 大区让利系数的销量加权归一因子: 使 12 个月的销量加权平均恰为 1.
// 未归一化时, "旺季后让利" 的让利月若与全局高销量月重叠, 该大区全年毛利率会被
// 系统性拉低甚至转负; 归一化只保留月度形态, 消除年度直流偏移.
var marginRegionNorm = func() [6]float64 {
	var norm [6]float64
	for r, p := range regionSeasonParam {
		var weighted, volume float64
		for m := 1; m <= 12; m++ {
			v := monthTrend[m-1] * seasonHarmonic(p[0], p[1], m) // 该大区 m 月的相对销量
			f := 1 - marginRegionKappa*(seasonHarmonic(marginRegionAmp[r], p[1]+discountSeasonLag, m)-1)
			weighted += v * f
			volume += v
		}
		norm[r] = weighted / volume
	}
	return norm
}()

// industryIndexMap 客户行业 -> 行业索引, 与 industries 数组顺序一致.
var industryIndexMap = func() map[string]int {
	m := make(map[string]int, len(industries))
	for i, name := range industries {
		m[name] = i
	}
	return m
}()

// seasonHarmonic 返回单维度月度季节谐波系数: 1 + amp x cos(2π(month-peak)/12).
// 系数在旺月 peak 处最大 (1+amp), 相位错开的维度因此拥有互不相同的淡旺季形态;
// 余弦在 12 个月上恰好取满一个整周期, 年均值严格为 1, 不改变整体量级.
//   - amp, 波动幅度.
//   - peak, 旺月位置 (可为小数, 如 9.5 表示 9-10 月之间).
//   - month, 月份 (1-12).
//
// 返回值 float64, 季节谐波系数.
func seasonHarmonic(amp, peak float64, month int) float64 {
	return 1 + amp*math.Cos(2*math.Pi*(float64(month)-peak)/12)
}

// seasonYearDrift 从 (维度索引, 年份) 确定性导出年度漂移系数 [1-amp, 1+amp] (整型乘法散列),
// 使同一维度在不同年份的整体水平存在高低分化, 按年/半年聚合的序列不再形状一致;
// 漂移均值为 1, 不改变整体量级; 确定性保证全量与增量两次生成取值一致.
//   - dim, 维度索引 (已叠加命名空间偏移).
//   - year, 年份.
//   - amp, 漂移幅度 (如 0.15 表示 [0.85, 1.15]).
//
// 返回值 float64, 年度漂移系数.
func seasonYearDrift(dim, year int, amp float64) float64 {
	x := uint64(dim)*0x9E3779B97F4A7C15 + uint64(year)*2654435761 // #nosec G115 小正整数转 uint64, 无溢出风险
	x ^= x >> 13
	x *= 1274126177
	x ^= x >> 16
	return 1 - amp + 2*amp*float64(x%1000003)/1000003.0
}

// dimSeason 返回维度在 (年, 月) 的综合季节系数: 月度谐波 x 年度漂移.
//   - amp, peak, 谐波幅度与旺月.
//   - dim, 维度索引 (已叠加命名空间偏移).
//   - year, month, 年与月份.
//   - drift, 年度漂移幅度.
//
// 返回值 float64, 综合季节系数.
func dimSeason(amp, peak float64, dim, year, month int, drift float64) float64 {
	return seasonHarmonic(amp, peak, month) * seasonYearDrift(dim, year, drift)
}

// categoryIndex 返回产品分类 (如 "A类") 的季节性索引, A=0 至 J=9; 无法识别时返回 -1.
//   - category, 产品分类名称.
//
// 返回值 int, 分类索引, 未知为 -1.
func categoryIndex(category string) int {
	if category == "" || category[0] < 'A' || category[0] > 'J' {
		return -1
	}
	return int(category[0] - 'A')
}

// regionIndex 返回门店城市所属大区的季节性索引 (0-5, 对应东区/西区/南区/北区/中区/港澳台);
// 城市或大区未知时返回 -1.
//   - cityID, 门店城市 ID.
//
// 返回值 int, 大区索引, 未知为 -1.
func (g *Generator) regionIndex(cityID int) int {
	pid, ok := g.ds.CityToProvince[cityID]
	if !ok {
		return -1
	}
	idx := g.ds.ProvinceByID[pid].RegionID - 1
	if idx < 0 || idx >= len(regionSeasonParam) {
		return -1
	}
	return idx
}

// storeFactorIdx 从门店编号确定性导出 [0, n) 的哈希索引 (整型乘法散列),
// 供门店级离散因子 (折扣策略类别/客流系数/单均件数系数) 使用:
// 大区交错序列对门店编号的模式是确定性的, 直接按编号取模会使编号段与因子段固定耦合,
// 造成个别大区系统性偏高/偏低; 哈希索引打散该耦合, 且全量与增量两次生成取值一致.
//   - id, 门店编号; n, 索引上界.
//
// 返回值 int, [0, n) 的哈希索引.
func storeFactorIdx(id, n int) int {
	x := uint64(id)*0x9E3779B97F4A7C15 + 0xABC98372F // #nosec G115 小正整数转 uint64, 无溢出风险
	x ^= x >> 13
	x *= 1274126177
	x ^= x >> 16
	return int(x % uint64(n)) // #nosec G115 哈希取模结果小于 n, 无溢出风险
}

// regionOrderSeasonFactor 返回门店所属大区在 (年, 月) 的销量端季节系数 (订单量与每单种类数共用),
// 未知大区为 1.
//   - s, 门店.
//   - year, month, 年与月份.
//
// 返回值 float64, 季节系数.
func (g *Generator) regionOrderSeasonFactor(s *store, year, month int) float64 {
	ri := g.regionIndex(s.cityID)
	if ri < 0 {
		return 1
	}
	p := regionSeasonParam[ri]
	return dimSeason(p[0], p[1], seasonDimRegion+ri, year, month, seasonDriftRegion)
}

// categorySeasonFactor 返回产品分类在 (年, 月) 的销量季节系数, 未知分类为 1.
//   - category, 产品分类.
//   - year, month, 年与月份.
//
// 返回值 float64, 季节系数.
func categorySeasonFactor(category string, year, month int) float64 {
	i := categoryIndex(category)
	if i < 0 {
		return 1
	}
	p := categorySeasonParam[i]
	return dimSeason(p[0], p[1], seasonDimCategory+i, year, month, seasonDriftCategory)
}

// industrySeasonFactor 返回订单客户所属行业在 (年, 月) 的购买量季节系数, 无客户或未知行业为 1.
//   - customerIdx, 客户索引 (可为 -1 表示无客户).
//   - year, month, 年与月份.
//
// 返回值 float64, 季节系数.
func (g *Generator) industrySeasonFactor(customerIdx, year, month int) float64 {
	if customerIdx < 0 || customerIdx >= len(g.customers) {
		return 1
	}
	i, ok := industryIndexMap[g.customers[customerIdx].industry]
	if !ok {
		return 1
	}
	p := industrySeasonParam[i]
	return dimSeason(p[0], p[1], seasonDimIndustry+i, year, month, seasonDriftIndustry)
}

// discountSeasonFactor 返回 (大区, 产品分类, 年月) 的折扣修正系数, 围绕 1 浮动,
// 由调用处保证折扣比例不超过 1, 全部为确定性索引:
//  1. 大区毛利率水平: 各大区折扣水平逐年漂移 (毛利率基线分层, 层间排序逐年变化);
//  2. 大区毛利率季节: 旺季后一个季度让利 (相位与销量波正交), 毛利率随 月份x大区 波动;
//  3. 分类让利季节: 分类的旺季后让利, 毛利率随分类变化;
//  4. 分类 x 大区交互: 按 (分类, 大区, 年) 散列, 打破 毛利率(分类,大区) 的可分离性.
//
// 仅使用月度谐波 (不含年度漂移): 漂移为年份常量不随相位移动, 若参与让利会反向抵消金额端分化.
//   - s, 门店 (用于定位大区).
//   - category, 产品分类.
//   - year, month, 年与月份.
//
// 返回值 float64, 折扣修正系数.
func (g *Generator) discountSeasonFactor(s *store, category string, year, month int) float64 {
	ri := g.regionIndex(s.cityID)
	f := 1.0
	if ri >= 0 {
		f *= seasonYearDrift(seasonDimMarginLv+ri, year, marginLevelDrift)
		f *= (1 - marginRegionKappa*(seasonHarmonic(marginRegionAmp[ri], regionSeasonParam[ri][1]+discountSeasonLag, month)-1)) /
			marginRegionNorm[ri]
	}
	if ci := categoryIndex(category); ci >= 0 {
		f *= 1 - discountSeasonKappa*(seasonHarmonic(categorySeasonParam[ci][0], categorySeasonParam[ci][1]+discountSeasonLag, month)-1)
		if ri >= 0 {
			f *= seasonYearDrift(seasonDimInteract+ci*6+ri, year, interactDrift)
		}
	}
	return f
}
