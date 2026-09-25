// FilePath    : internal/generator/province.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 省份销售规模五梯队客观规律: 双通道梯队权重 + 梯队内随机梯度 + 梯队交界互换.

package generator

import (
	"sort"

	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// 背景与设计 (实际规律: 省份销售额按五个梯队自高到低分布, 替代早期的大区级排序):
//  1. 双通道分工: 门店配额通道只做温和分层 (相邻梯队配额之比 1.10-1.18,
//     默认 55 家门店下每省期望 1.1-1.8 家, 任何省份都不会落空, 也避免 1->2 家的
//     取整摆动直接淹没排序); 单店订单量通道承担梯队间的主要差距,
//     两通道相乘的梯队销售额份额约为 50% / 22% / 16% / 9% / 3%;
//     第 1/2 交界的单省销售额之比仅约 1.4 倍 (8 省与 5 省均摊相近份额),
//     该交界在产物中存在部分混排, 属 "份额与排序折中" 的既定取舍;
//  2. 梯队内随机: 每次生成对各梯队内部洗牌, 名次自上而下叠加小幅递减梯度,
//     梯队内的相对次序随种子变化 (随机序列不可复现是预期行为);
//  3. 梯队交界互换: 每个交界处, 下梯队的第一名上探一个梯队 (权重落在上梯队带内的
//     下半段, 即仅超越上梯队的尾部成员), 上梯队的末位滑落一个梯队 (权重位于下梯队
//     带顶, 期望名次保持下梯队前 3); 两侧同时发生, 各梯队块的大小保持不变;
//  4. 量级不变: 配额权重省数加权均值恒为 1, 订单量端再按 ΣQ/Σ(QxV) 归一,
//     使按门店数加权的平均订单量系数为 1, 整体数据量级与不注入规律时一致;
//  5. 增量一致: 省份权重的任何取值都只落在 本梯队带 或 相邻梯队带 内, 增量更新时
//     重新洗牌后每个省份的期望销售额仍不会跨越超过一个梯队, 客观规律整体不被破坏.

// provinceTiers 省份销售规模五梯队 (从高到低), 按省份简称 (D01 的 ProvinceShort2) 命名.
var provinceTiers = [5][]string{
	{"广东", "北京", "上海", "重庆", "天津", "四川", "浙江", "江苏"},
	{"福建", "山东", "安徽", "台湾", "河北"},
	{"贵州", "陕西", "山西", "河南", "云南", "湖北", "湖南"},
	{"新疆", "海南", "黑龙江", "辽宁", "内蒙古", "广西", "江西", "香港"},
	{"吉林", "甘肃", "西藏", "宁夏", "青海", "澳门"},
}

// provinceQuotaRatio 相邻梯队的门店配额基准之比 (第 i 梯队 / 第 i+1 梯队).
// 配额通道刻意保持接近 1: 每省至少分得 1 家门店 (地图可视不落空), 且 1->2 家的
// 取整摆动 (销售额 x2) 由订单量通道的陡峭比值吸收.
var provinceQuotaRatio = [4]float64{1.18, 1.15, 1.12, 1.10}

// provinceVolumeRatio 相邻梯队的单店订单量基准之比 (第 i 梯队 / 第 i+1 梯队).
// 与配额比相乘后的梯队销售额份额约为 50% / 22% / 16% / 9% / 3% (相对温和的头部集中).
// 第 1/2 交界的份额比受梯队省数制约 (8 省 vs 5 省), 单省销售额之比仅约 1.4 倍,
// 该交界在产物中存在部分混排, 属 "份额与排序折中" 的既定取舍; 越靠后的交界
// 省内门店数越少 (1-2 家), 门店数取整的 ±2 倍摆动越大, 比值仍需递增.
var provinceVolumeRatio = [4]float64{1.2, 1.67, 1.81, 2.05}

// provinceTierSpread 梯队内权重梯度幅度: 梯队带为 基准权重 x [1-spread/2, 1+spread/2],
// 自梯队内第一名至末名线性递减; 梯度仅供交界互换的成员落位, 幅度远小于门店级噪声,
// 梯队内实际名次仍由随机性主导.
const provinceTierSpread = 0.08

// provinceQuotaBase 梯队门店配额基准, 由 provinceQuotaRatio 自第 5 梯队逐级导出并归一化,
// 满足 Σ(梯队省数 x 基准) = 全部梯队省数 (省数加权均值为 1).
var provinceQuotaBase = func() [5]float64 {
	var q [5]float64
	q[len(q)-1] = 1
	for t := len(q) - 2; t >= 0; t-- {
		q[t] = q[t+1] * provinceQuotaRatio[t]
	}
	return normalizeTierBase(q)
}()

// provinceVolumeBase 梯队单店订单量基准, 由 provinceVolumeRatio 自第 5 梯队逐级导出
// (不预归一, 量级由运行期的 provinceVolumeScale 统一压缩).
var provinceVolumeBase = func() [5]float64 {
	var v [5]float64
	v[len(v)-1] = 1
	for t := len(v) - 2; t >= 0; t-- {
		v[t] = v[t+1] * provinceVolumeRatio[t]
	}
	return v
}()

// normalizeTierBase 将梯队基准归一化为 省数加权均值 1 (不改变整体量级).
//   - base, 待归一的梯队基准.
//
// 返回值 [5]float64, 归一后的梯队基准.
func normalizeTierBase(base [5]float64) [5]float64 {
	var members, weighted float64
	for t, tier := range provinceTiers {
		members += float64(len(tier))
		weighted += float64(len(tier)) * base[t]
	}
	for t := range base {
		base[t] *= members / weighted
	}
	return base
}

// ensureProvinceWeights 懒初始化当次生成的省份销售权重体系:
// 各梯队内部 Fisher-Yates 洗牌后叠加递减梯度 (配额与订单量两通道共用同一梯度),
// 再在每个交界处执行 "下梯队头名上探 + 上梯队末位滑落" 的互换;
// 全量与增量两条路径首次取用权重时各自构建一次, 同一次生成内取值一致.
func (g *Generator) ensureProvinceWeights() {
	if g.provinceQuotaWeight != nil {
		return
	}
	blocks := g.provinceTierBlocks()
	quota := make(map[int]float64, 34)
	volume := make(map[int]float64, 34)
	applyTierGradient(blocks, quota, volume)
	g.provinceBlocks = applyBoundaryExchange(blocks, quota, volume, g.rnd)

	// 配额精确归一 (省数加权均值 1); 订单量按 Σ(q/cf) / Σ(q/cf x V) 归一:
	// 门店配额与 q/cf 成正比 (城市系数补偿), 以其为基数归一可使
	// 按门店数加权的平均订单量系数恰为 1, 且各梯队销售额份额贴近期望 (整体量级不变).
	var sumQ, sumEff, sumEffV float64
	for pid, q := range quota {
		sumQ += q
		eff := q / g.provinceCityFactor[pid]
		sumEff += eff
		sumEffV += eff * volume[pid]
	}
	scale := float64(len(quota)) / sumQ
	for pid := range quota {
		quota[pid] *= scale
	}
	g.provinceQuotaWeight = quota
	g.provinceVolumeWeight = volume
	g.provinceVolumeScale = sumEff / sumEffV
	g.provinceSalesRank = rankProvinces(quota, volume)
}

// provinceTierBlocks 将五梯队定义解析为省 ID 块并各自 Fisher-Yates 洗牌
// (梯队内随机排序; 省份表顺序即 CSV 行序, 解析结果确定, 仅洗牌随种子变化).
//
// 返回值 [][]int, 五个梯队块 (省 ID 列表, 原地洗牌).
func (g *Generator) provinceTierBlocks() [][]int {
	nameTier := make(map[string]int, 34)
	for t, tier := range provinceTiers {
		for _, name := range tier {
			nameTier[name] = t
		}
	}
	blocks := make([][]int, len(provinceTiers))
	for _, p := range g.ds.Provinces {
		if t, ok := nameTier[p.Short2]; ok {
			blocks[t] = append(blocks[t], p.ProvinceID)
		}
	}
	for _, block := range blocks {
		shuffleProvinceIDs(block, g.rnd)
	}
	return blocks
}

// applyTierGradient 按梯队内洗牌名次为各省份叠加递减梯度:
// 第一名 1+spread/2, 末名 1-spread/2, 配额与订单量两通道共用同一梯度系数.
//   - blocks, 梯队块 (省 ID 列表); quota/volume, 输出的两通道权重表 (原地写入).
func applyTierGradient(blocks [][]int, quota, volume map[int]float64) {
	for t, block := range blocks {
		for j, pid := range block {
			f := tierBandFactor(len(block), j)
			quota[pid] = provinceQuotaBase[t] * f
			volume[pid] = provinceVolumeBase[t] * f
		}
	}
}

// applyBoundaryExchange 在每个梯队交界执行 "下梯队头名上探 + 上梯队末位滑落" 的互换:
// 上探者两通道权重都落在上梯队带下半段 (仅超越上梯队尾部成员), 滑落者两通道权重都
// 位于下梯队带顶之上 (期望名次保持下梯队前 3, 通常为第 1); 落位后每个名次块的成员数
// 仍等于梯队省数 (滑落者让出的位置恰好由上探者补上).
//   - blocks, 梯队块; quota/volume, 两通道权重表 (原地改写互换成员的取值).
//   - rnd, 随机源 (上探/滑落在带内的精确落位随种子变化).
//
// 返回值 [][]int, 互换后的五个名次块.
func applyBoundaryExchange(blocks [][]int, quota, volume map[int]float64, rnd *util.Rand) [][]int {
	ranked := make([][]int, len(blocks))
	for t := range blocks {
		ranked[t] = append([]int(nil), blocks[t]...)
	}
	for t := 0; t+1 < len(blocks); t++ {
		upper, lower := blocks[t], blocks[t+1]
		if len(upper) == 0 || len(lower) < 2 {
			continue
		}
		promoted, demoted := lower[0], upper[len(upper)-1]
		fUp := 1 - provinceTierSpread*0.5*rnd.F()
		fDown := 1 + provinceTierSpread*(1+0.5*rnd.F())
		quota[promoted] = provinceQuotaBase[t] * fUp
		volume[promoted] = provinceVolumeBase[t] * fUp
		quota[demoted] = provinceQuotaBase[t+1] * fDown
		volume[demoted] = provinceVolumeBase[t+1] * fDown
		ranked[t][len(ranked[t])-1] = promoted
		ranked[t+1][0] = demoted
	}
	return ranked
}

// rankProvinces 返回按期望销售额 (配额 x 订单量) 降序的省份名次表,
// 同权重按省 ID 升序保证确定.
//   - quota/volume, 两通道权重表.
//
// 返回值 map[int]int, 省 ID -> 名次 (0 最靠前).
func rankProvinces(quota, volume map[int]float64) map[int]int {
	order := make([]int, 0, len(quota))
	for pid := range quota {
		order = append(order, pid)
	}
	sort.Slice(order, func(a, b int) bool {
		wa, wb := quota[order[a]]*volume[order[a]], quota[order[b]]*volume[order[b]]
		if wa != wb {
			return wa > wb
		}
		return order[a] < order[b]
	})
	rank := make(map[int]int, len(order))
	for r, pid := range order {
		rank[pid] = r
	}
	return rank
}

// tierBandFactor 返回梯队内第 j 名 (0-based) 的梯队带系数: 第一名 1+spread/2,
// 末名 1-spread/2, 线性递减; 单成员梯队恒为 1.
//   - n, 梯队省数; j, 梯队内名次 (0-based).
//
// 返回值 float64, 梯队带系数.
func tierBandFactor(n, j int) float64 {
	if n <= 1 {
		return 1
	}
	return 1 + provinceTierSpread*(0.5-float64(j)/float64(n-1))
}

// shuffleProvinceIDs 对省 ID 列表做 Fisher-Yates 洗牌 (梯队内随机排序).
//   - ids, 待洗牌的省 ID 列片 (原地修改).
//   - rnd, 随机源.
func shuffleProvinceIDs(ids []int, rnd *util.Rand) {
	for i := len(ids) - 1; i > 0; i-- {
		j := int(float64(i+1) * rnd.F())
		ids[i], ids[j] = ids[j], ids[i]
	}
}

// provinceBaseFactor 返回门店所属省份的订单量端权重 (单店订单量基准 x 梯度 x 量级
// 归一系数 x 门店数取整补偿, 未知省份为 1), 与门店配额通道同向叠加, 保证省份梯队
// 排序与销售额份额均稳健, 且整体量级不变.
//   - s, 门店.
//
// 返回值 float64, 订单量端权重.
func (g *Generator) provinceBaseFactor(s *store) float64 {
	g.ensureProvinceWeights()
	pid, ok := g.ds.CityToProvince[s.cityID]
	if !ok {
		return 1
	}
	v, ok := g.provinceVolumeWeight[pid]
	if !ok {
		return 1
	}
	f := v * g.provinceVolumeScale
	if comp, ok := g.provinceStoreComp[pid]; ok {
		f *= comp
	}
	return f
}
