// FilePath    : internal/generator/province_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 省份五梯队销售规模规律的覆盖性, 权重模型数学性质与生成链路回归测试.

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

// tierOfProvince 返回各省 ID 的归属梯队 (0-4), 供按名次块定位上探/滑落成员.
//   - ds, 基础数据集.
//
// 返回值 map[int]int, 省 ID -> 梯队索引.
func tierOfProvince(ds *data.Dataset) map[int]int {
	nameTier := make(map[string]int, 34)
	for t, tier := range provinceTiers {
		for _, name := range tier {
			nameTier[name] = t
		}
	}
	m := make(map[int]int, len(ds.Provinces))
	for _, p := range ds.Provinces {
		if t, ok := nameTier[p.Short2]; ok {
			m[p.ProvinceID] = t
		}
	}
	return m
}

// TestProvinceTierCoverage 校验五梯队定义恰好覆盖全部 34 个省份且互不重复.
func TestProvinceTierCoverage(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	seen := make(map[string]int, 34)
	for ti, tier := range provinceTiers {
		if len(tier) == 0 {
			t.Errorf("第 %d 梯队为空", ti+1)
		}
		for _, name := range tier {
			if prev, ok := seen[name]; ok {
				t.Errorf("省份 %q 同时出现在第 %d 与第 %d 梯队", name, prev+1, ti+1)
			}
			seen[name] = ti
		}
	}
	for _, p := range ds.Provinces {
		if _, ok := seen[p.Short2]; !ok {
			t.Errorf("省份 %q (ID %d) 未归属任何梯队", p.Short2, p.ProvinceID)
		}
	}
	if got := len(seen); got != len(ds.Provinces) {
		t.Errorf("梯队省份数 = %d, 数据集省份数 = %d, 须一一对应", got, len(ds.Provinces))
	}
}

// TestProvinceWeightModel 校验省份权重体系的数学性质 (多次随机洗牌重复验证):
// 配额权重恒正且省数加权均值恰为 1 (量级不变); 期望销售额 (配额 x 订单量) 的名次块间
// 严格分层 (块内最小值高于下一块内最大值); 上探成员两通道权重都落在上梯队带下半段,
// 滑落成员高于下梯队块内其余全部成员 (期望名次保持下梯队第 1); 块大小等于梯队省数;
// 量级归一系数等于 ΣQ/Σ(QxV).
func TestProvinceWeightModel(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	tierOf := tierOfProvince(ds)
	var sumTier float64
	for ti, tier := range provinceTiers {
		sumTier += float64(len(tier)) * provinceQuotaBase[ti]
	}
	if math.Abs(sumTier/float64(len(ds.Provinces))-1) > 1e-9 {
		t.Errorf("梯队配额基准的省数加权均值 = %.9f, 期望恰为 1 (不改变整体量级)", sumTier/float64(len(ds.Provinces)))
	}
	for ti := 0; ti+1 < len(provinceQuotaBase); ti++ {
		if math.Abs(provinceQuotaBase[ti]/provinceQuotaBase[ti+1]-provinceQuotaRatio[ti]) > 1e-9 {
			t.Errorf("第 %d/%d 梯队配额基准之比 = %.6f, 期望为 %.2f",
				ti+1, ti+2, provinceQuotaBase[ti]/provinceQuotaBase[ti+1], provinceQuotaRatio[ti])
		}
		if math.Abs(provinceVolumeBase[ti]/provinceVolumeBase[ti+1]-provinceVolumeRatio[ti]) > 1e-9 {
			t.Errorf("第 %d/%d 梯队订单量基准之比 = %.6f, 期望为 %.2f",
				ti+1, ti+2, provinceVolumeBase[ti]/provinceVolumeBase[ti+1], provinceVolumeRatio[ti])
		}
	}

	// composite 返回省份的期望销售额相对值 (配额 x 订单量).
	composite := func(g *Generator, pid int) float64 {
		return g.provinceQuotaWeight[pid] * g.provinceVolumeWeight[pid]
	}
	for seed := int64(1); seed <= 50; seed++ {
		g := New(&config.Config{OutputDir: t.TempDir()}, ds, nil)
		g.rnd = util.NewRandSeed(seed)
		g.ensureProvinceWeights()

		var sumQ, sumEff, sumEffV float64
		for pid, q := range g.provinceQuotaWeight {
			if q <= 0 || g.provinceVolumeWeight[pid] <= 0 {
				t.Fatalf("种子 %d: 省 %d 权重非正 (配额 %v, 订单量 %v)", seed, pid, q, g.provinceVolumeWeight[pid])
			}
			eff := q / g.provinceCityFactor[pid]
			sumQ += q
			sumEff += eff
			sumEffV += eff * g.provinceVolumeWeight[pid]
		}
		if math.Abs(sumQ/float64(len(g.provinceQuotaWeight))-1) > 1e-9 {
			t.Errorf("种子 %d: 配额权重均值 = %.9f, 期望恰为 1", seed, sumQ/float64(len(g.provinceQuotaWeight)))
		}
		if math.Abs(g.provinceVolumeScale-sumEff/sumEffV) > 1e-12 {
			t.Errorf("种子 %d: 量级归一系数 = %.9f, 期望为 Σ(q/cf)/Σ(q/cfxV) = %.9f", seed, g.provinceVolumeScale, sumEff/sumEffV)
		}

		blockComposite := func(block []int) (lo, hi float64) {
			lo, hi = math.Inf(1), math.Inf(-1)
			for _, pid := range block {
				c := composite(g, pid)
				lo, hi = min(lo, c), max(hi, c)
			}
			return lo, hi
		}
		for b := 0; b < len(g.provinceBlocks); b++ {
			block := g.provinceBlocks[b]
			if len(block) != len(provinceTiers[b]) {
				t.Fatalf("种子 %d: 名次块 %d 大小 = %d, 期望为梯队省数 %d", seed, b, len(block), len(provinceTiers[b]))
			}
			// 块间严格分层: 本块期望销售额最小值高于下一块最大值 (相隔一个梯队更不可能翻转).
			if b+1 < len(g.provinceBlocks) {
				lo, _ := blockComposite(block)
				nextLo, nextHi := blockComposite(g.provinceBlocks[b+1])
				if lo <= nextHi {
					t.Errorf("种子 %d: 名次块 %d 期望销售额最小 %.4f 未高于块 %d 最大 %.4f (块底 %.4f)",
						seed, b, lo, b+1, nextHi, nextLo)
				}
				// 交界互换定位: 上探者 (本块中归属下一梯队的成员) 两通道权重应落在上梯队带下半段;
				// 滑落者 (下一块中归属本梯队的成员) 期望销售额应高于下一块其余全部成员.
				var promoted, demoted int
				for _, pid := range block {
					if tierOf[pid] == b+1 {
						promoted = pid
					}
				}
				for _, pid := range g.provinceBlocks[b+1] {
					if tierOf[pid] == b {
						demoted = pid
					}
				}
				if promoted == 0 || demoted == 0 {
					t.Fatalf("种子 %d: 交界 %d 缺少上探/滑落成员", seed, b)
				}
				qBase, vBase := provinceQuotaBase[b], provinceVolumeBase[b]
				// 上下界各放宽 1%: 交界互换会使全表权重和偏离基准约 ±0.5%,
				// ensureProvinceWeights 末尾的整体归一随之带来同向漂移.
				for _, w := range []struct {
					name string
					got  float64
					base float64
				}{
					{"配额", g.provinceQuotaWeight[promoted], qBase},
					{"订单量", g.provinceVolumeWeight[promoted], vBase},
				} {
					if w.got <= w.base*(1-provinceTierSpread/2)*0.99 || w.got > w.base*1.01 {
						t.Errorf("种子 %d: 上探省 %d %s权重 %.4f 超出上梯队带下半段 (%.4f, %.4f]",
							seed, promoted, w.name, w.got, w.base*(1-provinceTierSpread/2), w.base)
					}
				}
				for _, pid := range g.provinceBlocks[b+1] {
					if pid != demoted && composite(g, pid) >= composite(g, demoted) {
						t.Errorf("种子 %d: 滑落省 %d 期望销售额 %.4f 未高于下块其余省 %d 的 %.4f",
							seed, demoted, composite(g, demoted), pid, composite(g, pid))
					}
				}
			}
		}
	}
}

// provinceSalesFromCSV 从产物 T01/T04/T05 汇总各省销售额 (订单 -> 门店 -> 城市 -> 省).
//   - t, 测试对象.
//   - dir, 产物目录.
//   - ds, 基础数据集.
//
// 返回值 map[int]float64, 省 ID -> 销售额.
func provinceSalesFromCSV(t *testing.T, dir string, ds *data.Dataset) map[int]float64 {
	t.Helper()
	storeProvince := make(map[string]int)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileStore))[1:] {
		pid, ok := ds.CityToProvince[atoiSafe(rec[5])]
		if !ok {
			continue
		}
		storeProvince[rec[1]] = pid
	}
	orderProvince := make(map[string]int)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrder))[1:] {
		orderProvince[rec[1]] = storeProvince[rec[2]]
	}
	sales := make(map[int]float64, 34)
	for _, rec := range readCSVWithBOM(t, filepath.Join(dir, model.FileOrderItem))[1:] {
		if pid, ok := orderProvince[rec[1]]; ok {
			sales[pid] += atofSafe(rec[6])
		}
	}
	return sales
}

// provinceShareBand 各名次块销售额份额的断言区间 {下限, 上限} (50 种子扫描实测带宽外放约 2pp).
// 份额折中目标约为 50% / 22% / 16% / 9% / 3% (相对温和的头部集中); 单次生成的份额
// 受门店级噪声影响存在 ±5pp 左右的摆动, 均值贴近期望.
var provinceShareBand = [5][2]float64{
	{0.42, 0.58},
	{0.16, 0.30},
	{0.10, 0.19},
	{0.06, 0.13},
	{0.015, 0.048},
}

// provinceInversionBudget 各相邻块逆序对的单种子预算 (50 种子实测最大值 [24 6 15 7] 外放).
// 交界互换本身不产生逆序, 预算容纳门店级噪声; 第 1/2 交界为 "份额与排序折中" 的既定取舍
// (8 省与 5 省均摊相近份额, 单省销售额之比仅约 1.4 倍), 存在部分混排属预期行为.
var provinceInversionBudget = [4]int{30, 8, 17, 9}

// TestProvinceSalesOrder 回归测试: 省份销售额按五梯队名次块自高到低分布, 份额贴近期望:
//  1. 默认 55 家门店下每个省份都有销售 (配额通道保底, 地图可视不落空);
//  2. 五个名次块的销售额份额落在 provinceShareBand 区间内 (折中目标 50/22/16/9/3);
//  3. 相邻块逆序对受 provinceInversionBudget 约束, 且第 1/2 交界的多种子平均逆序
//     低于随机混排基线 (块大小 8x5 的一半 = 20), 即头部仍保持部分有序;
//  4. 相隔 >= 2 个梯队中, 块 1/3 与 2/4 严格有序, 块 0/2 允许个位数违例
//     (头部份额折中削弱了 0/2 的间隔, 块 0 成员偶发低于块 2 头部属预期).
func TestProvinceSalesOrder(t *testing.T) {
	ds, err := data.Load()
	if err != nil {
		t.Fatalf("load data: %v", err)
	}
	end := time.Date(2026, 8, 31, 0, 0, 0, 0, time.UTC)
	base := config.Config{
		Locale:         config.LocaleZhCN,
		ProductCount:   120,
		StoreCount:     55,
		InventoryCycle: 14,
		StartDate:      end.AddDate(0, 0, -1120),
		EndDate:        end,
	}
	headInversions := 0
	for seed := int64(1); seed <= 10; seed++ {
		dir, g := genWithSeed(t, ds, seed, &base)
		sales := provinceSalesFromCSV(t, dir, ds)
		positive := func(block []int) []float64 {
			var vals []float64
			for _, pid := range block {
				if v, ok := sales[pid]; ok && v > 0 {
					vals = append(vals, v)
				}
			}
			return vals
		}
		if len(sales) != len(g.provinceQuotaWeight) {
			t.Errorf("种子 %d: 有销售的省份数 = %d, 期望覆盖全部 %d 个省份",
				seed, len(sales), len(g.provinceQuotaWeight))
		}
		var total float64
		for _, v := range sales {
			total += v
		}
		for b, block := range g.provinceBlocks {
			var blockSum float64
			for _, pid := range block {
				blockSum += sales[pid]
			}
			share := blockSum / total
			if share < provinceShareBand[b][0] || share > provinceShareBand[b][1] {
				t.Errorf("种子 %d: 名次块 %d 销售额份额 = %.1f%%, 超出区间 [%.0f%%, %.0f%%]",
					seed, b, share*100, provinceShareBand[b][0]*100, provinceShareBand[b][1]*100)
			}
		}
		for b := 0; b+1 < len(g.provinceBlocks); b++ {
			hi, lo := positive(g.provinceBlocks[b]), positive(g.provinceBlocks[b+1])
			if len(hi) == 0 {
				t.Errorf("种子 %d: 名次块 %d 无正销售额省份", seed, b)
				continue
			}
			inversions := 0
			for _, va := range hi {
				for _, vb := range lo {
					if va < vb {
						inversions++
					}
				}
			}
			if inversions > provinceInversionBudget[b] {
				t.Errorf("种子 %d: 名次块 %d/%d 逆序对 = %d, 超过噪声预算 %d",
					seed, b, b+1, inversions, provinceInversionBudget[b])
			}
			if b == 0 {
				headInversions += inversions
			}
		}
		// 相隔一个梯队: 块 1/3 与 2/4 严格有序; 块 0/2 允许个位数违例 (头部份额折中).
		for b, allow := range map[int]int{0: 6, 1: 0, 2: 0} {
			hi, lo := positive(g.provinceBlocks[b]), positive(g.provinceBlocks[b+2])
			bad := 0
			for _, va := range hi {
				for _, vb := range lo {
					if va <= vb {
						bad++
					}
				}
			}
			if bad > allow {
				t.Errorf("种子 %d: 块 %d/%d 相隔违例对 = %d, 超过允许 %d", seed, b, b+2, bad, allow)
			}
		}
	}
	// 头部交界保持部分有序: 平均逆序低于随机混排基线 20 (8x5 块的一半).
	if avg := headInversions / 10; avg >= 20 {
		t.Errorf("第 1/2 交界 10 种子平均逆序对 = %d, 未低于随机混排基线 20 (应保持部分有序)", avg)
	}
}
