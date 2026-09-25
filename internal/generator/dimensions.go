// FilePath    : internal/generator/dimensions.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 维度表 (D00-D03) 与产品/门店/客户表 (T00-T02) 生成.

package generator

import (
	"math"
	"sort"
	"strconv"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/csvw"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// ff 将浮点数格式化为最短十进制字符串, 去除多余的尾随零.
//   - f, 待格式化的浮点数.
//
// 返回值 string, 格式化结果.
func ff(f float64) string {
	return strconv.FormatFloat(f, 'f', -1, 64)
}

// dateLayout CSV 与前后端交互使用的日期格式.
const dateLayout = "2006-01-02"

// dateStr 将日期格式化为 YYYY-MM-DD, 对应 VBA 的 Format(date, "YYYY-MM-DD").
//   - t, 日期.
//
// 返回值 string, 格式化结果.
func dateStr(t time.Time) string {
	return t.Format(dateLayout)
}

// addDays 返回 base 之后 days 天的日期 (days 可为负), 对应 VBA 的日期加减.
//   - base, 基准日期.
//   - days, 偏移天数.
//
// 返回值 time.Time, 结果日期.
func addDays(base time.Time, days int) time.Time {
	return base.AddDate(0, 0, days)
}

// writeDimensions 写出 D00-D03 四张维度表 (与原 DataTableD0-D3 对应).
// 返回值 error, 出错时非 nil.
func (g *Generator) writeDimensions() error {
	// D00 大区表
	if err := writeAll(g.cfg.OutputDir, model.FileRegion, model.HeaderRegion, len(g.ds.Regions),
		func(i int) []string {
			r := g.ds.Regions[i]
			return []string{strconv.Itoa(i + 1), strconv.Itoa(r.RegionID), r.Name, r.Manager,
				strconv.Itoa(r.CityID), r.CityName, ff(r.Lat), ff(r.Lng)}
		}); err != nil {
		return err
	}
	// D01 省份表
	if err := writeAll(g.cfg.OutputDir, model.FileProvince, model.HeaderProvince, len(g.ds.Provinces),
		func(i int) []string {
			p := g.ds.Provinces[i]
			return []string{strconv.Itoa(i + 1), strconv.Itoa(p.RegionID), strconv.Itoa(p.ProvinceID),
				p.Full, p.Short1, p.Short2, ff(p.Lat), ff(p.Lng)}
		}); err != nil {
		return err
	}
	// D02 城市表
	if err := writeAll(g.cfg.OutputDir, model.FileCity, model.HeaderCity, len(g.ds.Cities),
		func(i int) []string {
			c := g.ds.Cities[i]
			return []string{strconv.Itoa(i + 1), strconv.Itoa(c.ProvinceID), strconv.Itoa(c.CityID),
				c.Name, ff(c.Lat), ff(c.Lng)}
		}); err != nil {
		return err
	}
	// D03 区县表
	return writeAll(g.cfg.OutputDir, model.FileDistrict, model.HeaderDistrict, len(g.ds.Districts),
		func(i int) []string {
			d := g.ds.Districts[i]
			return []string{strconv.Itoa(i + 1), strconv.Itoa(d.CityID), strconv.Itoa(d.DistrictID),
				d.Name, ff(d.Lat), ff(d.Lng)}
		})
}

// 产品售价对数正态参数: 家具品类单价较高, 价格区间 [1000, 30000],
// 取 ln(6000) 为均值、0.85 为标准差, 中位数约 6000 (约为区间上限的 20%, 与原 [100, 15000] 设计比例一致),
// 多数偏低少数高价 (上下各约 2% 触界截断),
// 替代原 VBA 与分类强绑定的均匀分布 [5000, 10000].
const (
	priceLogMu    = 8.699514748210192 // ln(6000)
	priceLogSigma = 0.85
	priceMin      = 1000
	priceMax      = 30000
)

// stdNormalCDF 返回标准正态分布的累积分布函数 Φ(z), 用于将正态抽样映射为均匀分位.
//   - z, 标准正态随机数.
//
// 返回值 float64, 累积概率 [0, 1].
func stdNormalCDF(z float64) float64 {
	return 0.5 * (1 + math.Erf(z/math.Sqrt2))
}

// genProducts 生成产品表 T00, 对应 DataTableT0.
func (g *Generator) genProducts() {
	n := g.cfg.ProductCount
	g.products = make([]product, 0, n)
	seen := make(map[string]struct{}, n) // 产品名称去重
	for i := 1; i <= n; i++ {
		z := g.rnd.Norm()
		// 售价对数正态抽样并截断到 [1000, 30000]; 分类仍按价格分位取 A-J (A 类最便宜, J 类最贵).
		r4 := math.Min(math.Max(math.Exp(priceLogMu+priceLogSigma*z), priceMin), priceMax)
		letter := util.Letter(util.RoundInt(stdNormalCDF(z) * 9))
		// 成本比例独立抽样 (原 VBA 与售价共用同一随机数): 约 28% 的产品为 0.18, 其余在 [0.28, 1) 内.
		var r5 float64
		if ratio := g.rnd.F(); ratio >= 0.28 {
			r5 = r4 * ratio
		} else {
			r5 = r4 * 0.18
		}
		// 产品名称缩短为 产品 + 3 位字母数字组合 (如 产品A7X), 替代原 VBA 的 产品B0122 格式, 名称保证唯一.
		var name string
		for {
			name = "产品" + g.randProductCode()
			if _, ok := seen[name]; !ok {
				seen[name] = struct{}{}
				break
			}
		}
		g.products = append(g.products, product{
			id:        i,
			code:      "SKU_" + util.PadInt(i, 6),
			category:  letter + "类",
			name:      name,
			salePrice: util.RoundBankers(r4, 2),
			costPrice: util.RoundBankers(r5, 2),
		})
	}
}

// randProductCode 生成 3 位 "字母+数字" 混合组合 (至少含一个字母与一个数字, 如 A7X),
// 作为产品名称 产品XXX 中的 XXX 部分.
// 返回值 string, 3 位字母数字组合.
func (g *Generator) randProductCode() string {
	for {
		code := ""
		hasLetter, hasDigit := false, false
		for range 3 {
			v := util.RoundInt(g.rnd.F() * 35) // 0-25 映射字母 A-Z, 26-35 映射数字 0-9 (乘 35 保证舍入后不超过 35)
			if v < 26 {
				code += util.Letter(v)
				hasLetter = true
			} else {
				code += strconv.Itoa(v - 26)
				hasDigit = true
			}
		}
		if hasLetter && hasDigit {
			return code
		}
	}
}

// writeProducts 将产品表写出为 CSV.
// 返回值 error, 出错时非 nil.
func (g *Generator) writeProducts() error {
	return writeAll(g.cfg.OutputDir, model.FileProduct, model.HeaderProduct, len(g.products),
		func(i int) []string {
			p := g.products[i]
			return []string{strconv.Itoa(p.id), p.code, p.category, p.name, ff(p.salePrice), ff(p.costPrice)}
		})
}

// fixedStore 原 VBA Arr7 中四直辖市 + 港澳台的固定门店信息 (地理字段对齐 VBA 版本使用城市).
type fixedStore struct {
	code    string
	manager string
	cityID  int
	city    string
	lat     float64
	lng     float64
}

// arr7 前七个优先命中的固定门店 (与原 VBA Arr7 一致), 城市 ID/名称/经纬度取自 D02_城市表 数据源.
var arr7 = []fixedStore{
	{"SC_0001", "焦阿大", 110000, "北京", 39.904989, 116.405285},
	{"SC_0002", "焦阿二", 120000, "天津", 39.125596, 117.190182},
	{"SC_0003", "焦阿三", 310000, "上海", 31.231706, 121.472644},
	{"SC_0004", "焦阿四", 500000, "重庆", 29.533155, 106.504962},
	{"SC_0005", "焦阿五", 710000, "台湾", 25.044332, 121.509062},
	{"SC_0006", "焦阿六", 810000, "香港", 22.320048, 114.173355},
	{"SC_0007", "焦阿七", 820000, "澳门", 22.198951, 113.54909},
}

// randStoreName 生成 N 个不重复的 "XYZ店" 随机门店名, 对应原 Dict1 逻辑.
//   - n, 需要的门店名数量.
//
// 返回值 []string, 门店名列表.
func (g *Generator) randStoreNames(n int) []string {
	seen := make(map[string]struct{}, n)
	names := make([]string, 0, n)
	for len(names) < n {
		nm := util.Letter(util.RoundInt(g.rnd.F()*25)) +
			util.Letter(util.RoundInt(g.rnd.F()*25)) +
			util.Letter(util.RoundInt(g.rnd.F()*25)) + "店"
		if _, ok := seen[nm]; ok {
			continue
		}
		seen[nm] = struct{}{}
		names = append(names, nm)
	}
	return names
}

// randName 生成随机姓名, 对应原 Sj<阈值 决定二字/三字 的逻辑.
//   - twoCharThreshold, 二字姓名的 Sj 阈值.
//   - sj, 已抽取的随机数.
//
// 返回值 string, 姓名.
func (g *Generator) randName(twoCharThreshold, sj float64) string {
	fn := g.ds.FirstNames
	ln := g.ds.LastNames
	fnUB0 := util.RoundInt(float64(len(fn)-1) * sj)
	fnUB1 := util.RoundInt(float64(len(fn)-1) * (1 - sj))
	lnUB := util.RoundInt(float64(len(ln)-1) * sj)
	if sj < twoCharThreshold {
		return ln[lnUB] + fn[fnUB0]
	}
	return ln[lnUB] + fn[fnUB0] + fn[fnUB1]
}

// provinceStoreSequence 生成 m 个随机门店的省份交错序列:
// 有效权重为 省份销售规模权重 / 该省城市订单量系数均值 (补偿城市系数与地理编码
// 相关的系统性偏差); 各省门店数取 "含固定门店的总配额" 的随机舍入 (期望精确等于
// 配额, 总量偏差按剩余赤字最小修正), 再按分位交错铺开与随机轮转.
// 固定门店的省份 (北京/天津/上海/重庆/台湾/香港/澳门) 预置计数, 随机分配自动补偿;
// 门店数取整的 ±2 倍摆动由 assignProvinceStoreComp 在订单量端反向吸收.
//   - m, 需要分配的门店数 (门店总数减去 7 个固定门店).
//
// 返回值 []int, 长度为 m 的省 ID 序列; map[int]float64, 各省的期望随机门店数 (供省内分层抽城).
func (g *Generator) provinceStoreSequence(m int) ([]int, map[int]float64) {
	g.ensureProvinceWeights()
	pids := sortedProvinceIDs(g.provinceQuotaWeight)
	w := make([]float64, len(pids))
	var sumW float64
	for i, pid := range pids {
		w[i] = g.provinceQuotaWeight[pid] / g.provinceCityFactor[pid]
		sumW += w[i]
	}
	counts := make([]int, len(pids)) // 固定门店预置计数
	for _, a := range arr7 {
		pid, ok := g.ds.CityToProvince[a.cityID]
		if !ok {
			continue
		}
		for i, p := range pids {
			if p == pid {
				counts[i]++
				break
			}
		}
	}
	// 期望随机门店数 = 含固定门店的总配额份额 - 已预置的固定门店数 (供省内分层抽城).
	ideal := make([]float64, len(pids))
	idealMap := make(map[int]float64, len(pids))
	for i, pid := range pids {
		ideal[i] = max(float64(m+len(arr7))*w[i]/sumW-float64(counts[i]), 0.1)
		idealMap[pid] = ideal[i]
	}
	target := stochasticStoreTargets(ideal, m, g.rnd)
	seq := make([]int, 0, m)
	for _, i := range interleaveProvinceSlots(target, g.rnd) {
		seq = append(seq, pids[i])
	}
	// 随机轮转: 使 编号->省份 的模式随种子变化, 避免编号哈希化的门店级因子
	// (折扣类别/客流/单均件数) 与固定模式耦合形成系统性偏差.
	if len(seq) > 1 {
		off := util.RoundInt(g.rnd.F() * float64(len(seq)))
		rotated := make([]int, len(seq))
		copy(rotated, seq[off:])
		copy(rotated[len(seq)-off:], seq[:off])
		return rotated, idealMap
	}
	return seq, idealMap
}

// stochasticStoreTargets 将各省的期望随机门店数随机舍入为整数目标:
// 期望精确等于配额份额. 确定性的最大余数取整会把 "+1" 固定给同一批省份,
// 使低权重梯队的门店数被系统性压低, 块级销售额份额偏离期望.
// 舍入的零星总量偏差按 |剩余赤字| 最小的省份逐个增减一, 保证总数恰为 m.
//   - ideal, 各省期望随机门店数; m, 需要分配的门店总数; rnd, 随机源.
//
// 返回值 []int, 各省的整数门店数目标.
func stochasticStoreTargets(ideal []float64, m int, rnd *util.Rand) []int {
	target := make([]int, len(ideal))
	sumT := 0
	for i, x := range ideal {
		//nolint:gosec // 期望为正, 随机舍入结果非负
		target[i] = stochasticFloor(x, rnd)
		sumT += target[i]
	}
	for sumT != m {
		best := 0
		bestGap := math.Inf(1)
		for i := range ideal {
			gap := math.Abs(ideal[i] - float64(target[i]))
			if (sumT > m && target[i] > 0 || sumT < m) && gap < bestGap {
				bestGap, best = gap, i
			}
		}
		if sumT > m {
			target[best]--
			sumT--
		} else {
			target[best]++
			sumT++
		}
	}
	return target
}

// stochasticFloor 返回 x 的随机舍入 (向下取整并以小数部分为概率进一, 期望恰为 x):
// 供门店配额与日订单量取整. 确定性四舍五入对 <0.5 的期望恒取 0, 会使省份五梯队
// 末端的单店整月无单、销售额归零; 随机舍入支持稀疏出单且期望保持无偏.
//   - x, 期望值 (须非负); rnd, 随机源.
//
// 返回值 int, 舍入结果.
func stochasticFloor(x float64, rnd *util.Rand) int {
	if x <= 0 {
		return 0
	}
	f := math.Floor(x)
	if rnd.F() < x-f {
		return int(f) + 1 // #nosec G115 x > 0 时结果为正整数, 无溢出
	}
	return int(f) // #nosec G115 同上
}

// interleaveProvinceSlots 将各省的门店数目标排成交错序列:
// 各省的第 k 家门店落在 (k+0.5)/n 分位附近 (加微抖动打散同分位), 使同省门店编号
// 均匀铺开; 索引并列时按省下序升序保证确定.
//   - target, 各省的门店数目标; rnd, 随机源 (分位抖动).
//
// 返回值 []int, 长度为 Σtarget 的省下标序列.
func interleaveProvinceSlots(target []int, rnd *util.Rand) []int {
	type slotEntry struct {
		slot float64
		idx  int
	}
	var total int
	for _, n := range target {
		total += n
	}
	entries := make([]slotEntry, 0, total)
	for i, n := range target {
		for k := range n {
			entries = append(entries, slotEntry{(float64(k)+0.5)/float64(n) + rnd.F()*0.5/float64(n), i})
		}
	}
	sort.Slice(entries, func(a, b int) bool {
		if entries[a].slot != entries[b].slot {
			return entries[a].slot < entries[b].slot
		}
		return entries[a].idx < entries[b].idx
	})
	order := make([]int, 0, total)
	for _, e := range entries {
		order = append(order, e.idx)
	}
	return order
}

// sortedProvinceIDs 返回省 ID 权重表的键升序列表, 保证轮询打分的遍历次序确定.
//   - weights, 省 ID -> 权重.
//
// 返回值 []int, 升序省 ID 列表.
func sortedProvinceIDs(weights map[int]float64) []int {
	ids := make([]int, 0, len(weights))
	for pid := range weights {
		ids = append(ids, pid)
	}
	sort.Ints(ids)
	return ids
}

// pickCityInProvince 在指定省内按城市订单量系数分层抽取城市:
// 城市已按系数升序排列, 第 k 家门店 (0-based) 取分位 (k+0.5)/期望门店数,
// 使省内被抽中城市的系数均值贴近期望, 消除单城抽样对省份销售额排序的噪声
// (替代原区内均匀抽城; 仅 1 个城市的省份恒取该城市).
//   - pid, 省 ID; assigned, 该省已分配的随机门店数; ideal, 该省的期望随机门店数.
//
// 返回值 data.City, 抽中的城市.
func (g *Generator) pickCityInProvince(pid, assigned int, ideal float64) data.City {
	cities := g.provinceCities[pid]
	q := (float64(assigned) + 0.5) / math.Max(ideal, 1)
	idx := util.RoundInt(min(q, 0.999) * float64(len(cities)-1))
	return cities[idx]
}

// balancedStoreClass 按门店在其省内的规模排名均衡分配折扣策略类别 (0-39):
// 每连续 4 家门店构成 1 深折 (0-9) + 2 中档 (10-30) + 1 溢价 (31-39), 与总体类别分布近似.
// 排名以预期规模 (营业天数 x 城市系数 x 客流系数) 降序计算, 保证每个省
// 深折扣门店的 "数量与权重" 双均衡, 不再随机聚集干扰省份梯队排序;
// 规模要素全部可由产物 CSV 重建, 全量生成与增量载入取值一致.
//   - j, 门店在其省内的规模排名 (0-based, 0 为预期规模最大).
//
// 返回值 int, 折扣策略类别.
func balancedStoreClass(j int) int {
	switch j % 4 {
	case 0:
		return j / 4 % 10
	case 3:
		return 31 + j/4%9
	default:
		return 10 + (j/4*2+j%2)%21
	}
}

// assignDiscountClasses 为全部门店按省分配折扣策略类别:
// 同省门店先按预期规模 (营业天数 x 城市订单量系数 x 客流系数) 降序排名, 再轮转取
// balancedStoreClass 的均衡类别; 门店数不足 3 家的省份全部取中档类别 (10-30) 并按
// 名次错开, 避免个位数门店的省份被单独一家深折/溢价门店主导而干扰省份梯队排序.
// 不同省份的类别权重构成一致, 且规模要素全部可由产物 CSV 重建, 全量生成与增量载入取值一致.
func (g *Generator) assignDiscountClasses() {
	idx := make(map[int][]int, 34)
	for i := range g.stores {
		pid, ok := g.ds.CityToProvince[g.stores[i].cityID]
		if !ok {
			pid = -1
		}
		idx[pid] = append(idx[pid], i)
	}
	for _, list := range idx {
		sort.Slice(list, func(a, b int) bool {
			wa, wb := g.storeExpectedVolume(&g.stores[list[a]]), g.storeExpectedVolume(&g.stores[list[b]])
			if wa != wb {
				return wa > wb
			}
			return g.stores[list[a]].id < g.stores[list[b]].id
		})
		if len(list) < 3 {
			for rank, i := range list {
				g.stores[i].discountClass = 10 + (rank*7)%21
			}
			continue
		}
		for rank, i := range list {
			g.stores[i].discountClass = balancedStoreClass(rank)
		}
	}
}

// storeExpectedVolume 估算门店的预期销售规模 (相对值), 供折扣类别的省内规模排名;
// 仅使用可从 T01 重建的要素 (营业天数, 城市订单量系数, 客流系数).
//   - s, 门店.
//
// 返回值 float64, 预期规模.
func (g *Generator) storeExpectedVolume(s *store) float64 {
	end := g.cfg.EndDate
	if s.closeDate != nil && s.closeDate.Before(end) {
		end = *s.closeDate
	}
	days := end.Sub(s.openDate).Hours() / 24
	if days < 0 {
		days = 0
	}
	return days * cityOrderFactor(s.cityID) * s.traffic
}

// genStores 生成门店表 T01, 对应 DataTableT1.
// 开店/关店日期不在创建时生成, 统一由 assignStratifiedDates 按省分层分配;
// 折扣策略类别由 assignDiscountClasses 按省内规模排名均衡分配.
func (g *Generator) genStores() {
	n := g.cfg.StoreCount
	names := g.randStoreNames(n)
	g.stores = make([]store, 0, n)

	appendStore := func(id int, code, name, manager string, cityID int, city string, lat, lng float64) {
		g.stores = append(g.stores, store{
			id: id, code: code, name: name, manager: manager,
			cityID: cityID, city: city, lat: lat, lng: lng,
		})
	}

	if n < 8 {
		for k := range n {
			a := arr7[k]
			appendStore(k+1, a.code, names[k], a.manager, a.cityID, a.city, a.lat, a.lng)
		}
		g.assignStratifiedDates()
		g.assignStoreTraffic()
		g.assignProvinceStoreComp()
		g.assignDiscountClasses()
		return
	}

	// N1 > 7: 先写 7 个固定门店
	for k := range 7 {
		a := arr7[k]
		appendStore(k+1, a.code, names[k], a.manager, a.cityID, a.city, a.lat, a.lng)
	}
	// 再生成剩余随机门店 (原 For i = 8 To N1), 省份按平滑加权轮询交错分配, 省内分层抽城.
	seq, ideal := g.provinceStoreSequence(n - 7)
	assigned := make(map[int]int, len(ideal))
	for i := 8; i <= n; i++ {
		sj := g.rnd.F()
		pid := seq[i-8]
		c := g.pickCityInProvince(pid, assigned[pid], ideal[pid])
		assigned[pid]++
		appendStore(i, "SC_"+util.PadInt(i, 4), names[i-1],
			g.randName(0.66, sj), c.CityID, c.Name,
			util.RoundBankers(c.Lat+g.rnd.F()*0.05, 6),
			util.RoundBankers(c.Lng+g.rnd.F()*0.05, 6))
	}
	g.assignStratifiedDates()
	g.assignStoreTraffic()
	g.assignProvinceStoreComp()
	g.assignDiscountClasses()
}

// assignStratifiedDates 为全部门店按省分层分配开店与关店日期:
// 每个省的开店时间在本省内均匀分层抽样 (替代原 VBA 的全省独立随机),
// 使各省营业天数总量贴近期望, 消除随机开店日期对省份梯队排序的噪声;
// 关店日期沿用原逻辑 (开店日 + 550 + Round(4320F), 早于窗口终点即记为已关店).
func (g *Generator) assignStratifiedDates() {
	span := max(g.windowDays-28, 1)
	idx := make(map[int][]int, 34)
	for i := range g.stores {
		pid, ok := g.ds.CityToProvince[g.stores[i].cityID]
		if !ok {
			pid = -1
		}
		idx[pid] = append(idx[pid], i)
	}
	// 按省 ID 升序遍历: 循环内消耗随机数, map 迭代顺序随机会导致同一种子两次生成结果不同.
	pids := make([]int, 0, len(idx))
	for pid := range idx {
		pids = append(pids, pid)
	}
	sort.Ints(pids)
	for _, pid := range pids {
		list := idx[pid]
		// 两店省份用同一 u 取抗互补分位 {u/2, 1-u/2}: 两店年龄之和恒定,
		// 消除省内营业天数总量的抽样方差.
		pairU := g.rnd.F()
		for j, i := range list {
			age := util.RoundInt(storeOpenQuantile(len(list), j, pairU, g.rnd)*float64(span)) + 28
			open := addDays(g.cfg.EndDate, -age)
			g.stores[i].openDate = open
			// dateGD > Now 表示尚未关店; 否则记录关店日期.
			if closeCandidate := addDays(open, 550+util.RoundInt(4320*g.rnd.F())); !closeCandidate.After(g.cfg.EndDate) {
				cd := closeCandidate
				g.stores[i].closeDate = &cd
			}
		}
	}
}

// storeOpenQuantile 返回省内第 j 家门店 (0-based) 的开店时间分位 (0-1):
// 单店省份收窄到窗口中段 (40%-60%), 消除单店营业天数在 [28, 窗口长] 上
// 均匀抽签的巨大方差; 两店省份用同一 u 取抗互补分位 {u/2, 1-u/2},
// 两店年龄之和恒定; 三店及以上沿用均匀分层抽样.
//   - n, 省内门店数; j, 门店序号 (0-based); u, 两店省份共享的随机数; rnd, 随机源.
//
// 返回值 float64, 开店时间分位.
func storeOpenQuantile(n, j int, u float64, rnd *util.Rand) float64 {
	switch n {
	case 1:
		return 0.4 + 0.2*rnd.F()
	case 2:
		if j == 0 {
			return u / 2
		}
		return 1 - u/2
	default:
		return (float64(j) + rnd.F()) / float64(n)
	}
}

// assignStoreTraffic 为全部门店分配客流系数: 同省门店按编号排名在 [0.75, 1.25] 线性
// 铺开, 省内均值恒为 1; 单店省份恒为 1, 消除单店省份客流抽签的方差.
// 分组与排名均可由 T01 (城市 -> 省, 门店编号) 确定性重建, 全量生成与增量载入一致.
func (g *Generator) assignStoreTraffic() {
	idx := g.groupStoresByProvince()
	for _, list := range idx {
		for j, i := range list {
			if len(list) == 1 {
				g.stores[i].traffic = 1
				continue
			}
			g.stores[i].traffic = 0.75 + 0.5*float64(j)/float64(len(list)-1)
		}
	}
}

// assignProvinceStoreComp 计算各省门店数取整的量级补偿:
// 补偿 = 含固定门店的总配额份额 / 实际门店数. 门店配额只能取整 (默认 55 家门店下
// 各省 1-2 家), 单纯依赖取整落点会使省份销售额出现 ±2 倍摆动, 且确定性取整会把 "+1"
// 固定给同一批省份形成块级系统偏置; 由单店订单量反向吸收取整偏差后,
// 省份期望销售额恒等于 配额 x 量级, 份额与块间排序同时保持稳定.
// 门店集合固定后调用 (全量生成与增量载入的 T01 相同, 补偿取值一致).
func (g *Generator) assignProvinceStoreComp() {
	g.ensureProvinceWeights()
	idx := g.groupStoresByProvince()
	// Σw 取全部梯队省份 (与门店配额同口径), 份额模型不受个别省份无门店影响.
	var sumW float64
	for pid, q := range g.provinceQuotaWeight {
		sumW += q / g.provinceCityFactor[pid]
	}
	total := float64(len(g.stores))
	g.provinceStoreComp = make(map[int]float64, len(idx))
	for pid, list := range idx {
		g.provinceStoreComp[pid] = total * (g.provinceQuotaWeight[pid] / g.provinceCityFactor[pid]) / sumW / float64(len(list))
	}
}

// groupStoresByProvince 按省聚合门店索引 (省内按门店编号升序, 确定性重建).
//
// 返回值 map[int][]int, 省 ID -> 门店索引列表.
func (g *Generator) groupStoresByProvince() map[int][]int {
	idx := make(map[int][]int, 34)
	for i := range g.stores {
		pid, ok := g.ds.CityToProvince[g.stores[i].cityID]
		if !ok {
			pid = -1
		}
		idx[pid] = append(idx[pid], i)
	}
	return idx
}

// writeStores 将门店表写出为 CSV.
// 返回值 error, 出错时非 nil.
func (g *Generator) writeStores() error {
	return writeAll(g.cfg.OutputDir, model.FileStore, model.HeaderStore, len(g.stores),
		func(i int) []string {
			s := g.stores[i]
			closeStr := ""
			if s.closeDate != nil {
				closeStr = dateStr(*s.closeDate)
			}
			return []string{strconv.Itoa(s.id), s.code, s.name, s.manager, dateStr(s.openDate),
				strconv.Itoa(s.cityID), s.city, ff(s.lat), ff(s.lng), closeStr}
		})
}

// genCustomers 生成客户表 T02, 对应 DataTableT2.
// 客户按门店规模注册, 关店门店不产生客户 (与原逻辑一致).
func (g *Generator) genCustomers() {
	g.customers = make([]customer, 0, 1024)
	halfWindow := float64(g.windowDays) / 2
	ii := 0
	for i, s := range g.stores {
		var n2 int
		if s.closeDate == nil {
			days := g.cfg.EndDate.Sub(s.openDate).Hours() / 24
			n2 = int(days * (1.2 + g.rnd.F())) // Int 截断
		}
		for k := 1; k <= n2; k++ {
			ii++
			sj := g.rnd.F()
			ageFactor := ageDist[ii%12]
			birth := addDays(g.cfg.EndDate, -(7500 + util.RoundInt((ageFactor+g.rnd.F())*7000)))
			reg := addDays(g.cfg.StartDate, util.RoundInt((ageFactor+g.rnd.F())*halfWindow))
			name := g.randName(0.8, sj)
			gender := "男"
			if sj >= 0.8 {
				gender = "女"
			}
			g.customers = append(g.customers, customer{
				id: ii, code: "CC_" + util.PadInt(ii, 7), name: name,
				birth: birth, gender: gender, regDate: reg,
				industry:   industries[util.RoundInt(g.rnd.F()*industryDist[i%7]*6)],
				profession: professions[util.RoundInt(g.rnd.F()*professionDist[i%7]*6)],
			})
		}
	}
}

// writeCustomers 将客户表写出为 CSV.
// 返回值 error, 出错时非 nil.
func (g *Generator) writeCustomers() error {
	return writeAll(g.cfg.OutputDir, model.FileCustomer, model.HeaderCustomer, len(g.customers),
		func(i int) []string {
			c := g.customers[i]
			return []string{strconv.Itoa(c.id), c.code, c.name, dateStr(c.birth), c.gender,
				dateStr(c.regDate), c.industry, c.profession}
		})
}

// writeAll 便捷函数: 创建 CSV 并逐行写出.
//   - dir, 目标目录.
//   - file, 文件名.
//   - header, 表头.
//   - count, 行数.
//   - row, 根据行索引生成一行记录的函数.
//
// 返回值 error, 出错时非 nil.
func writeAll(dir, file string, header []string, count int, row func(i int) []string) error {
	w, err := csvw.Create(dir, file, header)
	if err != nil {
		return err
	}
	for i := range count {
		if err = w.Write(row(i)); err != nil {
			util.SilentClose(w.Close)
			return err
		}
	}
	return w.Close()
}
