// FilePath    : internal/generator/orders.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 入库/订单主/订单子/销售目标表 (T03-T06) 生成, 数据生成核心.

package generator

import (
	"sort"
	"strconv"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/csvw"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// monthOf 返回日期的月份 (1-12), 对应 VBA 的 Month().
func monthOf(t time.Time) int { return int(t.Month()) }

// orderWriters 聚合入库/订单主/订单子三张流式写出的 CSV 写入器.
type orderWriters struct {
	inv   *csvw.Writer // T03
	order *csvw.Writer // T04
	item  *csvw.Writer // T05
}

// close 关闭三张表的写入器.
// 返回值 error, 首个出错的写入器错误.
func (w *orderWriters) close() error {
	for _, x := range []*csvw.Writer{w.inv, w.order, w.item} {
		if x == nil {
			continue
		}
		if err := x.Close(); err != nil {
			return err
		}
	}
	return nil
}

// genOrders 流式生成 T03/T04/T05, 对应 DataTableT345, 是数据量最大的核心逻辑.
// 生成过程中同时按省 x 年 x 月汇总销售额, 供 T06 使用.
//   - startOC, 订单编号起始序号 (全量为 0, 增量为已有最大序号).
//   - progLo, progHi, 本阶段占用的进度区间.
//
// 返回值 int, 结束时的订单序号; error, 出错时非 nil.
func (g *Generator) genOrders(w *orderWriters, startOC int, progLo, progHi float64) (int, error) {
	productMaxIdx := len(g.products) - 1   // 产品索引上界
	customerMaxIdx := len(g.customers) - 1 // 客户索引上界
	ocNumber := startOC
	total := len(g.stores)

	for i1, s := range g.stores {
		if err := g.genStoreOrders(w, &s, i1, productMaxIdx, customerMaxIdx, &ocNumber); err != nil {
			return ocNumber, err
		}
		if total > 0 {
			g.progress(progLo+(progHi-progLo)*float64(i1+1)/float64(total), "stageOrders")
		}
	}
	return ocNumber, nil
}

// genStoreOrders 生成单个门店的订单/入库数据.
//   - s, 当前门店.
//   - i1, 门店索引.
//   - productMaxIdx, customerMaxIdx, 产品/客户索引上界.
//   - ocNumber, 全局订单序号指针.
//
// 返回值 error, 出错时非 nil.
//
//nolint:gocognit,gocyclo // 忠实移植原 VBA 的多重条件分支, 保持业务逻辑一致.
func (g *Generator) genStoreOrders(w *orderWriters, s *store, i1, productMaxIdx, customerMaxIdx int, ocNumber *int) error {
	// 营业天数: 已关店按关店日期, 否则按结束日期.
	end := g.cfg.EndDate
	if s.closeDate != nil {
		end = *s.closeDate
	}
	yyts := util.RoundInt(end.Sub(s.openDate).Hours() / 24)
	if yyts < 1 {
		return nil
	}

	invDays := g.buildInventoryDays(yyts)
	ptr := 0

	dict3 := make(map[string]int) // 产品编号 -> 自上次入库以来累计销量
	var dict3keys []string

	for i4 := 1; i4 <= yyts; i4++ {
		dateDD := addDays(s.openDate, i4-1)
		month := monthOf(dateDD)
		nd := util.RoundInt(g.rnd.F() * 4 * monthTrend[month-1] * regionOrderFactor[s.cityID%34] *
			g.orderVolumeFactor(s, dateDD, month))

		for i := 1; i <= nd; i++ {
			*ocNumber++
			oc := "OC_" + util.PadInt(*ocNumber, 7)
			sj := (g.rnd.F() + customerDist[(i4*i)%12]) / 2
			customerIdx := util.RoundInt(float64(customerMaxIdx) * sj)
			customerIdx = g.selectCustomer(s, customerIdx, customerMaxIdx, *ocNumber)

			channel := "线上"
			if sj >= 0.7 {
				channel = "线下"
			}
			customerCode := ""
			if customerIdx >= 0 && customerIdx <= customerMaxIdx {
				customerCode = g.customers[customerIdx].code
			}
			// 送货日期与增量路径统一为 下单日期 + 1 + Round(4*sj+8) (原 openDate+i4+R 为等价写法, 无行为变更).
			deliveryDate := addDays(dateDD, util.RoundInt(4*sj+8)+1)
			g.idOrder++
			if err := w.order.Write([]string{strconv.Itoa(g.idOrder), oc, s.code, dateStr(dateDD), dateStr(deliveryDate), customerCode, channel}); err != nil {
				return err
			}
			g.nOrders++

			if err := g.genOrderItems(w, s, i1, productMaxIdx, customerIdx, oc, dateDD, month, dict3, &dict3keys); err != nil {
				return err
			}
		}

		// 生成入库信息: 命中调度日 (非末日) 或到达末日时, 将累计销量作为入库量写出.
		if ptr < len(invDays) && invDays[ptr] == i4 && i4 < yyts {
			ptr++
			if err := g.flushInventory(w, s, i4, dict3, dict3keys, 0); err != nil {
				return err
			}
			dict3, dict3keys = make(map[string]int), nil
		} else if i4 == yyts {
			if err := g.flushInventory(w, s, i4, dict3, dict3keys, util.RoundInt(g.rnd.F()*5)); err != nil {
				return err
			}
			dict3, dict3keys = make(map[string]int), nil
		}
	}
	return nil
}

// genOrderItems 生成单个订单的订单子表明细并累计入库/销售汇总.
//   - s, 门店; i1, 门店索引; productMaxIdx, 产品索引上界; customerIdx, 选中的客户索引.
//   - oc, 订单编号; dateDD, 下单日期; month, 月份.
//   - dict3, dict3keys, 入库累计器.
//
// 返回值 error, 出错时非 nil.
func (g *Generator) genOrderItems(w *orderWriters, s *store, i1, productMaxIdx, customerIdx int, oc string, dateDD time.Time, month int, dict3 map[string]int, dict3keys *[]string) error {
	k := util.RoundInt(5*g.rnd.F()) + 1 // 每单产品种类数上限, 均值约 3
	skuSet := make(map[int]struct{}, k)
	var skuOrder []int
	for n := 1; n <= k; n++ {
		var productIdx int
		if k < 4 {
			productIdx = util.RoundInt(float64(productMaxIdx) * g.rnd.F() / 5) // 往左偏移
		} else {
			productIdx = util.RoundInt(float64(productMaxIdx) * g.rnd.F())
		}
		if _, ok := skuSet[productIdx]; !ok {
			skuSet[productIdx] = struct{}{}
			skuOrder = append(skuOrder, productIdx)
		}
	}

	for _, skuIdx := range skuOrder {
		p := util.RoundInt(5*g.rnd.F()*unitCountFactor[i1%5]*unitCountFactor[skuIdx%5]*productPopularity[skuIdx%29]) + 1
		q := g.discount(i1, skuIdx, customerIdx, month)
		prod := g.products[skuIdx]
		amount := util.RoundBankers(prod.salePrice*float64(p)*q, 2)
		g.idItem++
		if err := w.item.Write([]string{strconv.Itoa(g.idItem), oc, prod.code, ff(prod.salePrice), ff(util.RoundBankers(q, 2)), strconv.Itoa(p), ff(amount)}); err != nil {
			return err
		}
		g.nItems++
		if _, ok := dict3[prod.code]; !ok {
			*dict3keys = append(*dict3keys, prod.code)
		}
		dict3[prod.code] += p
		g.accumulateProvinceSales(s.cityID, dateDD, amount)
	}
	return nil
}

// discount 计算订单子表的折扣比例, 忠实移植原 VBA 的多重条件判定.
//   - i1, 门店索引; skuIdx, 产品索引; customerIdx, 客户索引; month, 月份.
//
// 返回值 float64, 折扣比例.
func (g *Generator) discount(i1, skuIdx, customerIdx, month int) float64 {
	monthDiscountCoeff := discountMonth[month-1]
	switch {
	case i1%40 > 30:
		return discounts[0] * monthDiscountCoeff
	case i1%40 < 10:
		return discounts[5] * monthDiscountCoeff
	case skuIdx%8 < 1:
		return discounts[1] * monthDiscountCoeff
	case skuIdx%8 > 5:
		return discounts[3] * monthDiscountCoeff
	}
	if customerIdx >= 0 && customerIdx < len(g.customers) {
		c := g.customers[customerIdx]
		if _, ok := discountIndustries[c.industry]; ok {
			return discounts[2] * monthDiscountCoeff
		}
		if _, ok := discountProfessions[c.profession]; ok {
			return discounts[4] * monthDiscountCoeff
		}
	}
	return 1
}

// selectCustomer 依据注册日期与月份启发式选择下单客户, 忠实移植原 VBA 的搜索逻辑.
//   - s, 门店; customerIdx, 初始客户索引; customerMaxIdx, 客户索引上界; ocNumber, 当前订单序号.
//
// 返回值 int, 选中的客户索引; 无客户时为 -1.
//
//nolint:gocognit,gocyclo // 忠实移植原 VBA 的左右搜索分支.
func (g *Generator) selectCustomer(s *store, customerIdx, customerMaxIdx, ocNumber int) int {
	if customerMaxIdx < 0 {
		return -1
	}
	if customerIdx < 0 {
		customerIdx = 0
	}
	if customerIdx > customerMaxIdx {
		customerIdx = customerMaxIdx
	}
	open := s.openDate
	reg := func(i int) time.Time { return g.customers[i].regDate }

	if s.closeDate == nil { // 未关店
		if !reg(customerIdx).Before(open) && ocNumber%13 > 6 {
			return customerIdx
		}
		for i2 := customerIdx; i2 <= customerMaxIdx; i2++ { // 往右
			m := monthOf(reg(i2))
			if !reg(i2).Before(open) && (m > 8 || m%3 > 1) {
				return i2
			}
		}
		for i2 := customerIdx; i2 >= 0; i2-- { // 往左
			m := monthOf(reg(i2))
			if !reg(i2).Before(open) && (m <= 8 || m%3 <= 1) {
				return i2
			}
		}
		return customerIdx
	}

	// 已关店
	close := *s.closeDate
	inRange := func(i int) bool { return !reg(i).Before(open) && reg(i).Before(close) }
	if inRange(customerIdx) && ocNumber%13 < 6 {
		return customerIdx
	}
	for i2 := customerIdx; i2 <= customerMaxIdx; i2++ { // 往右
		m := monthOf(reg(i2))
		if inRange(i2) && m > 8 {
			return i2
		}
		if !reg(i2).Before(open) && m%3 > 1 {
			return i2
		}
	}
	for i2 := customerIdx; i2 >= 0; i2-- { // 往左
		m := monthOf(reg(i2))
		if inRange(i2) && (m <= 8 || m%3 <= 1) {
			return i2
		}
	}
	return customerIdx
}

// buildInventoryDays 构建入库调度日序列, 对应原 Dict3N.
//   - yyts, 营业天数.
//
// 返回值 []int, 递增的入库日 (1-based, 位于 [1, yyts]).
func (g *Generator) buildInventoryDays(yyts int) []int {
	d := yyts%g.cfg.InventoryCycle + 1
	days := []int{d}
	for {
		d += util.RoundInt(g.rnd.F()*2 + 5)
		if d > yyts {
			break
		}
		days = append(days, d)
	}
	return days
}

// flushInventory 将累计销量作为入库量写出到 T03, 并可附加末次缓冲量.
//   - s, 门店; i4, 当前营业日序号.
//   - dict3, dict3keys, 累计器; extra, 末次入库附加量.
//
// 返回值 error, 出错时非 nil.
func (g *Generator) flushInventory(w *orderWriters, s *store, i4 int, dict3 map[string]int, dict3keys []string, extra int) error {
	date := dateStr(addDays(s.openDate, i4-1)) // -1 保证有库存
	for _, code := range dict3keys {
		qty := dict3[code] + extra
		g.idInv++
		if err := w.inv.Write([]string{strconv.Itoa(g.idInv), code, strconv.Itoa(qty), s.code, date}); err != nil {
			return err
		}
		g.nInv++
	}
	return nil
}

// orderVolumeFactor 计算单店单日的订单量综合系数: 年景 x 月度噪声 x 星期 x 门店客流.
// 全部为按日期/门店 ID 确定性索引的固定系数, 同一日期与门店在全量与增量两次生成中取值一致.
//   - s, 门店; dateDD, 日期; month, 月份 (1-12).
//
// 返回值 float64, 综合系数.
func (g *Generator) orderVolumeFactor(s *store, dateDD time.Time, month int) float64 {
	year := dateDD.Year()
	return yearTrend[year%len(yearTrend)] *
		monthNoise[(year*12+month)%len(monthNoise)] *
		weekdayFactor[dateDD.Weekday()] *
		storeTrafficFactor[s.id%len(storeTrafficFactor)]
}

// accumulateProvinceSales 按省 x 年 x 月汇总销售额, 供 T06 以实绩为锚生成销售目标;
// 同时维护事实月份区间 [factMinYM, factMaxYM] (YYYYMM), T06 仅覆盖该区间.
//   - cityID, 门店城市 ID; dateDD, 下单日期; amount, 销售金额.
func (g *Generator) accumulateProvinceSales(cityID int, dateDD time.Time, amount float64) {
	pid, ok := g.ds.CityToProvince[cityID]
	if !ok {
		return
	}
	year, month := dateDD.Year(), monthOf(dateDD)
	g.provinceMonth[monthKey{pid: pid, year: year, month: month}] += amount
	ym := year*100 + month
	if g.factMinYM == 0 || ym < g.factMinYM {
		g.factMinYM = ym
	}
	if ym > g.factMaxYM {
		g.factMaxYM = ym
	}
}

// genSaleTargets 生成 (或增量重建时覆写) 销售目标表 T06.
// 与原 VBA 方案 (仅最近两年, 由去年均值外推目标, 完成率易整体偏离) 不同:
//  1. 覆盖区间为 首个销售事实年度的下一年 1 月 至 最后事实年度的 12 月 (年末),
//     如事实 [2022-08, 2026-08] 对应目标 [2023-01, 2026-12], 最后一年为完整年度目标;
//  2. 事实期内的月份以当月实绩为锚, 目标 = 实绩 x [0.92, 1.04] 确定性系数 x 年度偏置,
//     完成率 (实绩/目标) 落在约 [90.7%, 114.4%], 满足 90%-120% 的预期区间;
//  3. 事实期之后至年末的计划月份无实绩可锚, 以该省近 12 个月月均实绩 x 行业淡旺季外推;
//  4. 事实期内当月无实绩的省份目标记 0 (该省当月无门店产生销售);
//  5. 目标系数与年度偏置均为 (省, 年月/年份) 的确定性函数, 增量重建不改变未变化月份的目标值.
//
// 返回值 error, 出错时非 nil.
func (g *Generator) genSaleTargets() error {
	w, err := csvw.Create(g.cfg.OutputDir, model.FileSaleTarget, model.HeaderSaleTarget)
	if err != nil {
		return err
	}

	pidSet := make(map[int]struct{}, len(g.provinceMonth))
	for k := range g.provinceMonth {
		pidSet[k.pid] = struct{}{}
	}
	pids := make([]int, 0, len(pidSet))
	for pid := range pidSet {
		pids = append(pids, pid)
	}
	sort.Ints(pids)

	id := 0
	for _, pid := range pids {
		if id, err = g.writeProvinceTargets(w, pid, id); err != nil {
			util.SilentClose(w.Close)
			return err
		}
	}
	return w.Close()
}

// saleTargetRange 返回 T06 销售目标应覆盖的月份区间 (YYYYMM, 含).
// 规则: 从首个销售事实年度的下一年 1 月起, 到最后事实年度的 12 月止;
// 事实仅覆盖单一年度时退化为覆盖该年全年, 避免 T06 为空.
//
// 返回值 int, 起始月份 YYYYMM; int, 结束月份 YYYYMM.
func (g *Generator) saleTargetRange() (int, int) {
	startYear := g.factMinYM/100 + 1
	endYear := g.factMaxYM / 100
	if startYear > endYear {
		startYear = endYear
	}
	return startYear*100 + 1, endYear*100 + 12
}

// writeProvinceTargets 写出单个省在目标区间内逐月的目标行.
//   - w, CSV 写入器; pid, 省 ID; id, 当前的 F_00_自动编号.
//
// 返回值 int, 结束时的自动编号; error, 出错时非 nil.
func (g *Generator) writeProvinceTargets(w *csvw.Writer, pid, id int) (int, error) {
	short := g.ds.ProvinceByID[pid].Short2
	startYM, endYM := g.saleTargetRange()
	recent := g.provinceRecentMean(pid)
	for year := startYM / 100; year <= endYM/100; year++ {
		for k := 1; k <= 12; k++ {
			ym := year*100 + k
			if ym < startYM || ym > endYM {
				continue
			}
			id++
			row := []string{strconv.Itoa(id), strconv.Itoa(pid), short, monthDateStr(year, k),
				ff(g.saleTargetValue(pid, year, k, recent))}
			if err := w.Write(row); err != nil {
				return id, err
			}
		}
	}
	return id, nil
}

// saleTargetValue 计算某省某月的目标值 (银行家舍入到整数).
// 事实期内 (不超过最晚事实月份) 以当月实绩为锚; 之后的计划月份以近期月均 x 淡旺季外推.
// 目标系数由 (省 ID, 年月) 确定性导出并叠加年度偏置, 未变化月份在增量重建后目标值保持不变.
//   - pid, 省 ID; year, month, 目标年月; recent, 该省近 12 个事实月份的月均实绩 (计划期外推基数).
//
// 返回值 float64, 目标值.
func (g *Generator) saleTargetValue(pid, year, month int, recent float64) float64 {
	// 目标系数在 [0.92, 1.04] 内确定性浮动, 叠加年度偏置后乘积落在约 [0.874, 1.102],
	// 保证事实期内完成率约在 [90.7%, 114.4%], 满足 90%-120% 的预期区间.
	factor := 0.92 + targetFactorUniform(pid, year*100+month)*0.12
	factor *= 1 + targetYearBias[year%len(targetYearBias)]
	if year*100+month <= g.factMaxYM {
		return util.RoundBankers(g.provinceMonth[monthKey{pid: pid, year: year, month: month}]*factor, 0)
	}
	// 计划期 (事实期之后到年末): 近 12 个月月均 x 行业淡旺季 (月度趋势均值近似 1) 外推.
	return util.RoundBankers(recent*monthTrend[month-1]*factor, 0)
}

// targetFactorUniform 从 (省 ID, 年月) 确定性导出 [0, 1) 的伪随机数 (整型乘法散列),
// 供销售目标系数使用: 同一省月的目标系数在全量生成与增量重建两次生成中保持一致,
// 增量更新不会改变既有月份的目标值.
//   - pid, 省 ID; ym, 年月 YYYYMM.
//
// 返回值 float64, [0, 1) 的伪随机数.
func targetFactorUniform(pid, ym int) float64 {
	x := uint64(pid)*2654435761 + uint64(ym)*40503 // #nosec G115 正整数相乘, uint64 无溢出风险
	x ^= x >> 13
	x *= 1274126177
	x ^= x >> 16
	return float64(x%1000003) / 1000003.0
}

// provinceRecentMean 返回指定省近 12 个事实月份 (不足取全部) 有实绩月份的月均销售额,
// 作为事实期之后计划月份的目标外推基数; 近 12 个月均无实绩时退化为全部月份月均.
//   - pid, 省 ID.
//
// 返回值 float64, 月均销售额.
func (g *Generator) provinceRecentMean(pid int) float64 {
	start := ymAddMonths(g.factMaxYM, -11)
	var recentSum float64
	var recentN int
	var allSum float64
	var allN int
	for k, v := range g.provinceMonth {
		if k.pid != pid {
			continue
		}
		allSum += v
		allN++
		if k.year*100+k.month >= start {
			recentSum += v
			recentN++
		}
	}
	switch {
	case recentN > 0:
		return recentSum / float64(recentN)
	case allN > 0:
		return allSum / float64(allN)
	default:
		return 0
	}
}

// ymAddMonths 返回 YYYYMM 月份加减 n 个月后的值, 供月份区间回溯.
//   - ym, 起始月份 YYYYMM; n, 偏移月数 (可为负).
//
// 返回值 int, 结果月份 YYYYMM.
func ymAddMonths(ym, n int) int {
	total := ym/100*12 + ym%100 - 1 + n
	return total/12*100 + total%12 + 1
}

// monthDateStr 返回某年某月 1 号的日期字符串, 对应原 Format(... "YYYY-M-1").
//   - year, 年份; month, 月份.
//
// 返回值 string, 日期字符串.
func monthDateStr(year, month int) string {
	return time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.UTC).Format(dateLayout)
}
