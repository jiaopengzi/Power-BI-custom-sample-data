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
// 生成过程中同时按省汇总销售额, 供 T06 使用.
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
		nd := util.RoundInt(g.rnd.F() * 4 * monthTrend[month-1] * regionOrderFactor[s.districtID%34])

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
			deliveryDate := addDays(s.openDate, i4+util.RoundInt(4*sj+8))
			if err := w.order.Write([]string{oc, s.code, dateStr(dateDD), dateStr(deliveryDate), customerCode, channel}); err != nil {
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
		p := util.RoundInt(5*g.rnd.F()*unitCountFactor[i1%5]*unitCountFactor[skuIdx%5]) + 1
		q := g.discount(i1, skuIdx, customerIdx, month)
		prod := g.products[skuIdx]
		amount := util.RoundBankers(prod.salePrice*float64(p)*q, 2)
		if err := w.item.Write([]string{oc, prod.code, ff(prod.salePrice), ff(util.RoundBankers(q, 2)), strconv.Itoa(p), ff(amount)}); err != nil {
			return err
		}
		g.nItems++
		if _, ok := dict3[prod.code]; !ok {
			*dict3keys = append(*dict3keys, prod.code)
		}
		dict3[prod.code] += p
		g.accumulateProvinceSales(s.districtID, dateDD, amount)
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
		if err := w.inv.Write([]string{code, strconv.Itoa(qty), s.code, date}); err != nil {
			return err
		}
		g.nInv++
	}
	return nil
}

// accumulateProvinceSales 按省汇总销售额, 用于 T06 的销售目标测算.
// 汇总窗口与原 VBA 一致: 去年全年 (含前年 12 月) 与去年 Q4 (9-12 月).
//   - districtID, 门店区县 ID; dateDD, 下单日期; amount, 销售金额.
func (g *Generator) accumulateProvinceSales(districtID int, dateDD time.Time, amount float64) {
	pid, ok := g.ds.DistrictToProvince[districtID]
	if !ok {
		return
	}
	year, month := dateDD.Year(), monthOf(dateDD)
	if year == g.endYear-1 || (year == g.endYear-2 && month == 12) {
		g.provinceFull[pid] += amount
	}
	if year == g.endYear-1 && month >= 9 {
		g.provinceQ4[pid] += amount
	}
}

// genSaleTargets 生成销售目标表 T06, 对应 DataTableT6.
// 返回值 error, 出错时非 nil.
func (g *Generator) genSaleTargets() error {
	w, err := csvw.Create(g.cfg.OutputDir, model.FileSaleTarget, model.HeaderSaleTarget)
	if err != nil {
		return err
	}

	// Qn 为最后三个月系数之和 (索引 9,10,11), 与原逻辑一致.
	qn := monthTrend[9] + monthTrend[10] + monthTrend[11]

	pids := make([]int, 0, len(g.provinceFull))
	for pid := range g.provinceFull {
		pids = append(pids, pid)
	}
	sort.Ints(pids)

	id := 0
	for _, pid := range pids {
		full := g.provinceFull[pid]
		q4, hasQ4 := g.provinceQ4[pid]
		short := g.ds.ProvinceByID[pid].Short2

		// 月均基数 B: Q4 缺失或全年月均更大时取全年月均, 否则取 Q4 月均.
		var b float64
		switch {
		case !hasQ4:
			b = full / 12
		case full/12 > q4/qn:
			b = full / 12
		default:
			b = q4 / qn
		}

		// 去年目标 (区域差距系数固定为 1).
		for k := 1; k <= 12; k++ {
			id++
			target := util.RoundBankers(b*(0.7+g.rnd.F()*0.1)*monthTrend[k-1], 0)
			month := monthDateStr(g.endYear-1, k)
			if err = w.Write([]string{strconv.Itoa(id), strconv.Itoa(pid), short, month, ff(target)}); err != nil {
				util.SilentClose(w.Close)
				return err
			}
		}
		// 今年目标 (regionGapFactor 依据省 ID, 由于省 ID 为 6 位编码实际恒为 1, 与原逻辑一致).
		regionGapFactor := provinceRegionGapFactor(pid)
		for k := 1; k <= 12; k++ {
			id++
			target := util.RoundBankers(b*(0.6+g.rnd.F()*0.2*regionGapFactor)*monthTrend[k-1], 0)
			month := monthDateStr(g.endYear, k)
			if err = w.Write([]string{strconv.Itoa(id), strconv.Itoa(pid), short, month, ff(target)}); err != nil {
				util.SilentClose(w.Close)
				return err
			}
		}
	}
	return w.Close()
}

// provinceRegionGapFactor 依据省 ID 返回区域差距系数, 忠实移植原 VBA 条件 (省 ID 为大编码时恒返回 1).
//   - pid, 省 ID.
//
// 返回值 float64, 区域差距系数.
func provinceRegionGapFactor(pid int) float64 {
	switch {
	case pid < 5:
		return 1.5
	case pid < 11:
		return 1.2
	default:
		return 1
	}
}

// monthDateStr 返回某年某月 1 号的日期字符串, 对应原 Format(... "YYYY-M-1").
//   - year, 年份; month, 月份.
//
// 返回值 string, 日期字符串.
func monthDateStr(year, month int) string {
	return time.Date(year, time.Month(month), 1, 0, 0, 0, 0, time.UTC).Format(dateLayout)
}
