// FilePath    : internal/generator/incremental.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 增量更新逻辑 (原 VBA 没有的新功能).

package generator

import (
	"bufio"
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/csvw"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// ErrNoBaseData 指定目录不存在基础示例数据, 需先执行全量生成.
var ErrNoBaseData = errors.New("no base data in output directory")

// DateConflictError 增量更新的日期区间与现有数据发生重叠.
type DateConflictError struct {
	ExistingMin time.Time
	ExistingMax time.Time
}

// Error 实现 error 接口.
// 返回值 string, 冲突描述.
func (e *DateConflictError) Error() string {
	return fmt.Sprintf("date range conflicts with existing data [%s, %s]",
		e.ExistingMin.Format(dateLayout), e.ExistingMax.Format(dateLayout))
}

// IncrementalUpdate 依据 [StartDate, EndDate] 生成增量订单/入库数据并追加到现有 CSV.
// 这是原 VBA 版本没有的新功能. 复用现有的产品/门店/客户维度, 不重算维度表与销售目标表.
//   - cfg, 增量配置 (日期区间为增量窗口).
//   - ds, 基础数据集 (用于区县->省映射, 兼容性保留).
//   - progress, 进度回调, 可为 nil.
//
// 返回值 Result, 增量新增行数; error, 无基础数据 (ErrNoBaseData) 或日期冲突 (*DateConflictError) 时非 nil.
func IncrementalUpdate(cfg *config.Config, ds *data.Dataset, progress ProgressFunc) (Result, error) {
	if progress == nil {
		progress = func(float64, string) {}
	}
	// 基础数据存在性检查.
	for _, f := range []string{model.FileProduct, model.FileStore, model.FileCustomer, model.FileOrder} {
		if _, err := os.Stat(filepath.Join(cfg.OutputDir, f)); err != nil {
			return Result{}, ErrNoBaseData
		}
	}

	g := New(cfg, ds, progress)
	if err := g.loadExistingDimensions(); err != nil {
		return Result{}, err
	}

	ocMax, minDate, maxDate, err := g.scanExistingOrders()
	if err != nil {
		return Result{}, err
	}
	// 日期冲突检查: 增量窗口与现有订单区间重叠即冲突.
	if !minDate.IsZero() && !cfg.StartDate.After(maxDate) && !cfg.EndDate.Before(minDate) {
		return Result{}, &DateConflictError{ExistingMin: minDate, ExistingMax: maxDate}
	}

	w, err := openAppendWriters(cfg.OutputDir)
	if err != nil {
		return Result{}, err
	}
	if err = g.genIncrementalOrders(w, ocMax); err != nil {
		util.SilentClose(w.close)
		return Result{}, err
	}
	if err = w.close(); err != nil {
		return Result{}, err
	}
	progress(100, "stageDone")

	return Result{
		Products:  len(g.products),
		Stores:    len(g.stores),
		Customers: len(g.customers),
		Inventory: g.nInv,
		Orders:    g.nOrders,
		OrderItem: g.nItems,
	}, nil
}

// openAppendWriters 以追加模式打开 T03/T04/T05 三张表.
// 返回值 *orderWriters, 写入器集合; error, 出错时非 nil.
func openAppendWriters(dir string) (*orderWriters, error) {
	inv, err := csvw.OpenAppend(dir, model.FileInventory)
	if err != nil {
		return nil, err
	}
	order, err := csvw.OpenAppend(dir, model.FileOrder)
	if err != nil {
		util.CloseQuietly(inv)
		return nil, err
	}
	item, err := csvw.OpenAppend(dir, model.FileOrderItem)
	if err != nil {
		util.CloseQuietly(inv)
		util.CloseQuietly(order)
		return nil, err
	}
	return &orderWriters{inv: inv, order: order, item: item}, nil
}

// genIncrementalOrders 为增量窗口内的每个门店生成订单/入库数据.
//   - ocMax, 现有最大订单序号.
//
// 返回值 error, 出错时非 nil.
func (g *Generator) genIncrementalOrders(w *orderWriters, ocMax int) error {
	productMaxIdx := len(g.products) - 1
	customerMaxIdx := len(g.customers) - 1
	ocNumber := ocMax
	total := len(g.stores)
	for i1, s := range g.stores {
		if err := g.genStoreOrdersRange(w, &s, i1, productMaxIdx, customerMaxIdx, &ocNumber); err != nil {
			return err
		}
		if total > 0 {
			g.progress(float64(i1+1)/float64(total)*100, "stageOrders")
		}
	}
	return nil
}

// genStoreOrdersRange 仅生成落在增量窗口 [StartDate, EndDate] 内的门店订单/入库数据.
// 与全量的按营业天数遍历不同, 此处按自然日期遍历窗口交集, 入库按入库周期定期结算.
//   - s, 门店; i1, 门店索引; productMaxIdx, customerMaxIdx, 产品/客户索引上界; ocNumber, 订单序号指针.
//
// 返回值 error, 出错时非 nil.
//
//nolint:gocognit,gocyclo // 忠实移植原 VBA 的多重条件分支, 保持业务逻辑一致.
func (g *Generator) genStoreOrdersRange(w *orderWriters, s *store, i1, productMaxIdx, customerMaxIdx int, ocNumber *int) error {
	// 门店自然营业结束日 (已关店取关店日, 否则取窗口结束日).
	naturalEnd := g.cfg.EndDate
	if s.closeDate != nil && s.closeDate.Before(naturalEnd) {
		naturalEnd = *s.closeDate
	}
	from := maxTime(s.openDate, g.cfg.StartDate)
	to := minTime(naturalEnd, g.cfg.EndDate)
	if from.After(to) {
		return nil
	}

	dict3 := make(map[string]int)
	var dict3keys []string
	dayCount := 0

	for dateDD := from; !dateDD.After(to); dateDD = addDays(dateDD, 1) {
		dayCount++
		month := monthOf(dateDD)
		nd := util.RoundInt(g.rnd.F() * 4 * monthTrend[month-1] * regionOrderFactor[s.districtID%34])
		for i := 1; i <= nd; i++ {
			*ocNumber++
			oc := "OC_" + util.PadInt(*ocNumber, 7)
			sj := (g.rnd.F() + customerDist[(dayCount*i)%12]) / 2
			customerIdx := g.selectCustomer(s, util.RoundInt(float64(customerMaxIdx)*sj), customerMaxIdx, *ocNumber)
			channel := "线上"
			if sj >= 0.7 {
				channel = "线下"
			}
			customerCode := ""
			if customerIdx >= 0 && customerIdx <= customerMaxIdx {
				customerCode = g.customers[customerIdx].code
			}
			deliveryDate := addDays(dateDD, util.RoundInt(4*sj+8)+1)
			if err := w.order.Write([]string{oc, s.code, dateStr(dateDD), dateStr(deliveryDate), customerCode, channel}); err != nil {
				return err
			}
			g.nOrders++
			if err := g.genOrderItems(w, s, i1, productMaxIdx, customerIdx, oc, dateDD, month, dict3, &dict3keys); err != nil {
				return err
			}
		}
		// 按入库周期或窗口末日结算入库.
		if (dayCount%g.cfg.InventoryCycle == 0 && !dateDD.Equal(to)) || dateDD.Equal(to) {
			extra := 0
			if dateDD.Equal(to) {
				extra = util.RoundInt(g.rnd.F() * 5)
			}
			if err := g.flushInventoryOn(w, s, dateDD, dict3, dict3keys, extra); err != nil {
				return err
			}
			dict3, dict3keys = make(map[string]int), nil
		}
	}
	return nil
}

// flushInventoryOn 在指定日期结算入库 (增量模式使用绝对日期).
//   - s, 门店; date, 入库日期; dict3, dict3keys, 累计器; extra, 附加量.
//
// 返回值 error, 出错时非 nil.
func (g *Generator) flushInventoryOn(w *orderWriters, s *store, date time.Time, dict3 map[string]int, dict3keys []string, extra int) error {
	d := dateStr(date)
	for _, code := range dict3keys {
		if err := w.inv.Write([]string{code, strconv.Itoa(dict3[code] + extra), s.code, d}); err != nil {
			return err
		}
		g.nInv++
	}
	return nil
}

// loadExistingDimensions 从现有 CSV 载入产品/门店/客户维度.
// 返回值 error, 出错时非 nil.
func (g *Generator) loadExistingDimensions() error {
	if err := g.loadProducts(); err != nil {
		return err
	}
	if err := g.loadStores(); err != nil {
		return err
	}
	return g.loadCustomers()
}

// loadProducts 读取 T00 产品表.
func (g *Generator) loadProducts() error {
	return forEachRow(filepath.Join(g.cfg.OutputDir, model.FileProduct), func(r []string) {
		g.products = append(g.products, product{
			id: atoiSafe(r[0]), code: r[1], category: r[2], name: r[3],
			salePrice: atofSafe(r[4]), costPrice: atofSafe(r[5]),
		})
	})
}

// loadStores 读取 T01 门店表.
func (g *Generator) loadStores() error {
	return forEachRow(filepath.Join(g.cfg.OutputDir, model.FileStore), func(r []string) {
		s := store{
			id: atoiSafe(r[0]), code: r[1], name: r[2], manager: r[3],
			openDate: parseDate(r[4]), districtID: atoiSafe(r[5]), district: r[6],
			lat: atofSafe(r[7]), lng: atofSafe(r[8]),
		}
		if strings.TrimSpace(r[9]) != "" {
			cd := parseDate(r[9])
			s.closeDate = &cd
		}
		g.stores = append(g.stores, s)
	})
}

// loadCustomers 读取 T02 客户表.
func (g *Generator) loadCustomers() error {
	return forEachRow(filepath.Join(g.cfg.OutputDir, model.FileCustomer), func(r []string) {
		g.customers = append(g.customers, customer{
			id: atoiSafe(r[0]), code: r[1], name: r[2], birth: parseDate(r[3]),
			gender: r[4], regDate: parseDate(r[5]), industry: r[6], profession: r[7],
		})
	})
}

// scanExistingOrders 扫描 T04 订单主表, 求最大订单序号与下单日期区间.
// 返回值 int, 最大订单序号; time.Time, 最早下单日期; time.Time, 最晚下单日期; error, 出错时非 nil.
func (g *Generator) scanExistingOrders() (int, time.Time, time.Time, error) {
	var ocMax int
	var minDate, maxDate time.Time
	err := forEachRow(filepath.Join(g.cfg.OutputDir, model.FileOrder), func(r []string) {
		if n := ocSuffix(r[0]); n > ocMax {
			ocMax = n
		}
		d := parseDate(r[2])
		if minDate.IsZero() || d.Before(minDate) {
			minDate = d
		}
		if maxDate.IsZero() || d.After(maxDate) {
			maxDate = d
		}
	})
	return ocMax, minDate, maxDate, err
}

// forEachRow 逐行读取 CSV (跳过表头) 并回调.
//   - path, 文件路径; fn, 行回调.
//
// 返回值 error, 出错时非 nil.
func forEachRow(path string, fn func(r []string)) error {
	f, err := os.Open(path) // #nosec G304 路径来自受控配置目录
	if err != nil {
		return err
	}
	defer util.CloseQuietly(f)

	br := bufio.NewReader(f)
	// 跳过可能存在的 UTF-8 BOM.
	if bs, perr := br.Peek(3); perr == nil && len(bs) == 3 && bs[0] == 0xEF && bs[1] == 0xBB && bs[2] == 0xBF {
		if _, derr := br.Discard(3); derr != nil {
			return derr
		}
	}
	cr := csv.NewReader(br)
	cr.FieldsPerRecord = -1
	first := true
	for {
		rec, err := cr.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if first {
			first = false
			continue
		}
		fn(rec)
	}
	return nil
}

// ocSuffix 从订单编号 OC_0000001 中提取数字序号.
//   - code, 订单编号.
//
// 返回值 int, 序号, 解析失败为 0.
func ocSuffix(code string) int {
	if i := strings.LastIndex(code, "_"); i >= 0 {
		n, _ := strconv.Atoi(code[i+1:]) //nolint:errcheck // 宽松解析, 失败取 0
		return n
	}
	return 0
}

// parseDate 解析 YYYY-MM-DD 日期, 失败返回零值.
func parseDate(s string) time.Time {
	t, _ := time.Parse(dateLayout, strings.TrimSpace(s)) //nolint:errcheck // 宽松解析, 失败取零值
	return t
}

func atoiSafe(s string) int {
	n, _ := strconv.Atoi(strings.TrimSpace(s)) //nolint:errcheck // 宽松解析, 失败取 0
	return n
}

func atofSafe(s string) float64 {
	f, _ := strconv.ParseFloat(strings.TrimSpace(s), 64) //nolint:errcheck // 宽松解析, 失败取 0
	return f
}
func maxTime(a, b time.Time) time.Time {
	if a.After(b) {
		return a
	}
	return b
}
func minTime(a, b time.Time) time.Time {
	if a.Before(b) {
		return a
	}
	return b
}
