// FilePath    : internal/generator/dimensions.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 维度表 (D00-D03) 与产品/门店/客户表 (T00-T02) 生成.

package generator

import (
	"strconv"
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/csvw"
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

// genProducts 生成产品表 T00, 对应 DataTableT0.
func (g *Generator) genProducts() {
	n := g.cfg.ProductCount
	g.products = make([]product, 0, n)
	for i := 1; i <= n; i++ {
		sj := g.rnd.F()
		letter := util.Letter(util.RoundInt(sj * 9)) // A-J
		r4 := 5000 + sj*5000
		var r5 float64
		if sj < 0.28 {
			r5 = r4 * 0.18
		} else {
			r5 = r4 * sj
		}
		g.products = append(g.products, product{
			id:        i,
			code:      "SKU_" + util.PadInt(i, 6),
			category:  letter + "类",
			name:      "产品" + letter + util.PadInt(i, 4),
			salePrice: util.RoundBankers(r4, 2),
			costPrice: util.RoundBankers(r5, 2),
		})
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

// fixedStore 原 VBA Arr7 中四直辖市 + 港澳台的固定门店信息.
type fixedStore struct {
	code       string
	manager    string
	districtID int
	district   string
	lat        float64
	lng        float64
}

// arr7 前七个优先命中的固定门店 (与原 VBA Arr7 一致).
var arr7 = []fixedStore{
	{"SC_0001", "焦阿大", 110101, "东城区", 39.917544, 116.418757},
	{"SC_0002", "焦阿二", 120101, "和平区", 39.118328, 121.490318},
	{"SC_0003", "焦阿三", 310101, "黄浦区", 31.222778, 121.471518},
	{"SC_0004", "焦阿四", 500103, "渝中区", 29.556748, 106.562888},
	{"SC_0005", "焦阿五", 710000, "台湾", 25.044518, 121.509518},
	{"SC_0006", "焦阿六", 810001, "中西区", 22.28198088, 114.1543738},
	{"SC_0007", "焦阿七", 820001, "花地玛堂区", 22.207878, 113.5528958},
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

// randOpenDate 生成随机开店日期, 保证距结束日期至少 28 天 (对应原 Now-Round(Rnd*1500+28)).
// 返回值 time.Time, 开店日期.
func (g *Generator) randOpenDate() time.Time {
	span := g.windowDays - 28
	if span < 1 {
		span = 1
	}
	return addDays(g.cfg.EndDate, -(util.RoundInt(g.rnd.F()*float64(span)) + 28))
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

// genStores 生成门店表 T01, 对应 DataTableT1.
func (g *Generator) genStores() {
	n := g.cfg.StoreCount
	names := g.randStoreNames(n)
	g.stores = make([]store, 0, n)

	if n < 8 {
		for k := 0; k < n; k++ {
			a := arr7[k]
			g.stores = append(g.stores, store{
				id: k + 1, code: a.code, name: names[k], manager: a.manager,
				openDate: g.randOpenDate(), districtID: a.districtID, district: a.district,
				lat: a.lat, lng: a.lng,
			})
		}
		return
	}

	// N1 > 7: 先写 7 个固定门店
	for k := 0; k < 7; k++ {
		a := arr7[k]
		g.stores = append(g.stores, store{
			id: k + 1, code: a.code, name: names[k], manager: a.manager,
			openDate: g.randOpenDate(), districtID: a.districtID, district: a.district,
			lat: a.lat, lng: a.lng,
		})
	}
	// 再生成剩余随机门店 (原 For i = 8 To N1)
	for i := 8; i <= n; i++ {
		sj := g.rnd.F()
		d := g.ds.Districts[util.RoundInt(float64(len(g.ds.Districts)-1)*g.rnd.F())]
		openDate := g.randOpenDate()
		closeCandidate := addDays(openDate, 550+util.RoundInt(4320*g.rnd.F()))
		st := store{
			id: i, code: "SC_" + util.PadInt(i, 4), name: names[i-1],
			manager: g.randName(0.66, sj), openDate: openDate,
			districtID: d.DistrictID, district: d.Name,
			lat: util.RoundBankers(d.Lat+g.rnd.F()*0.05, 6),
			lng: util.RoundBankers(d.Lng+g.rnd.F()*0.05, 6),
		}
		// dateGD > Now 表示尚未关店; 否则记录关店日期.
		if !closeCandidate.After(g.cfg.EndDate) {
			cd := closeCandidate
			st.closeDate = &cd
		}
		g.stores = append(g.stores, st)
	}
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
				strconv.Itoa(s.districtID), s.district, ff(s.lat), ff(s.lng), closeStr}
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
