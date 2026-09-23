// FilePath    : internal/generator/types.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 生成器类型定义与业务系数数组.

// Package generator 移植自原 VBA 数据生成逻辑, 负责按业务规则生成各表数据并输出为 CSV.
// 与原版差异: 随机数序列不同 (逐字节不可复现), 但字段/表结构/行数规则/分布规律保持一致;
// 时间窗口由原来的 Now()-1500 基线参数化为用户指定的 [StartDate, EndDate].
package generator

import (
	"time"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/data"
	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// product 产品表中间结构 (T00).
type product struct {
	id        int
	code      string
	category  string
	name      string
	salePrice float64
	costPrice float64
}

// store 门店表中间结构 (T01).
type store struct {
	id        int
	code      string
	name      string
	manager   string
	openDate  time.Time
	cityID    int
	city      string
	lat       float64
	lng       float64
	closeDate *time.Time
}

// customer 客户表中间结构 (T02).
type customer struct {
	id         int
	code       string
	name       string
	birth      time.Time
	gender     string
	regDate    time.Time
	industry   string
	profession string
}

// ProgressFunc 进度回调.
//   - percent, 当前总进度百分比 [0, 100].
//   - stage, 当前阶段描述文案键.
type ProgressFunc func(percent float64, stage string)

// Generator 承载一次生成任务的全部状态.
type Generator struct {
	cfg      *config.Config
	ds       *data.Dataset
	rnd      *util.Rand
	progress ProgressFunc

	products  []product
	stores    []store
	customers []customer

	// 供 T06 汇总: 省 -> 去年全年销售额 / 去年 Q4 销售额.
	provinceFull map[int]float64
	provinceQ4   map[int]float64

	windowDays int // 时间窗口天数, 对应原逻辑中的 1500.
	endYear    int // 结束日期所在年份, 对应原逻辑 Now() 的年份.

	// 生成行数统计.
	nOrders int
	nItems  int
	nInv    int

	// 各表 F_00_自动编号 当前值 (全量从 0 起, 增量从已有最大值续起), 对应原 Access 的 IDENTITY 字段.
	idOrder int // T04 订单主表
	idItem  int // T05 订单子表
	idInv   int // T03 入库信息表
}

// Result 生成结果统计, 各表实际行数.
type Result struct {
	Products  int `json:"products"`
	Stores    int `json:"stores"`
	Customers int `json:"customers"`
	Inventory int `json:"inventory"`
	Orders    int `json:"orders"`
	OrderItem int `json:"orderItem"`
}

// New 创建生成器.
//   - cfg, 生成配置 (须已校验).
//   - ds, 基础数据集.
//   - progress, 进度回调, 可为 nil.
//
// 返回值 *Generator, 生成器实例.
func New(cfg *config.Config, ds *data.Dataset, progress ProgressFunc) *Generator {
	if progress == nil {
		progress = func(float64, string) {}
	}
	return &Generator{
		cfg:          cfg,
		ds:           ds,
		rnd:          util.NewRand(),
		progress:     progress,
		provinceFull: make(map[int]float64),
		provinceQ4:   make(map[int]float64),
		windowDays:   int(cfg.EndDate.Sub(cfg.StartDate).Hours() / 24),
		endYear:      cfg.EndDate.Year(),
	}
}

// - - - 业务系数数组, 与原 VBA 完全一致 - - -

// ageDist 年龄分布 ArrSjNL, count=12.
var ageDist = [12]float64{0, 0.1, 0.2, 0.3, 0.3, 0.3, 0.3, 0.8, 0.8, 0.8, 0.9, 1}

// industries 客户行业 ArrHY.
var industries = [7]string{"建筑业", "制造业", "互联网", "农业", "餐饮", "物流", "汽车"}

// professions 客户职业 ArrZY.
var professions = [7]string{"个体户", "HR", "运营", "IT", "财务", "销售", "研发"}

// industryDist 行业分布 ArrSjHY.
var industryDist = [7]float64{0.2, 0.5, 0.5, 0.8, 0.8, 1, 0.9}

// professionDist 职业分布 ArrSjZY.
var professionDist = [7]float64{0.3, 0.7, 0.6, 1, 1, 0.8, 0.1}

// unitCountFactor 单均件数系数 ArrDjjsxs, count=5.
var unitCountFactor = [5]float64{0.7, 0.8, 1, 1.2, 1.3}

// monthTrend 行业淡旺季趋势 ArrDdslxsMonth, count=12.
var monthTrend = [12]float64{1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1}

// regionOrderFactor 区域订单系数 (近似正态) ArrDdslxsSC, count=34.
var regionOrderFactor = [34]float64{
	0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4,
	1.4, 1.35, 1.3, 1.25, 1.2, 1.15, 1.1, 1.05, 1, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6,
}

// discounts 折扣信息分布 ArrZK, count=6.
var discounts = [6]float64{1, 0.9, 0.8, 0.7, 0.6, 0.5}

// discountMonth 折扣月份分布 ArrZKMonth, count=12.
var discountMonth = [12]float64{0.95, 0.9, 1, 0.98, 0.85, 1, 0.98, 0.88, 0.8, 0.86, 0.92, 0.98}

// customerDist 客户分布 ArrSjKF, count=12.
var customerDist = [12]float64{0, 0.1, 0.5, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.8, 0.9, 1}

// discountIndustries 参与折扣的行业集合 (原 ArrHY 折扣准备).
var discountIndustries = map[string]struct{}{"保险": {}, "互联网": {}, "汽车": {}, "制造业": {}}

// discountProfessions 参与折扣的职业集合 (原 ArrZY 折扣准备).
var discountProfessions = map[string]struct{}{"HR": {}, "财务": {}, "销售": {}, "运营": {}}
