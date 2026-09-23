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

// monthKey 省 x 年 x 月的复合键, 供 T06 按月汇总实绩使用.
type monthKey struct {
	pid   int
	year  int
	month int
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

	// 供 T06 汇总: (省, 年, 月) -> 实际销售额; factMinYM/factMaxYM 为事实月份区间 (YYYYMM, 含).
	provinceMonth map[monthKey]float64
	factMinYM     int
	factMaxYM     int

	windowDays int // 时间窗口天数, 对应原逻辑中的 1500.

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
		cfg:           cfg,
		ds:            ds,
		rnd:           util.NewRand(),
		progress:      progress,
		provinceMonth: make(map[monthKey]float64),
		windowDays:    int(cfg.EndDate.Sub(cfg.StartDate).Hours() / 24),
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

// - - - 以下为新增随机因子 (原 VBA 没有的增强), 用于打破时间与产品维度上 "几乎一致" 的形态;
// 与既有系数数组方案保持同一风格: 固定数组 + 确定性索引 (按年份/日期/ID 取模),
// 均值近似 1 以维持整体量级不变, 且同一日期/ID 在全量与增量两次生成中取值一致 - - -

// yearTrend 年景系数 (宏观景气周期), 按年份取模索引, count=5.
// 使不同年份的订单量存在整体高低波动, 打破逐年一致.
var yearTrend = [5]float64{0.88, 0.94, 1, 1.07, 1.14}

// monthNoise 月度随机噪声, 按 年*12+月 取模索引, count=37 (与 12 互质, 约 3 年不重复).
// 在行业淡旺季趋势之上叠加逐年不同的月度扰动, 打破 "每年淡旺季形态完全一致".
var monthNoise = [37]float64{
	0.97, 1.03, 0.99, 1.05, 0.95, 1.02, 0.98, 1.06, 0.93, 1.01,
	0.96, 1.04, 0.99, 1.02, 0.94, 1.05, 0.98, 1.03, 0.97, 1.01,
	0.95, 1.06, 0.99, 1.02, 0.96, 1.04, 0.98, 1.03, 0.94, 1.05,
	0.97, 1.01, 0.99, 1.04, 0.96, 1.02, 1,
}

// weekdayFactor 星期系数 (周日为下标 0), count=7. 周末与周五客流更高, 工作日偏低的客观规律.
var weekdayFactor = [7]float64{1.05, 0.9, 0.9, 0.9, 0.95, 1.1, 1.2}

// storeTrafficFactor 门店客流系数 (近似正态的钟形分布), 按门店 ID 取模索引, count=21.
// 使同类区域门店的日常客流天生存在高低差异.
var storeTrafficFactor = [21]float64{
	0.75, 0.8, 0.85, 0.9, 0.96, 1.01, 1.07, 1.12, 1.17, 1.23,
	1.28, 1.23, 1.17, 1.12, 1.07, 1.01, 0.96, 0.9, 0.85, 0.8, 0.75,
}

// productPopularity 产品人气系数 (近似正态的钟形分布), 按产品索引取模, count=29 (与既有 %5/%8 错开).
// 使不同产品的销售数量天生存在畅销/滞销差异.
var productPopularity = [29]float64{
	0.7, 0.74, 0.78, 0.82, 0.86, 0.9, 0.94, 0.98, 1.02, 1.06,
	1.1, 1.14, 1.18, 1.22, 1.26, 1.22, 1.18, 1.14, 1.1, 1.06,
	1.02, 0.98, 0.94, 0.9, 0.86, 0.82, 0.78, 0.74, 0.7,
}

// targetYearBias 销售目标年度偏置, 按年份取模索引, count=5.
// 使不同年度的目标整体偏乐观 (正值, 完成率低于 100%) 或偏保守 (负值, 完成率高于 100%),
// 年度完成率因此围绕 100% 上下分化, 更贴近真实业务的年度节奏.
var targetYearBias = [5]float64{-0.05, 0.06, -0.02, 0.04, -0.04}
