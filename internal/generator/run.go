// FilePath    : internal/generator/run.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 全量生成编排与进度上报.

package generator

import (
	"jiaopengzi/Power-BI-custom-sample-data/internal/csvw"
	"jiaopengzi/Power-BI-custom-sample-data/internal/model"
	"jiaopengzi/Power-BI-custom-sample-data/internal/util"
)

// openOrderWriters 创建 T03/T04/T05 三张表的 CSV 写入器并写入表头.
//   - dir, 目标目录.
//
// 返回值 *orderWriters, 写入器集合; error, 出错时非 nil.
func openOrderWriters(dir string) (*orderWriters, error) {
	inv, err := csvw.Create(dir, model.FileInventory, model.HeaderInventory)
	if err != nil {
		return nil, err
	}
	order, err := csvw.Create(dir, model.FileOrder, model.HeaderOrder)
	if err != nil {
		util.CloseQuietly(inv)
		return nil, err
	}
	item, err := csvw.Create(dir, model.FileOrderItem, model.HeaderOrderItem)
	if err != nil {
		util.CloseQuietly(inv)
		util.CloseQuietly(order)
		return nil, err
	}
	return &orderWriters{inv: inv, order: order, item: item}, nil
}

// GenerateAll 执行全量数据生成, 依次生成 D00-D03, T00-T06 并写出 CSV.
// 返回值 Result, 各表行数统计; error, 出错时非 nil.
func (g *Generator) GenerateAll() (Result, error) {
	g.progress(0, "stageStart")

	// 维度表 D00-D03 (D03 区县表数据量较大, 占用较多进度).
	if err := g.writeDimensions(); err != nil {
		return Result{}, err
	}
	g.progress(30, "stageDimensions")

	// T00 产品表
	g.genProducts()
	if err := g.writeProducts(); err != nil {
		return Result{}, err
	}
	g.progress(35, "stageProducts")

	// T01 门店表
	g.genStores()
	if err := g.writeStores(); err != nil {
		return Result{}, err
	}
	g.progress(45, "stageStores")

	// T02 客户表
	g.genCustomers()
	if err := g.writeCustomers(); err != nil {
		return Result{}, err
	}
	g.progress(55, "stageCustomers")

	// T03/T04/T05 入库/订单主/订单子 (流式写出)
	w, err := openOrderWriters(g.cfg.OutputDir)
	if err != nil {
		return Result{}, err
	}
	if _, err = g.genOrders(w, 0, 55, 90); err != nil {
		util.SilentClose(w.close)
		return Result{}, err
	}
	if err = w.close(); err != nil {
		return Result{}, err
	}
	g.progress(90, "stageOrders")

	// T06 销售目标表
	if err = g.genSaleTargets(); err != nil {
		return Result{}, err
	}
	g.progress(100, "stageDone")

	return Result{
		Products:  len(g.products),
		Stores:    len(g.stores),
		Customers: len(g.customers),
		Inventory: g.nInv,
		Orders:    g.nOrders,
		OrderItem: g.nItems,
	}, nil
}
