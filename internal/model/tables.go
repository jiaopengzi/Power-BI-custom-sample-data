// FilePath    : internal/model/tables.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 各数据表的文件名与表头定义.

// Package model 定义各数据表的文件名与表头, 与原 VBA/Access 的字段完全对应.
package model

// 表文件名 (CSV), 与原 Access 表名一致.
const (
	FileProduct    = "T00_产品表.csv"
	FileStore      = "T01_门店表.csv"
	FileCustomer   = "T02_客户表.csv"
	FileInventory  = "T03_入库信息表.csv"
	FileOrder      = "T04_订单主表.csv"
	FileOrderItem  = "T05_订单子表.csv"
	FileSaleTarget = "T06_销售目标表.csv"
	FileRegion     = "D00_大区表.csv"
	FileProvince   = "D01_省份表.csv"
	FileCity       = "D02_城市表.csv"
	FileDistrict   = "D03_区县表.csv"
)

// colAutoID 各表统一的自增主键列名 (原 Access IDENTITY 字段).
const colAutoID = "F_00_自动编号"

// 各表表头, 保留原始字段名 (含 F_00_自动编号 等) 以保证与原产物字段一致.
var (
	HeaderProduct    = []string{colAutoID, "F_01_产品编号", "F_02_产品分类", "F_03_产品名称", "F_04_产品销售价格", "F_05_产品成本价格"}
	HeaderStore      = []string{colAutoID, "F_01_门店编号", "F_02_门店名称", "F_03_门店负责人", "F_04_开店日期", "F_05_城市ID", "F_06_城市", "F_07_纬度", "F_08_经度", "F_09_关店日期"}
	HeaderCustomer   = []string{colAutoID, "F_01_客户编号", "F_02_客户名称", "F_03_客户生日", "F_04_客户性别", "F_05_注册日期", "F_06_客户行业", "F_07_客户职业"}
	HeaderInventory  = []string{colAutoID, "F_01_入库产品编号", "F_02_入库产品数量", "F_03_入库门店编号", "F_04_入库日期"}
	HeaderOrder      = []string{colAutoID, "F_01_订单编号", "F_02_门店编号", "F_03_下单日期", "F_04_送货日期", "F_05_客户编号", "F_06_销售渠道"}
	HeaderOrderItem  = []string{colAutoID, "F_01_订单编号", "F_02_产品编号", "F_03_产品销售价格", "F_04_折扣比例", "F_05_产品销售数量", "F_06_产品销售金额"}
	HeaderSaleTarget = []string{colAutoID, "F_01_省ID", "F_02_省简称", "F_03_月份", "F_04_销售目标"}
	HeaderRegion     = []string{colAutoID, "F_01_大区ID", "F_02_大区", "F_03_大区负责人", "F_04_办公地城市ID", "F_05_办公地城市", "F_06_纬度", "F_07_经度"}
	HeaderProvince   = []string{colAutoID, "F_01_大区ID", "F_02_省ID", "F_03_省全称", "F_04_省简称1", "F_05_省简称2", "F_06_纬度", "F_07_经度"}
	HeaderCity       = []string{colAutoID, "F_01_省ID", "F_02_城市ID", "F_03_城市", "F_04_纬度", "F_05_经度"}
	HeaderDistrict   = []string{colAutoID, "F_01_城市ID", "F_02_区县ID", "F_03_区县", "F_04_纬度", "F_05_经度"}

	// AllFiles 全部产物文件名, 用于增量更新前的存在性检查.
	AllFiles = []string{
		FileProduct, FileStore, FileCustomer, FileInventory, FileOrder,
		FileOrderItem, FileSaleTarget, FileRegion, FileProvince, FileCity, FileDistrict,
	}
)
