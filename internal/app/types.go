// FilePath    : internal/app/types.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 与 GUI 无关的编排层参数, 响应与结果码定义.

// Package app 提供与具体 GUI 框架无关的应用编排层:
// 参数解析校验, 调用 generator 执行全量生成/增量更新, 并将结果归一为统一响应.
// 该层不依赖 Fyne, 便于单元测试与将来替换界面.
package app

import "jiaopengzi/Power-BI-custom-sample-data/internal/generator"

// Params 界面提交的生成参数, 日期以 YYYY-MM-DD 字符串传递.
//   - OutputDir, CSV 产物存放目录.
//   - Locale, 界面/产物语言 (zh-cn / en-us).
//   - ProductCount, 产品数量 (原 N0).
//   - StoreCount, 门店数量 (原 N1).
//   - InventoryCycle, 入库间隔最大数 (原 N3).
//   - StartDate, 数据时间窗口起点.
//   - EndDate, 数据时间窗口终点.
type Params struct {
	OutputDir      string
	Locale         string
	ProductCount   int
	StoreCount     int
	InventoryCycle int
	StartDate      string
	EndDate        string
}

// Response 统一返回结构.
//   - Code, 结果码: CodeOK / CodeNoBaseData / CodeDateConflict / CodeError.
//   - Message, 附加消息 (如冲突区间描述).
//   - Result, 各表行数统计.
type Response struct {
	Code    string
	Message string
	Result  generator.Result
}

// 响应结果码, 与界面 renderResponse 的分支一一对应.
const (
	// CodeOK 执行成功.
	CodeOK = "ok"
	// CodeError 参数非法或执行失败.
	CodeError = "error"
	// CodeNoBaseData 增量更新时目录缺少基础数据.
	CodeNoBaseData = "no_base_data"
	// CodeDateConflict 增量日期区间与现有数据冲突.
	CodeDateConflict = "date_conflict"
)
