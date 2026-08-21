// FilePath    : internal/config/config.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 应用配置, 全局参数 (N0/N1/N3) 与校验规则.

// Package config 定义应用配置, 全局参数与校验规则.
// 参数含义与取值范围继承自原 VBA 版本: N0 产品数量, N1 门店数量, N3 入库间隔最大数.
package config

import (
	"fmt"
	"time"
)

// Locale 界面与产物语言标识.
type Locale string

const (
	// LocaleZhCN 简体中文 (默认).
	LocaleZhCN Locale = "zh-cn"
	// LocaleEnUS 英文.
	LocaleEnUS Locale = "en-us"
)

// 参数取值范围, 与原 VBA 表单 Form_main 的校验保持一致.
const (
	MinProductCount = 1
	MaxProductCount = 10000
	MinStoreCount   = 1
	MaxStoreCount   = 10000
	MinInventory    = 5
	MaxInventory    = 180
)

// DateLayout 前后端交互与 CSV 使用的日期格式.
const DateLayout = "2006-01-02"

// Config 生成任务的完整配置.
//   - OutputDir, CSV 产物存放目录.
//   - Locale, 界面/产物语言.
//   - ProductCount, 产品数量 (原 N0).
//   - StoreCount, 门店数量 (原 N1).
//   - InventoryCycle, 入库间隔最大数 (原 N3).
//   - StartDate, 数据时间窗口起点 (替代原 Now()-1500 基线).
//   - EndDate, 数据时间窗口终点 (替代原 Now()).
type Config struct {
	OutputDir      string    `json:"outputDir"`
	Locale         Locale    `json:"locale"`
	ProductCount   int       `json:"productCount"`
	StoreCount     int       `json:"storeCount"`
	InventoryCycle int       `json:"inventoryCycle"`
	StartDate      time.Time `json:"-"`
	EndDate        time.Time `json:"-"`
}

// Validate 校验配置参数是否在允许范围内.
// 返回值 error, 参数非法时非 nil.
func (c *Config) Validate() error {
	if c.OutputDir == "" {
		return fmt.Errorf("output directory is empty")
	}
	if c.ProductCount < MinProductCount || c.ProductCount > MaxProductCount {
		return fmt.Errorf("product count must be in [%d, %d]", MinProductCount, MaxProductCount)
	}
	if c.StoreCount < MinStoreCount || c.StoreCount > MaxStoreCount {
		return fmt.Errorf("store count must be in [%d, %d]", MinStoreCount, MaxStoreCount)
	}
	if c.InventoryCycle < MinInventory || c.InventoryCycle > MaxInventory {
		return fmt.Errorf("inventory cycle must be in [%d, %d]", MinInventory, MaxInventory)
	}
	if !c.EndDate.After(c.StartDate) {
		return fmt.Errorf("end date must be after start date")
	}
	// 窗口至少需覆盖门店最小营业期 (原逻辑保证开店距结束至少 28 天)
	if c.EndDate.Sub(c.StartDate) < 60*24*time.Hour {
		return fmt.Errorf("date window is too small, need at least 60 days")
	}
	return nil
}
