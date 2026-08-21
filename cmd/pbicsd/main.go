// FilePath    : cmd/pbicsd/main.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 可执行入口, 引导 Fyne 界面.

// Command pbicsd 是 Power BI Custom Sample Data 的桌面程序入口:
// 仅负责启动 internal/ui 的 Fyne 界面, 全部业务逻辑位于 internal 各包.
package main

import "jiaopengzi/Power-BI-custom-sample-data/internal/ui"

// main 程序入口.
func main() {
	ui.Run()
}
