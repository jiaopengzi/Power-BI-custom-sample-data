// FilePath    : internal/ui/app.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : Fyne 应用引导入口.

// Package ui 使用原生 Go Fyne 实现桌面界面, 复刻原 Wails+Vue3 版本的交互与外观.
// 业务逻辑仍由 internal/app 编排层驱动, 界面层只负责渲染与事件.
package ui

import (
	fyneapp "fyne.io/fyne/v2/app"
)

// appID 应用唯一标识, 供 Fyne 存储/通知等使用.
const appID = "com.jiaopengzi.pbicsd"

// Run 引导并启动 Fyne 应用: 应用品牌主题与图标, 显示主窗口, 进入事件循环.
func Run() {
	a := fyneapp.NewWithID(appID)
	th := newBrandTheme()
	a.Settings().SetTheme(th)
	a.SetIcon(appIcon())

	w := newMainWindow(a)
	w.win.ShowAndRun()
}
