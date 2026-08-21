// FilePath    : internal/ui/scrollentry.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 关闭内部文本滚动的单行输入框, 使鼠标滚轮穿透给外层页面滚动容器.

package ui

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/widget"
)

// scrollEntry 单行输入框, 关闭内部文本滚动 (Scroll=None), 使悬停时滚轮直接作用于页面滚动容器.
// Fyne 单行 Entry 内部自带文本滚动容器会吞掉滚轮, 关掉它即可让页面滚动生效.
type scrollEntry struct {
	widget.Entry
}

// newScrollEntry 创建关闭内部滚动的输入框.
// 返回值 *scrollEntry, 输入框实例.
func newScrollEntry() *scrollEntry {
	e := &scrollEntry{}
	e.Scroll = fyne.ScrollNone
	e.ExtendBaseWidget(e)
	return e
}
