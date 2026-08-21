// FilePath    : internal/ui/datefield.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : YYYY-MM-DD 日期输入控件, 文本框 + 核心 Calendar 弹窗选择.

package ui

import (
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/i18n"
)

// DateField 日期输入控件: 只读文本显示 YYYY-MM-DD, 旁置按钮弹出日历选择.
type DateField struct {
	entry    *widget.Entry
	button   *widget.Button
	obj      fyne.CanvasObject
	win      fyne.Window
	mgr      *i18n.Manager
	titleKey string
}

// NewDateField 创建日期输入控件.
//   - win, 承载日历弹窗的父窗口.
//   - mgr, 文案管理器 (弹窗标题/关闭按钮随语言切换).
//   - titleKey, 弹窗标题的文案键.
//   - initial, 初始日期.
//
// 返回值 *DateField, 控件实例.
func NewDateField(win fyne.Window, mgr *i18n.Manager, titleKey string, initial time.Time) *DateField {
	entry := widget.NewEntry()
	entry.SetText(initial.Format(config.DateLayout))
	entry.Validator = dateValidator
	df := &DateField{entry: entry, win: win, mgr: mgr, titleKey: titleKey}
	df.button = widget.NewButtonWithIcon("", theme.MenuDropDownIcon(), df.showCalendar)
	df.obj = container.NewBorder(nil, nil, nil, df.button, entry)
	return df
}

// Object 返回可布局的控件对象.
// 返回值 fyne.CanvasObject, 控件对象.
func (df *DateField) Object() fyne.CanvasObject {
	return df.obj
}

// Text 返回当前日期文本 (YYYY-MM-DD).
// 返回值 string, 日期文本.
func (df *DateField) Text() string {
	return df.entry.Text
}

// SetDisabled 启用或禁用输入与选择按钮.
//   - disabled, 为 true 时禁用.
func (df *DateField) SetDisabled(disabled bool) {
	if disabled {
		df.entry.Disable()
		df.button.Disable()
		return
	}
	df.entry.Enable()
	df.button.Enable()
}

// showCalendar 弹出日历, 选择后写回文本并关闭.
func (df *DateField) showCalendar() {
	cur, err := time.Parse(config.DateLayout, df.entry.Text)
	if err != nil {
		cur = time.Now()
	}
	var popup *dialog.CustomDialog
	cal := widget.NewCalendar(cur, func(t time.Time) {
		df.entry.SetText(t.Format(config.DateLayout))
		if popup != nil {
			popup.Hide()
		}
	})
	popup = dialog.NewCustom(df.mgr.T(df.titleKey), df.mgr.T("common.cancel"), cal, df.win)
	popup.Show()
}

// dateValidator 校验文本是否为合法的 YYYY-MM-DD 日期.
//   - s, 文本.
//
// 返回值 error, 非法时非 nil.
func dateValidator(s string) error {
	_, err := time.Parse(config.DateLayout, s)
	return err
}
