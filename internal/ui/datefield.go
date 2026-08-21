// FilePath    : internal/ui/datefield.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : YYYY-MM-DD 日期输入控件, 文本框 + 日历弹窗 (年切换/快捷选择) + 非法红框与本地化提示.

package ui

import (
	"strconv"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/i18n"
)

// DateField 日期输入控件: 文本显示 YYYY-MM-DD, 旁置按钮弹出日历选择;
// 非法日期时输入框红框 + 下方本地化文字提示.
type DateField struct {
	entry     *scrollEntry
	border    *errorBorder
	hint      *widget.Label
	button    *widget.Button
	obj       fyne.CanvasObject
	win       fyne.Window
	mgr       *i18n.Manager
	titleKey  string
	extErrKey string // 跨字段校验的错误文案键, 空表示无
	// OnChanged 日期文本变化时回调 (供表单重新校验), 可为 nil.
	OnChanged func()
}

// NewDateField 创建日期输入控件.
//   - win, 承载日历弹窗的父窗口.
//   - mgr, 文案管理器 (弹窗标题/关闭按钮随语言切换).
//   - titleKey, 弹窗标题的文案键.
//   - initial, 初始日期.
//
// 返回值 *DateField, 控件实例.
func NewDateField(win fyne.Window, mgr *i18n.Manager, titleKey string, initial time.Time) *DateField {
	entry := newScrollEntry()
	entry.SetText(initial.Format(config.DateLayout))
	df := &DateField{entry: entry, win: win, mgr: mgr, titleKey: titleKey}
	df.button = widget.NewButtonWithIcon("", calendarIcon(), df.showCalendar)
	df.border = newErrorBorder(entry)

	df.hint = widget.NewLabel("")
	df.hint.Importance = widget.DangerImportance
	df.hint.Hide()

	entry.OnChanged = df.onChanged
	row := container.NewBorder(nil, nil, nil, df.button, df.border.Object())
	df.obj = container.NewVBox(row, df.hint)
	return df
}

// onChanged 文本变化时刷新校验状态并回调表单.
//   - s, 当前文本.
func (df *DateField) onChanged(s string) {
	df.refreshState()
	if df.OnChanged != nil {
		df.OnChanged()
	}
}

// SetExternalError 设置或清除来自跨字段校验的错误文案键 (空字符串表示清除).
//   - key, 错误文案键.
func (df *DateField) SetExternalError(key string) {
	df.extErrKey = key
	df.refreshState()
}

// refreshState 根据格式合法性与外部错误刷新红框与提示 (格式错优先于跨字段错).
func (df *DateField) refreshState() {
	key := ""
	switch {
	case !df.Valid():
		key = "form.dateInvalid"
	case df.extErrKey != "":
		key = df.extErrKey
	}
	if key == "" {
		df.border.setInvalid(false)
		df.hint.Hide()
		return
	}
	df.border.setInvalid(true)
	df.hint.SetText(df.mgr.T(key))
	df.hint.Show()
}

// Valid 返回当前文本是否为合法的 YYYY-MM-DD 日期.
// 返回值 bool, 合法返回 true.
func (df *DateField) Valid() bool {
	_, err := time.Parse(config.DateLayout, df.entry.Text)
	return err == nil
}

// RefreshLocale 语言切换时刷新提示文案.
func (df *DateField) RefreshLocale() {
	df.refreshState()
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

// currentDate 返回输入框当前日期, 解析失败时回退为今天.
// 返回值 time.Time, 当前日期.
func (df *DateField) currentDate() time.Time {
	if t, err := time.Parse(config.DateLayout, df.entry.Text); err == nil {
		return t
	}
	return time.Now()
}

// showCalendar 弹出日历: 顶部年份快速切换, 中部日历, 底部今日/年初/月初/月末/年末快捷选择.
func (df *DateField) showCalendar() {
	view := df.currentDate()
	content := container.NewVBox()
	var popup *dialog.CustomDialog

	pick := func(t time.Time) {
		df.entry.SetText(t.Format(config.DateLayout))
		if popup != nil {
			popup.Hide()
		}
	}

	var rebuild func()
	rebuild = func() {
		yearLabel := widget.NewLabelWithStyle(strconv.Itoa(view.Year()), fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
		prevYear := widget.NewButtonWithIcon("", theme.NavigateBackIcon(), func() { view = view.AddDate(-1, 0, 0); rebuild() })
		nextYear := widget.NewButtonWithIcon("", theme.NavigateNextIcon(), func() { view = view.AddDate(1, 0, 0); rebuild() })
		yearRow := container.NewBorder(nil, nil, prevYear, nextYear, yearLabel)

		cal := widget.NewCalendar(view, pick)

		quick := container.NewGridWithColumns(5,
			widget.NewButton(df.mgr.T("date.today"), func() { pick(time.Now()) }),
			widget.NewButton(df.mgr.T("date.yearStart"), func() { pick(dayIn(view.Year(), time.January, 1)) }),
			widget.NewButton(df.mgr.T("date.monthStart"), func() { pick(dayIn(view.Year(), view.Month(), 1)) }),
			widget.NewButton(df.mgr.T("date.monthEnd"), func() { pick(monthEnd(view)) }),
			widget.NewButton(df.mgr.T("date.yearEnd"), func() { pick(dayIn(view.Year(), time.December, 31)) }),
		)

		content.Objects = []fyne.CanvasObject{yearRow, cal, quick}
		content.Refresh()
	}
	rebuild()

	popup = dialog.NewCustom(df.mgr.T(df.titleKey), df.mgr.T("common.cancel"), content, df.win)
	popup.Show()
}

// dayIn 构造指定年月日的本地日期.
//   - year, month, day, 年/月/日.
//
// 返回值 time.Time, 对应日期.
func dayIn(year int, month time.Month, day int) time.Time {
	return time.Date(year, month, day, 0, 0, 0, 0, time.Local)
}

// monthEnd 返回给定日期所在月份的最后一天.
//   - t, 参考日期.
//
// 返回值 time.Time, 月末日期.
func monthEnd(t time.Time) time.Time {
	return dayIn(t.Year(), t.Month(), 1).AddDate(0, 1, -1)
}
