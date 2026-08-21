// FilePath    : internal/ui/numberfield.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 基于 fyne-x NumericalEntry 的整数输入控件 (min/max 夹取 + ▲▼ 步进).

package ui

import (
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	xwidget "fyne.io/x/fyne/widget"
)

// NumberField 整数输入控件: 复用 fyne-x NumericalEntry 的数字过滤,
// 叠加 [min, max] 夹取与上下步进按钮, 避免自造数字校验/步进逻辑.
type NumberField struct {
	entry    *xwidget.NumericalEntry
	up       *widget.Button
	down     *widget.Button
	obj      fyne.CanvasObject
	min      int
	max      int
	step     int
	value    int
	updating bool
	// OnChanged 值变化回调 (夹取后的整数), 可为 nil.
	OnChanged func(int)
}

// NewNumberField 创建整数输入控件.
//   - minVal, maxVal, 取值区间 (含端点).
//   - step, 步进量.
//   - initial, 初始值 (超出区间会被夹取).
//
// 返回值 *NumberField, 控件实例.
func NewNumberField(minVal, maxVal, step, initial int) *NumberField {
	entry := xwidget.NewNumericalEntry()
	nf := &NumberField{
		entry: entry,
		min:   minVal,
		max:   maxVal,
		step:  step,
		value: clampInt(initial, minVal, maxVal),
	}
	entry.SetText(strconv.Itoa(nf.value))
	entry.OnChanged = nf.onEntryChanged

	nf.up = widget.NewButtonWithIcon("", theme.MenuDropUpIcon(), func() { nf.add(nf.step) })
	nf.down = widget.NewButtonWithIcon("", theme.MenuDropDownIcon(), func() { nf.add(-nf.step) })
	nf.up.Importance = widget.LowImportance
	nf.down.Importance = widget.LowImportance

	steppers := container.NewGridWithRows(2, nf.up, nf.down)
	nf.obj = container.NewBorder(nil, nil, nil, steppers, entry)
	return nf
}

// Object 返回可布局的控件对象.
// 返回值 fyne.CanvasObject, 控件对象.
func (nf *NumberField) Object() fyne.CanvasObject {
	return nf.obj
}

// Value 返回当前夹取后的整数值.
// 返回值 int, 当前值.
func (nf *NumberField) Value() int {
	if n, ok := parseIntLoose(nf.entry.Text); ok {
		nf.value = clampInt(n, nf.min, nf.max)
	}
	return nf.value
}

// SetDisabled 启用或禁用输入与步进按钮.
//   - disabled, 为 true 时禁用.
func (nf *NumberField) SetDisabled(disabled bool) {
	if disabled {
		nf.entry.Disable()
		nf.up.Disable()
		nf.down.Disable()
		return
	}
	nf.entry.Enable()
	nf.up.Enable()
	nf.down.Enable()
}

// onEntryChanged 文本变化时解析并夹取, 越界时回写规范值.
//   - s, 当前文本.
func (nf *NumberField) onEntryChanged(s string) {
	if nf.updating {
		return
	}
	n, ok := parseIntLoose(s)
	if !ok {
		return
	}
	c := clampInt(n, nf.min, nf.max)
	nf.value = c
	if c != n {
		nf.setText(strconv.Itoa(c))
	}
	if nf.OnChanged != nil {
		nf.OnChanged(c)
	}
}

// add 在当前值基础上增减 delta 并夹取.
//   - delta, 增减量.
func (nf *NumberField) add(delta int) {
	nf.value = clampInt(nf.value+delta, nf.min, nf.max)
	nf.setText(strconv.Itoa(nf.value))
	if nf.OnChanged != nil {
		nf.OnChanged(nf.value)
	}
}

// setText 以防重入的方式设置文本.
//   - s, 目标文本.
func (nf *NumberField) setText(s string) {
	nf.updating = true
	nf.entry.SetText(s)
	nf.updating = false
}

// clampInt 将 v 夹取到 [lo, hi].
//   - v, 原值; lo, hi, 区间端点.
//
// 返回值 int, 夹取结果.
func clampInt(v, lo, hi int) int {
	if v < lo {
		return lo
	}
	if v > hi {
		return hi
	}
	return v
}

// parseIntLoose 宽松解析: 仅保留数字字符后转为 int (容忍千分位分隔符等).
//   - s, 原始文本.
//
// 返回值 int, 解析值; bool, 是否成功.
func parseIntLoose(s string) (int, bool) {
	digits := make([]rune, 0, len(s))
	for _, r := range s {
		if r >= '0' && r <= '9' {
			digits = append(digits, r)
		}
	}
	if len(digits) == 0 {
		return 0, false
	}
	n, err := strconv.Atoi(string(digits))
	if err != nil {
		return 0, false
	}
	return n, true
}
