// FilePath    : internal/ui/numberfield.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 仅数字整数输入控件 (min/max 夹取 + −/+ 步进 + 非法值红色边框).

package ui

import (
	"strconv"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
)

// intEntry 仅接受数字的输入框: 不设置 Validator (避免对钩图标), 并复用 scrollEntry 的滚轮转发.
type intEntry struct {
	scrollEntry
}

// newIntEntry 创建仅数字输入框.
// 返回值 *intEntry, 输入框实例.
func newIntEntry() *intEntry {
	e := &intEntry{}
	e.Scroll = fyne.ScrollNone
	e.ExtendBaseWidget(e)
	return e
}

// TypedRune 只放行 0-9 的字符, 过滤其余输入.
//   - r, 输入的字符.
func (e *intEntry) TypedRune(r rune) {
	if r >= '0' && r <= '9' {
		e.Entry.TypedRune(r)
	}
}

// NumberField 整数输入控件: 纯数字输入 + [min, max] 夹取 + 左右 −/+ 步进,
// 非法值 (空或越界) 时以红色边框提示.
type NumberField struct {
	entry    *intEntry
	border   *errorBorder
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
	entry := newIntEntry()
	nf := &NumberField{
		entry: entry,
		min:   minVal,
		max:   maxVal,
		step:  step,
		value: clampInt(initial, minVal, maxVal),
	}
	entry.SetText(strconv.Itoa(nf.value))
	entry.OnChanged = nf.onEntryChanged

	nf.up = widget.NewButtonWithIcon("", theme.ContentAddIcon(), func() { nf.add(nf.step) })
	nf.down = widget.NewButtonWithIcon("", theme.ContentRemoveIcon(), func() { nf.add(-nf.step) })
	nf.up.Importance = widget.LowImportance
	nf.down.Importance = widget.LowImportance

	nf.border = newErrorBorder(entry)
	nf.obj = container.New(tightRowLayout{}, nf.down, nf.border.Object(), nf.up)
	nf.updateSteppers()
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

// Valid 返回当前文本是否为区间内的合法整数.
// 返回值 bool, 合法返回 true.
func (nf *NumberField) Valid() bool {
	n, ok := parseIntLoose(nf.entry.Text)
	return ok && n >= nf.min && n <= nf.max
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
	nf.updateSteppers()
}

// updateSteppers 依据当前值与边界启停加/减按钮 (值达上限禁用加号, 达下限禁用减号).
func (nf *NumberField) updateSteppers() {
	if nf.value >= nf.max {
		nf.up.Disable()
	} else {
		nf.up.Enable()
	}
	if nf.value <= nf.min {
		nf.down.Disable()
	} else {
		nf.down.Enable()
	}
}

// onEntryChanged 文本变化时解析并校验: 空或越界则标红, 否则记录夹取值.
//   - s, 当前文本.
func (nf *NumberField) onEntryChanged(s string) {
	if nf.updating {
		return
	}
	n, ok := parseIntLoose(s)
	nf.border.setInvalid(!ok || n < nf.min || n > nf.max)
	if ok {
		nf.value = clampInt(n, nf.min, nf.max)
	}
	nf.updateSteppers()
	if nf.OnChanged != nil {
		nf.OnChanged(nf.value)
	}
}

// add 在当前值基础上增减 delta 并夹取 (步进后必为合法值).
//   - delta, 增减量.
func (nf *NumberField) add(delta int) {
	nf.value = clampInt(nf.value+delta, nf.min, nf.max)
	nf.setText(strconv.Itoa(nf.value))
	nf.border.setInvalid(false)
	nf.updateSteppers()
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

// tightRowLayout 横向排列 [左, 中, 右] 三个元素: 两端按各自最小宽度贴边,
// 中间拉伸填满剩余宽度, 元素之间零间距, 使加减按钮与输入框贴合成一个整体.
type tightRowLayout struct{}

// MinSize 返回三个子元素横向堆叠所需的最小尺寸.
//   - objs, 子元素 (需为 [左, 中, 右] 三个).
//
// 返回值 fyne.Size, 最小尺寸.
func (tightRowLayout) MinSize(objs []fyne.CanvasObject) fyne.Size {
	var width, height float32
	for _, o := range objs {
		m := o.MinSize()
		width += m.Width
		if m.Height > height {
			height = m.Height
		}
	}
	return fyne.NewSize(width, height)
}

// Layout 将左右元素按最小宽度贴边, 中间元素占据剩余宽度, 全部等高.
//   - objs, 子元素 (需为 [左, 中, 右] 三个).
//   - size, 容器尺寸.
func (tightRowLayout) Layout(objs []fyne.CanvasObject, size fyne.Size) {
	if len(objs) != 3 {
		return
	}
	left, center, right := objs[0], objs[1], objs[2]
	lw := left.MinSize().Width
	rw := right.MinSize().Width
	left.Resize(fyne.NewSize(lw, size.Height))
	left.Move(fyne.NewPos(0, 0))
	center.Resize(fyne.NewSize(size.Width-lw-rw, size.Height))
	center.Move(fyne.NewPos(lw, 0))
	right.Resize(fyne.NewSize(rw, size.Height))
	right.Move(fyne.NewPos(size.Width-rw, 0))
}
