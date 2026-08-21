// FilePath    : internal/ui/errorborder.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 输入控件的红色错误边框叠层, 用于非法值即时提示 (不依赖 Fyne 自带校验的对钩图标).

package ui

import (
	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/theme"
)

// errorBorder 在输入控件之上叠加一个可切换的红色描边矩形.
// 非法时显示错误色边框, 合法时透明, 从而在不引入对钩图标的前提下完成状态校验提示.
type errorBorder struct {
	rect *canvas.Rectangle
	obj  *fyne.Container
}

// newErrorBorder 用给定内容创建错误边框叠层 (初始为合法/透明).
//   - content, 被包裹的输入控件.
//
// 返回值 *errorBorder, 叠层实例.
func newErrorBorder(content fyne.CanvasObject) *errorBorder {
	rect := canvas.NewRectangle(color.Transparent)
	rect.StrokeWidth = 2
	rect.StrokeColor = color.Transparent
	rect.CornerRadius = theme.Size(theme.SizeNameInputRadius)
	return &errorBorder{
		rect: rect,
		obj:  container.NewStack(content, rect),
	}
}

// Object 返回可布局的叠层对象.
// 返回值 fyne.CanvasObject, 叠层对象.
func (b *errorBorder) Object() fyne.CanvasObject {
	return b.obj
}

// setInvalid 切换非法状态: 非法显示错误色边框, 合法恢复透明.
//   - invalid, 是否非法.
func (b *errorBorder) setInvalid(invalid bool) {
	if invalid {
		b.rect.StrokeColor = theme.Color(theme.ColorNameError)
	} else {
		b.rect.StrokeColor = color.Transparent
	}
	b.rect.Refresh()
}
