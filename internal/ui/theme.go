// FilePath    : internal/ui/theme.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 品牌主题, 强制浅色 + 藏蓝主色 + 金色超链接 + 内嵌中文字体.

package ui

import (
	"image/color"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/theme"
)

// 品牌主色: 深藏蓝 + 金 (与原前端 --pbi-primary / --pbi-secondary 一致).
var (
	brandNavy = color.NRGBA{R: 0x1e, G: 0x28, B: 0x58, A: 0xff}
	brandGold = color.NRGBA{R: 0xc8, G: 0x98, B: 0x28, A: 0xff}
)

// brandTheme 自定义主题: 强制浅色变体, 覆盖主色/超链接色, 并统一使用内嵌中文字体.
type brandTheme struct {
	base fyne.Theme
	font fyne.Resource
}

// newBrandTheme 创建品牌主题.
// 返回值 *brandTheme, 主题实例.
func newBrandTheme() *brandTheme {
	return &brandTheme{base: theme.DefaultTheme(), font: chineseFont()}
}

// Color 返回主题颜色, 强制浅色变体, 并将主色改为藏蓝, 超链接改为金色.
//   - name, 颜色名.
//
// 返回值 color.Color, 颜色.
func (t *brandTheme) Color(name fyne.ThemeColorName, _ fyne.ThemeVariant) color.Color {
	switch name {
	case theme.ColorNamePrimary:
		return brandNavy
	case theme.ColorNameHyperlink:
		return brandGold
	default:
		return t.base.Color(name, theme.VariantLight)
	}
}

// Font 返回内嵌中文字体资源 (所有字重复用同一资源, 保证中文不显示为方块).
//   - style, 文本样式 (此处忽略, 统一返回同一字体).
//
// 返回值 fyne.Resource, 字体资源.
func (t *brandTheme) Font(_ fyne.TextStyle) fyne.Resource {
	return t.font
}

// Icon 委托默认主题.
//   - name, 图标名.
//
// 返回值 fyne.Resource, 图标资源.
func (t *brandTheme) Icon(name fyne.ThemeIconName) fyne.Resource {
	return t.base.Icon(name)
}

// Size 委托默认主题.
//   - name, 尺寸名.
//
// 返回值 float32, 尺寸值.
func (t *brandTheme) Size(name fyne.ThemeSizeName) float32 {
	return t.base.Size(name)
}
