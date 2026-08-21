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

// brandTheme 自定义主题: 强制浅色变体, 覆盖主色/超链接色, 并统一使用内嵌中文字体 (常规/粗体).
type brandTheme struct {
	base     fyne.Theme
	font     fyne.Resource
	boldFont fyne.Resource
}

// newBrandTheme 创建品牌主题.
// 返回值 *brandTheme, 主题实例.
func newBrandTheme() *brandTheme {
	return &brandTheme{base: theme.DefaultTheme(), font: chineseFont(), boldFont: chineseBoldFont()}
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

// Font 返回内嵌中文字体资源: 加粗样式返回粗体, 其余返回常规, 保证中文不显示为方块.
//   - style, 文本样式 (仅区分是否加粗).
//
// 返回值 fyne.Resource, 字体资源.
func (t *brandTheme) Font(style fyne.TextStyle) fyne.Resource {
	if style.Bold {
		return t.boldFont
	}
	return t.font
}

// Icon 委托默认主题.
//   - name, 图标名.
//
// 返回值 fyne.Resource, 图标资源.
func (t *brandTheme) Icon(name fyne.ThemeIconName) fyne.Resource {
	return t.base.Icon(name)
}

// Size 覆盖部分尺寸以获得更简约大气的间距, 其余委托默认主题.
//   - name, 尺寸名.
//
// 返回值 float32, 尺寸值.
func (t *brandTheme) Size(name fyne.ThemeSizeName) float32 {
	switch name {
	case theme.SizeNamePadding:
		return 8
	case theme.SizeNameInnerPadding:
		return 10
	default:
		return t.base.Size(name)
	}
}

// accentTheme 在品牌主题基础上把主色改为金色, 供进度条等需要副主题色的局部子树使用.
type accentTheme struct {
	*brandTheme
}

// newAccentTheme 创建以金色为主色的强调主题.
// 返回值 *accentTheme, 强调主题实例.
func newAccentTheme() *accentTheme {
	return &accentTheme{brandTheme: newBrandTheme()}
}

// Color 将主色替换为品牌金, 其余沿用品牌主题.
//   - name, 颜色名.
//   - variant, 主题变体.
//
// 返回值 color.Color, 颜色.
func (t *accentTheme) Color(name fyne.ThemeColorName, variant fyne.ThemeVariant) color.Color {
	if name == theme.ColorNamePrimary {
		return brandGold
	}
	return t.brandTheme.Color(name, variant)
}
