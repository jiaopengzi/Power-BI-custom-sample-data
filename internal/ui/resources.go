// FilePath    : internal/ui/resources.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 内嵌字体与图标资源 (go:embed 无法跨目录, 故资源置于本包内).

package ui

import (
	_ "embed"

	"fyne.io/fyne/v2"
)

// chineseFontData 内嵌的中文字体 (Noto Sans SC, OFL 许可, 见 assets/OFL.txt).
//
//go:embed assets/NotoSansSC.ttf
var chineseFontData []byte

// appIconData 内嵌的应用图标 (源自 build/appicon.png).
//
//go:embed assets/appicon.png
var appIconData []byte

// chineseFont 返回内嵌中文字体资源.
// 返回值 fyne.Resource, 字体资源.
func chineseFont() fyne.Resource {
	return fyne.NewStaticResource("NotoSansSC.ttf", chineseFontData)
}

// appIcon 返回内嵌应用图标资源.
// 返回值 fyne.Resource, 图标资源.
func appIcon() fyne.Resource {
	return fyne.NewStaticResource("appicon.png", appIconData)
}
