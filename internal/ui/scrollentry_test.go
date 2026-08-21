// FilePath    : Power-BI-custom-sample-data\internal\ui\scrollentry_test.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 验证 scrollEntry 在长文本下不会吞掉外层滚轮事件, 以便用户可以滚动页面.
package ui

import (
	"image/color"
	"strings"
	"testing"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/test"

	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/i18n"
)

// absoluteHoverPoint 返回对象在画布中的可见命中点坐标, 避免长文本输入框中心点落到视口外.
//   - t, 当前测试实例.
//   - obj, 目标对象.
//
// 返回值 fyne.Position, 对象内部靠左的绝对坐标.
func absoluteHoverPoint(t *testing.T, obj fyne.CanvasObject) fyne.Position {
	t.Helper()

	positioner, ok := fyne.CurrentApp().Driver().(interface {
		AbsolutePositionForObject(fyne.CanvasObject) fyne.Position
	})
	if !ok {
		t.Fatal("current driver does not expose AbsolutePositionForObject")
	}

	abs := positioner.AbsolutePositionForObject(obj)
	size := obj.Size()
	hoverX := float32(20)
	if size.Width < hoverX {
		hoverX = size.Width / 2
	}
	return abs.Add(fyne.NewPos(hoverX, size.Height/2))
}

// newScrollHarness 创建带外层垂直滚动容器的测试窗口, 用于验证子控件上的滚轮是否会传递给页面.
//   - t, 当前测试实例.
//   - content, 放在页面顶部的目标内容.
//
// 返回值 fyne.Window, 测试窗口; *container.Scroll, 外层滚动容器.
func newScrollHarness(t *testing.T, content fyne.CanvasObject) (fyne.Window, *container.Scroll) {
	t.Helper()

	spacer := canvas.NewRectangle(color.Transparent)
	spacer.SetMinSize(fyne.NewSize(320, 1200))
	scroll := container.NewVScroll(container.NewVBox(content, spacer))
	win := test.NewTempWindow(t, scroll)
	win.Resize(fyne.NewSize(360, 240))
	return win, scroll
}

// TestDateFieldScrollWheelPassesToOuterScroll 验证日期输入框上的滚轮会驱动外层页面滚动.
func TestDateFieldScrollWheelPassesToOuterScroll(t *testing.T) {
	mgr := i18n.NewManager(config.LocaleZhCN)
	dialogWin := test.NewTempWindow(t, canvas.NewRectangle(color.Transparent))
	df := NewDateField(dialogWin, mgr, "form.startDate", time.Date(2026, 8, 21, 0, 0, 0, 0, time.Local))
	win, scroll := newScrollHarness(t, df.Object())

	test.Scroll(win.Canvas(), absoluteHoverPoint(t, df.entry), 0, -120)
	if scroll.Offset.Y <= 0 {
		t.Fatalf("expected outer scroll to move when wheel is over date entry, got offset=%v", scroll.Offset)
	}
}

// TestNumberFieldScrollWheelPassesToOuterScroll 验证数字输入框上的滚轮会驱动外层页面滚动.
func TestNumberFieldScrollWheelPassesToOuterScroll(t *testing.T) {
	nf := NewNumberField(config.MinProductCount, config.MaxProductCount, 10, config.MinProductCount+10)
	win, scroll := newScrollHarness(t, nf.Object())

	test.Scroll(win.Canvas(), absoluteHoverPoint(t, nf.entry), 0, -120)
	if scroll.Offset.Y <= 0 {
		t.Fatalf("expected outer scroll to move when wheel is over number entry, got offset=%v", scroll.Offset)
	}
}

// TestScrollEntryLongTextPassesToOuterScroll 验证长目录路径文本不会让 scrollEntry 重新吞掉外层滚轮.
func TestScrollEntryLongTextPassesToOuterScroll(t *testing.T) {
	entry := newScrollEntry()
	entry.SetText(strings.Repeat("C:/Users/jiaopengzi/Desktop/Power-BI-custom-sample-data/", 4))
	win, scroll := newScrollHarness(t, entry)

	test.Scroll(win.Canvas(), absoluteHoverPoint(t, entry), 0, -120)
	if scroll.Offset.Y <= 0 {
		t.Fatalf("expected outer scroll to move when wheel is over long-text entry, got offset=%v", scroll.Offset)
	}
}

// TestMainWindowScrollWheelPassesToOuterScrollOnInputs 验证真实主窗口里, 日期框、数字框和目录框上的滚轮会驱动页面滚动.
func TestMainWindowScrollWheelPassesToOuterScrollOnInputs(t *testing.T) {
	a := test.NewApp()
	t.Cleanup(a.Quit)

	mw := newMainWindow(a)
	spacer := canvas.NewRectangle(color.Transparent)
	spacer.SetMinSize(fyne.NewSize(900, 480))
	mw.resultBox.Objects = append(mw.resultBox.Objects, spacer)
	mw.resultBox.Show()
	mw.resultBox.Refresh()
	mw.win.Resize(fyne.NewSize(900, 520))
	if mw.bodyScroll == nil {
		t.Fatal("main window body scroll was not initialized")
	}

	targets := []struct {
		name  string
		entry fyne.CanvasObject
	}{
		{name: "start date", entry: mw.startDF.entry},
		{name: "end date", entry: mw.endDF.entry},
		{name: "product count", entry: mw.productNF.entry},
		{name: "store count", entry: mw.storeNF.entry},
		{name: "inventory cycle", entry: mw.invNF.entry},
		{name: "output dir", entry: mw.dirEntry},
	}

	for _, target := range targets {
		mw.bodyScroll.Offset = fyne.NewPos(0, 0)
		mw.bodyScroll.Refresh()

		test.Scroll(mw.win.Canvas(), absoluteHoverPoint(t, target.entry), 0, -120)
		if mw.bodyScroll.Offset.Y <= 0 {
			t.Fatalf("expected outer scroll to move when wheel is over %s entry, got offset=%v", target.name, mw.bodyScroll.Offset)
		}
	}
}
