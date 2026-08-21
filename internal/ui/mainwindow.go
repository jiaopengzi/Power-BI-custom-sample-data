// FilePath    : internal/ui/mainwindow.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 主窗口, 复刻原 App.vue 的布局/交互/校验/运行时语言切换.

package ui

import (
	"fmt"
	"math"
	"net/url"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"

	"jiaopengzi/Power-BI-custom-sample-data/internal/app"
	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
	"jiaopengzi/Power-BI-custom-sample-data/internal/generator"
	"jiaopengzi/Power-BI-custom-sample-data/internal/i18n"
)

// 窗口与文档常量.
const (
	windowTitle = "Power BI Custom Sample Data"
	docsURL     = "https://jiaopengzi.com/?post_id=19051044919050241"
)

// 语言下拉的显示名 (以各自语言书写, 不随界面语言变化).
const (
	langChinese = "简体中文"
	langEnglish = "English"
)

// 运行类型.
const (
	kindFull = "full"
	kindInc  = "inc"
)

// 表单默认值 (与原前端 App.vue 一致).
const (
	defaultStartOffsetDays = -1600
	defaultProductCount    = 200
	defaultStoreCount      = 5
	defaultInventoryCycle  = 14
	numberStep             = 1
)

// mainWindow 持有窗口, 编排控制器, 文案管理器与全部控件引用,
// 以支持运行时语言切换与运行态启用/禁用.
type mainWindow struct {
	win    fyne.Window
	ctl    *app.Controller
	mgr    *i18n.Manager
	locale config.Locale

	// 表单控件.
	startDF   *DateField
	endDF     *DateField
	productNF *NumberField
	storeNF   *NumberField
	invNF     *NumberField
	dirEntry  *widget.Entry

	// 动作按钮.
	generateBtn    *widget.Button
	incrementalBtn *widget.Button
	chooseBtn      *widget.Button
	openBtn        *widget.Button

	// 进度区.
	progressBox  *fyne.Container
	progressBar  *widget.ProgressBar
	stageLabel   *widget.Label
	percentLabel *widget.Label

	// 结果区.
	resultBox    *fyne.Container
	resultValues []*widget.Label
	result       generator.Result
	hasResult    bool

	// 运行时状态.
	relabels     []func()
	running      bool
	lastStageKey string
}

// newMainWindow 创建并初始化主窗口.
//   - a, Fyne 应用实例.
//
// 返回值 *mainWindow, 主窗口.
func newMainWindow(a fyne.App) *mainWindow {
	w := &mainWindow{
		ctl:    app.New(),
		mgr:    i18n.NewManager(config.LocaleZhCN),
		locale: config.LocaleZhCN,
	}
	w.win = a.NewWindow(windowTitle)
	w.win.SetIcon(appIcon())
	w.win.SetContent(w.buildContent())
	w.win.Resize(fyne.NewSize(1000, 760))
	w.win.CenterOnScreen()
	return w
}

// buildContent 组装窗口内容: 顶部品牌栏, 中部表单/进度/结果, 底部文档链接.
// 返回值 fyne.CanvasObject, 根内容对象.
func (w *mainWindow) buildContent() fyne.CanvasObject {
	header := w.buildHeader()
	form := w.buildForm()
	w.buildProgress()
	w.buildResult()
	footer := w.buildFooter()

	panel := container.NewVBox(form, w.progressBox, w.resultBox)
	body := container.NewVScroll(panel)
	content := container.NewBorder(header, footer, nil, nil, body)
	return container.NewPadded(content)
}

// buildHeader 构建顶部品牌栏与语言下拉.
// 返回值 fyne.CanvasObject, 顶部栏.
func (w *mainWindow) buildHeader() fyne.CanvasObject {
	logo := canvas.NewImageFromResource(appIcon())
	logo.FillMode = canvas.ImageFillContain
	logo.SetMinSize(fyne.NewSize(40, 40))

	title := widget.NewLabelWithStyle(w.mgr.T("subtitle"), fyne.TextAlignLeading, fyne.TextStyle{Bold: true})
	w.bind(func() { title.SetText(w.mgr.T("subtitle")) })
	brand := container.NewHBox(logo, title)

	localeSelect := widget.NewSelect([]string{langChinese, langEnglish}, w.onLocaleChange)
	localeSelect.SetSelected(langChinese)

	return container.NewBorder(nil, nil, brand, localeSelect)
}

// buildForm 构建日期/数量/目录表单与动作按钮.
// 返回值 fyne.CanvasObject, 表单区.
func (w *mainWindow) buildForm() fyne.CanvasObject {
	now := time.Now()
	w.startDF = NewDateField(w.win, w.mgr, "form.startDate", now.AddDate(0, 0, defaultStartOffsetDays))
	w.endDF = NewDateField(w.win, w.mgr, "form.endDate", now)
	w.productNF = NewNumberField(config.MinProductCount, config.MaxProductCount, numberStep, defaultProductCount)
	w.storeNF = NewNumberField(config.MinStoreCount, config.MaxStoreCount, numberStep, defaultStoreCount)
	w.invNF = NewNumberField(config.MinInventory, config.MaxInventory, numberStep, defaultInventoryCycle)

	dates := container.NewGridWithColumns(2,
		field(w.label("form.startDate"), w.startDF.Object()),
		field(w.label("form.endDate"), w.endDF.Object()),
	)
	counts := container.NewGridWithColumns(3,
		field(w.label("form.productCount"), w.productNF.Object()),
		field(w.label("form.storeCount"), w.storeNF.Object()),
		field(w.label("form.inventoryCycle"), w.invNF.Object()),
	)

	w.dirEntry = widget.NewEntry()
	w.dirEntry.SetText(w.ctl.DefaultOutputDir())
	w.dirEntry.SetPlaceHolder(w.mgr.T("form.outputDirPlaceholder"))
	w.bind(func() { w.dirEntry.SetPlaceHolder(w.mgr.T("form.outputDirPlaceholder")) })

	w.chooseBtn = w.iconButton("buttons.chooseDir", theme.FolderOpenIcon(), w.onChooseDir)
	w.openBtn = w.iconButton("buttons.openDir", theme.FolderIcon(), w.onOpenDir)
	dirRow := container.NewBorder(nil, nil, nil,
		container.NewHBox(w.chooseBtn, w.openBtn), w.dirEntry)

	w.generateBtn = w.iconButton("buttons.generate", theme.DocumentIcon(), func() { w.run(kindFull) })
	w.generateBtn.Importance = widget.HighImportance
	w.incrementalBtn = w.iconButton("buttons.incremental", theme.HistoryIcon(), func() { w.run(kindInc) })
	actions := container.NewHBox(w.generateBtn, w.incrementalBtn)

	w.updateOpenButton()

	return container.NewVBox(
		dates,
		counts,
		field(w.label("form.outputDir"), dirRow),
		actions,
	)
}

// buildProgress 构建进度区 (阶段文案 + 百分比 + 进度条), 初始隐藏.
func (w *mainWindow) buildProgress() {
	w.stageLabel = widget.NewLabel("")
	w.percentLabel = widget.NewLabel("")
	w.progressBar = widget.NewProgressBar()
	head := container.NewBorder(nil, nil, w.stageLabel, w.percentLabel)
	w.progressBox = container.NewVBox(head, w.progressBar)
	w.progressBox.Hide()
}

// buildResult 构建结果区 (6 项行数统计), 初始隐藏.
func (w *mainWindow) buildResult() {
	title := widget.NewLabelWithStyle(w.mgr.T("result.title"), fyne.TextAlignLeading, fyne.TextStyle{Bold: true})
	w.bind(func() { title.SetText(w.mgr.T("result.title")) })

	keys := []string{
		"result.products", "result.stores", "result.customers",
		"result.inventory", "result.orders", "result.orderItem",
	}
	w.resultValues = make([]*widget.Label, len(keys))
	cells := make([]fyne.CanvasObject, 0, len(keys))
	for i, k := range keys {
		val := widget.NewLabel("")
		w.resultValues[i] = val
		cells = append(cells, container.NewHBox(w.label(k), val))
	}
	grid := container.NewGridWithColumns(3, cells...)
	w.resultBox = container.NewVBox(title, grid)
	w.resultBox.Hide()
}

// buildFooter 构建底部文档链接.
// 返回值 fyne.CanvasObject, 页脚.
func (w *mainWindow) buildFooter() fyne.CanvasObject {
	link := widget.NewHyperlink(w.mgr.T("buttons.docs"), mustURL(docsURL))
	w.bind(func() { link.SetText(w.mgr.T("buttons.docs")) })
	return container.NewCenter(link)
}

// onLocaleChange 语言下拉变化时切换界面与产物语言并刷新全部文案.
//   - sel, 下拉选中的显示名.
func (w *mainWindow) onLocaleChange(sel string) {
	loc := config.LocaleZhCN
	if sel == langEnglish {
		loc = config.LocaleEnUS
	}
	w.mgr.SetLocale(loc)
	w.locale = loc
	w.refreshTexts()
}

// refreshTexts 重新应用所有已注册文案, 并刷新阶段/结果的动态文案.
func (w *mainWindow) refreshTexts() {
	for _, fn := range w.relabels {
		fn()
	}
	if w.lastStageKey != "" {
		w.stageLabel.SetText(w.mgr.T("stages." + w.lastStageKey))
	}
	w.renderResult()
}

// onChooseDir 弹出目录选择, 选定后写回目录输入框.
func (w *mainWindow) onChooseDir() {
	dialog.ShowFolderOpen(func(uri fyne.ListableURI, err error) {
		if err != nil || uri == nil {
			return
		}
		w.dirEntry.SetText(normalizeDir(uri.Path()))
		w.updateOpenButton()
	}, w.win)
}

// onOpenDir 在系统文件管理器中打开当前目录.
func (w *mainWindow) onOpenDir() {
	dir := strings.TrimSpace(w.dirEntry.Text)
	if dir == "" {
		return
	}
	if err := w.ctl.OpenOutputDir(dir); err != nil {
		w.fail(err.Error())
	}
}

// run 校验并触发一次生成 (全量时若已有数据先弹覆盖确认).
//   - kind, kindFull 或 kindInc.
func (w *mainWindow) run(kind string) {
	if w.running || !w.validate() {
		return
	}
	if kind == kindFull && w.ctl.HasData(strings.TrimSpace(w.dirEntry.Text)) {
		w.confirmOverwrite(func() { w.execute(kind) })
		return
	}
	w.execute(kind)
}

// validate 复刻 App.vue 的校验顺序与提示文案.
// 返回值 bool, 全部通过返回 true.
func (w *mainWindow) validate() bool {
	if strings.TrimSpace(w.dirEntry.Text) == "" {
		w.info("msg.chooseDirFirst")
		return false
	}
	if p := w.productNF.Value(); p < config.MinProductCount || p > config.MaxProductCount {
		w.info("msg.productRange")
		return false
	}
	if s := w.storeNF.Value(); s < config.MinStoreCount || s > config.MaxStoreCount {
		w.info("msg.storeRange")
		return false
	}
	if iv := w.invNF.Value(); iv < config.MinInventory || iv > config.MaxInventory {
		w.info("msg.inventoryRange")
		return false
	}
	if w.endDF.Text() <= w.startDF.Text() {
		w.info("msg.invalidRange")
		return false
	}
	return true
}

// confirmOverwrite 弹出覆盖确认, 确认后执行 onConfirm.
//   - onConfirm, 确认回调.
func (w *mainWindow) confirmOverwrite(onConfirm func()) {
	d := dialog.NewConfirm(
		w.mgr.T("msg.overwriteTitle"),
		w.mgr.T("msg.overwriteContent"),
		func(ok bool) {
			if ok {
				onConfirm()
			}
		},
		w.win,
	)
	d.SetConfirmText(w.mgr.T("msg.overwriteConfirm"))
	d.SetDismissText(w.mgr.T("msg.overwriteCancel"))
	d.Show()
}

// execute 在后台执行生成/增量, 通过 fyne.Do 编组进度与结果到 UI 线程.
//   - kind, kindFull 或 kindInc.
func (w *mainWindow) execute(kind string) {
	w.setRunning(true)
	w.hideResult()
	w.setStage(0, "stageStart")
	params := w.params()

	go func() {
		prog := func(pct float64, stage string) {
			fyne.Do(func() { w.setStage(pct, stage) })
		}
		var resp app.Response
		if kind == kindFull {
			resp = w.ctl.GenerateSample(params, prog)
		} else {
			resp = w.ctl.IncrementalUpdate(params, prog)
		}
		fyne.Do(func() {
			w.setRunning(false)
			w.renderResponse(kind, resp)
		})
	}()
}

// params 从当前控件读取生成参数.
// 返回值 app.Params, 生成参数.
func (w *mainWindow) params() app.Params {
	return app.Params{
		OutputDir:      strings.TrimSpace(w.dirEntry.Text),
		Locale:         string(w.locale),
		ProductCount:   w.productNF.Value(),
		StoreCount:     w.storeNF.Value(),
		InventoryCycle: w.invNF.Value(),
		StartDate:      w.startDF.Text(),
		EndDate:        w.endDF.Text(),
	}
}

// renderResponse 按响应码复刻 App.vue 的分支处理.
//   - kind, 运行类型 (决定成功文案).
//   - resp, 控制器响应.
func (w *mainWindow) renderResponse(kind string, resp app.Response) {
	switch resp.Code {
	case app.CodeOK:
		w.showResult(resp.Result)
		if kind == kindFull {
			w.info("msg.generateSuccess")
		} else {
			w.info("msg.incrementalSuccess")
		}
	case app.CodeNoBaseData:
		w.info("msg.noBaseData")
	case app.CodeDateConflict:
		w.failKey("msg.dateConflict", map[string]string{"range": resp.Message})
	default:
		w.failKey("msg.failed", map[string]string{"msg": resp.Message})
	}
}

// setStage 更新进度条/阶段文案/百分比并显示进度区.
//   - pct, 进度百分比 (0..100).
//   - stageKey, 阶段标识 (对应 stages.* 文案).
func (w *mainWindow) setStage(pct float64, stageKey string) {
	w.lastStageKey = stageKey
	w.progressBar.SetValue(pct / 100)
	w.percentLabel.SetText(fmt.Sprintf("%d%%", int(math.Round(pct))))
	w.stageLabel.SetText(w.mgr.T("stages." + stageKey))
	w.progressBox.Show()
}

// setRunning 切换运行态: 禁用/启用控件并按需显示进度区.
//   - running, 是否运行中.
func (w *mainWindow) setRunning(running bool) {
	w.running = running
	w.setControlsDisabled(running)
	if running {
		w.progressBox.Show()
	}
	w.updateOpenButton()
}

// setControlsDisabled 统一启用/禁用输入控件与动作按钮 (打开目录按钮除外).
//   - disabled, 是否禁用.
func (w *mainWindow) setControlsDisabled(disabled bool) {
	toggle := func(b *widget.Button) {
		if disabled {
			b.Disable()
			return
		}
		b.Enable()
	}
	toggle(w.generateBtn)
	toggle(w.incrementalBtn)
	toggle(w.chooseBtn)
	w.startDF.SetDisabled(disabled)
	w.endDF.SetDisabled(disabled)
	w.productNF.SetDisabled(disabled)
	w.storeNF.SetDisabled(disabled)
	w.invNF.SetDisabled(disabled)
}

// updateOpenButton 依据运行态与目录是否为空, 切换打开目录按钮可用性.
func (w *mainWindow) updateOpenButton() {
	if w.running || strings.TrimSpace(w.dirEntry.Text) == "" {
		w.openBtn.Disable()
		return
	}
	w.openBtn.Enable()
}

// showResult 记录并显示生成结果.
//   - res, 各表行数统计.
func (w *mainWindow) showResult(res generator.Result) {
	w.result = res
	w.hasResult = true
	w.renderResult()
	w.resultBox.Show()
}

// hideResult 隐藏结果区.
func (w *mainWindow) hideResult() {
	w.hasResult = false
	w.resultBox.Hide()
}

// renderResult 以当前语言刷新结果各项文案 (含行数单位).
func (w *mainWindow) renderResult() {
	if !w.hasResult {
		return
	}
	rows := w.mgr.T("result.rows")
	vals := []int{
		w.result.Products, w.result.Stores, w.result.Customers,
		w.result.Inventory, w.result.Orders, w.result.OrderItem,
	}
	for i, l := range w.resultValues {
		l.SetText(fmt.Sprintf("%d %s", vals[i], rows))
	}
}

// info 以提示样式弹出一条信息.
//   - key, 文案键.
func (w *mainWindow) info(key string) {
	w.notify(w.mgr.T("common.info"), w.mgr.T(key))
}

// fail 以错误样式弹出一条原始消息 (用于系统错误).
//   - msg, 消息文本.
func (w *mainWindow) fail(msg string) {
	w.notify(w.mgr.T("common.error"), msg)
}

// failKey 以错误样式弹出带占位符替换的文案.
//   - key, 文案键.
//   - vars, 占位符替换表.
func (w *mainWindow) failKey(key string, vars map[string]string) {
	w.notify(w.mgr.T("common.error"), w.mgr.Tf(key, vars))
}

// notify 弹出一个带标题与关闭按钮的自定义对话框.
//   - title, 标题.
//   - msg, 正文.
func (w *mainWindow) notify(title, msg string) {
	dialog.ShowCustom(title, w.mgr.T("common.ok"), widget.NewLabel(msg), w.win)
}

// bind 注册一个语言刷新闭包 (语言切换时统一执行).
//   - fn, 刷新闭包.
func (w *mainWindow) bind(fn func()) {
	w.relabels = append(w.relabels, fn)
}

// label 创建随语言刷新的文本标签.
//   - key, 文案键.
//
// 返回值 *widget.Label, 标签.
func (w *mainWindow) label(key string) *widget.Label {
	l := widget.NewLabel(w.mgr.T(key))
	w.bind(func() { l.SetText(w.mgr.T(key)) })
	return l
}

// iconButton 创建随语言刷新文本的带图标按钮.
//   - key, 文案键.
//   - icon, 图标资源.
//   - tapped, 点击回调.
//
// 返回值 *widget.Button, 按钮.
func (w *mainWindow) iconButton(key string, icon fyne.Resource, tapped func()) *widget.Button {
	b := widget.NewButtonWithIcon(w.mgr.T(key), icon, tapped)
	w.bind(func() { b.SetText(w.mgr.T(key)) })
	return b
}

// field 将标题标签与输入控件竖直组合为一个表单项.
//   - caption, 标题标签.
//   - input, 输入控件.
//
// 返回值 *fyne.Container, 表单项容器.
func field(caption, input fyne.CanvasObject) *fyne.Container {
	return container.NewVBox(caption, input)
}

// mustURL 解析 URL, 解析失败时返回空 URL (docsURL 为常量, 正常不会失败).
//   - raw, 原始地址.
//
// 返回值 *url.URL, 解析结果.
func mustURL(raw string) *url.URL {
	u, err := url.Parse(raw)
	if err != nil {
		return &url.URL{}
	}
	return u
}

// normalizeDir 规范化目录对话框返回的路径 (Windows 下去除前导斜杠并转反斜杠).
//   - p, 原始路径.
//
// 返回值 string, 规范化路径.
func normalizeDir(p string) string {
	if runtime.GOOS == "windows" {
		p = strings.TrimPrefix(p, "/")
	}
	return filepath.FromSlash(p)
}
