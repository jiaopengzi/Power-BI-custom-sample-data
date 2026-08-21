// FilePath    : internal/ui/mainwindow.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 主窗口, 复刻原 App.vue 的布局/交互/校验/运行时语言切换.

package ui

import (
	"fmt"
	"image/color"
	"math"
	"net/url"
	"path/filepath"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/ncruces/zenity"

	"jiaopengzi/Power-BI-custom-sample-data/internal/app"
	"jiaopengzi/Power-BI-custom-sample-data/internal/config"
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
	dateTemplateKey        = "date"
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
	dirEntry  *scrollEntry
	dirBorder *errorBorder

	// 动作按钮.
	generateBtn    *widget.Button
	incrementalBtn *widget.Button
	chooseBtn      *widget.Button
	openBtn        *widget.Button
	clearBtn       *widget.Button
	cutoffLabel    *widget.Label
	cutoffDate     time.Time
	hasCutoff      bool

	// 进度区.
	progressBox  *fyne.Container
	progressBar  *widget.ProgressBar
	stageLabel   *widget.Label
	percentLabel *widget.Label

	// 结果区.
	resultBox   *fyne.Container
	resultTitle *widget.Label
	resultTable *fyne.Container
	tables      []app.TableStat
	hasResult   bool
	bodyScroll  *container.Scroll

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
	w.win.Resize(fyne.NewSize(1440, 810))
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

	// 软件加载时: 若目录已有数据则展示历史结果表格, 否则隐藏.
	w.refreshResult()

	panel := container.NewVBox(form, w.progressBox, w.resultBox)
	w.bodyScroll = container.NewVScroll(container.NewPadded(panel))

	top := container.NewVBox(header, widget.NewSeparator())
	bottom := container.NewVBox(widget.NewSeparator(), footer)
	content := container.NewBorder(top, bottom, nil, nil, w.bodyScroll)
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

	return container.NewPadded(container.NewBorder(nil, nil, brand, localeSelect))
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
		container.New(layout.NewCustomPaddedLayout(0, 0, 0, 16),
			field(w.label("form.startDate"), w.startDF.Object())),
		container.New(layout.NewCustomPaddedLayout(0, 0, 16, 0),
			field(w.label("form.endDate"), w.endDF.Object())),
	)
	w.bind(func() { w.startDF.RefreshLocale(); w.endDF.RefreshLocale() })
	counts := container.NewGridWithColumns(3,
		container.New(layout.NewCustomPaddedLayout(0, 0, 0, 16),
			field(w.label("form.productCount"), w.productNF.Object())),
		container.New(layout.NewCustomPaddedLayout(0, 0, 8, 8),
			field(w.label("form.storeCount"), w.storeNF.Object())),
		container.New(layout.NewCustomPaddedLayout(0, 0, 16, 0),
			field(w.label("form.inventoryCycle"), w.invNF.Object())),
	)

	w.dirEntry = newScrollEntry()
	w.dirEntry.SetText(w.ctl.DefaultOutputDir())
	w.dirEntry.SetPlaceHolder(w.mgr.T("form.outputDirPlaceholder"))
	w.bind(func() { w.dirEntry.SetPlaceHolder(w.mgr.T("form.outputDirPlaceholder")) })
	w.dirBorder = newErrorBorder(w.dirEntry)
	w.dirEntry.OnChanged = func(s string) {
		w.dirBorder.setInvalid(strings.TrimSpace(s) == "")
		w.updateOpenButton()
		w.updateActionState()
		w.updateClearButton()
		w.refreshCutoff()
		w.refreshResult()
	}

	// 输入变化时联动执行按钮的可用性 (非法值禁用).
	w.startDF.OnChanged = func() { w.onDateChanged(w.startDF) }
	w.endDF.OnChanged = func() { w.onDateChanged(w.endDF) }
	w.productNF.OnChanged = func(int) { w.updateActionState() }
	w.storeNF.OnChanged = func(int) { w.updateActionState() }
	w.invNF.OnChanged = func(int) { w.updateActionState() }

	w.chooseBtn = w.iconButton("buttons.chooseDir", theme.FolderOpenIcon(), w.onChooseDir)
	w.openBtn = w.iconButton("buttons.openDir", theme.FolderIcon(), w.onOpenDir)
	dirRow := container.NewBorder(nil, nil, nil,
		container.NewHBox(w.chooseBtn, w.openBtn), w.dirBorder.Object())

	w.generateBtn = w.iconButton("buttons.generate", theme.DocumentIcon(), func() { w.run(kindFull) })
	w.generateBtn.Importance = widget.HighImportance
	w.incrementalBtn = w.iconButton("buttons.incremental", theme.HistoryIcon(), func() { w.run(kindInc) })
	w.incrementalBtn.Importance = widget.HighImportance
	w.clearBtn = w.iconButton("buttons.clearData", theme.DeleteIcon(), w.onClearData)
	w.clearBtn.Importance = widget.DangerImportance
	w.cutoffLabel = widget.NewLabel("")
	w.cutoffLabel.Hide()
	// 两主按钮置于等宽两列网格保持尺寸一致; 清空(危险色)与截止日期标签置于其旁.
	buttons := container.NewGridWithColumns(2, w.generateBtn, w.incrementalBtn)
	actions := container.NewHBox(buttons, w.clearBtn, w.cutoffLabel)

	w.updateOpenButton()
	w.updateActionState()
	w.updateClearButton()
	w.refreshCutoff()

	return container.NewVBox(
		dates,
		gap(),
		counts,
		gap(),
		field(w.label("form.outputDir"), dirRow),
		gap(),
		actions,
	)
}

// buildProgress 构建进度区 (阶段文案 + 百分比 + 进度条), 初始隐藏.
func (w *mainWindow) buildProgress() {
	w.stageLabel = widget.NewLabel("")
	w.percentLabel = widget.NewLabel("")
	w.progressBar = widget.NewProgressBar()
	head := container.NewBorder(nil, nil, w.stageLabel, w.percentLabel)
	// 进度条使用副主题色 (金), 仅在该子树内覆盖主色.
	bar := container.NewThemeOverride(w.progressBar, newAccentTheme())
	w.progressBox = container.NewVBox(head, bar)
	w.progressBox.Hide()
}

// buildResult 构建结果区 (表名/行数/大小 表格), 初始隐藏.
func (w *mainWindow) buildResult() {
	w.resultTitle = widget.NewLabelWithStyle(w.mgr.T("result.title"), fyne.TextAlignLeading, fyne.TextStyle{Bold: true})
	w.bind(func() { w.resultTitle.SetText(w.mgr.T("result.title")) })
	w.resultTable = container.NewVBox()
	w.resultBox = container.NewVBox(w.resultTitle, w.resultTable)
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

// refreshTexts 重新应用所有已注册文案, 并刷新阶段/结果/截止日期的动态文案.
func (w *mainWindow) refreshTexts() {
	for _, fn := range w.relabels {
		fn()
	}
	if w.lastStageKey != "" {
		w.stageLabel.SetText(w.mgr.T("stages." + w.lastStageKey))
	}
	if w.hasCutoff {
		w.cutoffLabel.SetText(w.mgr.Tf("info.cutoff", map[string]string{dateTemplateKey: w.cutoffDate.Format(config.DateLayout)}))
	}
	w.renderResult()
}

// onChooseDir 弹出系统原生目录选择, 选定后写回目录输入框.
func (w *mainWindow) onChooseDir() {
	start := strings.TrimSpace(w.dirEntry.Text)
	go func() {
		dir, err := zenity.SelectFile(
			zenity.Title(w.mgr.T("buttons.chooseDir")),
			zenity.Directory(),
			zenity.Filename(start),
		)
		if err != nil || strings.TrimSpace(dir) == "" {
			return
		}
		fyne.Do(func() {
			w.dirEntry.SetText(filepath.Clean(dir))
			w.updateOpenButton()
		})
	}()
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
	if kind == kindInc && w.hasCutoff {
		expected := w.cutoffDate.AddDate(0, 0, 1)
		if start, err := time.Parse(config.DateLayout, w.startDF.Text()); err != nil || !start.Equal(expected) {
			w.failKey("msg.incStartMismatch", map[string]string{"date": expected.Format(config.DateLayout)})
			return
		}
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
	label := widget.NewLabel(w.mgr.T("msg.overwriteContent"))
	label.Wrapping = fyne.TextWrapWord
	// 透明占位撑宽弹窗, 避免正文过短时对话框显得狭小.
	spacer := canvas.NewRectangle(color.Transparent)
	spacer.SetMinSize(fyne.NewSize(460, 0))
	content := container.NewVBox(spacer, label)
	d := dialog.NewCustomConfirm(
		w.mgr.T("msg.overwriteTitle"),
		w.mgr.T("msg.overwriteConfirm"),
		w.mgr.T("msg.overwriteCancel"),
		content,
		func(ok bool) {
			if ok {
				onConfirm()
			}
		},
		w.win,
	)
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
			w.renderResponse(resp)
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

// renderResponse 按响应码处理结果: 完成后隐藏进度条; 成功展示表格, 异常弹提示.
//   - resp, 控制器响应.
func (w *mainWindow) renderResponse(resp app.Response) {
	w.resetProgress()
	switch resp.Code {
	case app.CodeOK:
		w.showResult(resp.Tables)
		w.refreshCutoff()
		w.updateClearButton()
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
	w.updateActionState()
	w.updateClearButton()
	if running {
		w.progressBox.Show()
	}
	w.updateOpenButton()
}

// resetProgress 将进度区恢复到初始态 (隐藏并清零), 用于校验失败/执行异常后.
func (w *mainWindow) resetProgress() {
	w.lastStageKey = ""
	w.progressBar.SetValue(0)
	w.progressBox.Hide()
}

// setControlsDisabled 统一启用/禁用输入控件 (执行按钮由 updateActionState 根据有效性管理, 打开目录按钮除外).
//   - disabled, 是否禁用.
func (w *mainWindow) setControlsDisabled(disabled bool) {
	if disabled {
		w.chooseBtn.Disable()
	} else {
		w.chooseBtn.Enable()
	}
	w.startDF.SetDisabled(disabled)
	w.endDF.SetDisabled(disabled)
	w.productNF.SetDisabled(disabled)
	w.storeNF.SetDisabled(disabled)
	w.invNF.SetDisabled(disabled)
}

// formValid 判断当前表单是否均为合法值 (目录非空/数量在区间/日期合法且结束晚于开始).
// 返回值 bool, 全部合法返回 true.
func (w *mainWindow) formValid() bool {
	if strings.TrimSpace(w.dirEntry.Text) == "" {
		return false
	}
	if !w.productNF.Valid() || !w.storeNF.Valid() || !w.invNF.Valid() {
		return false
	}
	if !w.startDF.Valid() || !w.endDF.Valid() {
		return false
	}
	return w.endDF.Text() > w.startDF.Text()
}

// updateActionState 根据运行态与表单有效性切换生成/增量按钮的可用性.
func (w *mainWindow) updateActionState() {
	if w.running || !w.formValid() {
		w.generateBtn.Disable()
		w.incrementalBtn.Disable()
		return
	}
	w.generateBtn.Enable()
	w.incrementalBtn.Enable()
}

// onDateChanged 日期变化时做先后校验: 两端格式均合法但开始晚于/等于结束时,
// 将错误标在刚编辑的字段下方, 并清除另一字段的跨字段错误.
//   - changed, 刚发生变化的日期字段.
func (w *mainWindow) onDateChanged(changed *DateField) {
	if w.startDF.Valid() && w.endDF.Valid() && w.endDF.Text() <= w.startDF.Text() {
		changed.SetExternalError("form.dateOrder")
		w.otherDate(changed).SetExternalError("")
	} else {
		w.startDF.SetExternalError("")
		w.endDF.SetExternalError("")
	}
	w.updateActionState()
}

// otherDate 返回两个日期字段中另一个.
//   - d, 当前字段.
//
// 返回值 *DateField, 另一个日期字段.
func (w *mainWindow) otherDate(d *DateField) *DateField {
	if d == w.startDF {
		return w.endDF
	}
	return w.startDF
}

// refreshCutoff 重读事实表截止日期并更新增量按钮旁的提示 (启动/目录变化/生成后调用).
func (w *mainWindow) refreshCutoff() {
	if w.cutoffLabel == nil {
		return
	}
	if last, ok := w.ctl.LastFactDate(strings.TrimSpace(w.dirEntry.Text)); ok {
		w.cutoffDate = last
		w.hasCutoff = true
		w.cutoffLabel.SetText(w.mgr.Tf("info.cutoff", map[string]string{"date": last.Format(config.DateLayout)}))
		w.cutoffLabel.Show()
		return
	}
	w.hasCutoff = false
	w.cutoffLabel.Hide()
}

// refreshResult 依据目录是否已有数据决定展示/隐藏结果表格 (启动与目录变化时调用).
func (w *mainWindow) refreshResult() {
	if w.resultBox == nil {
		return
	}
	if tables := w.ctl.ExistingTables(strings.TrimSpace(w.dirEntry.Text)); len(tables) > 0 {
		w.showResult(tables)
		return
	}
	w.hideResult()
}

// updateClearButton 依据运行态与目录是否已有数据切换清空按钮可用性.
func (w *mainWindow) updateClearButton() {
	if w.clearBtn == nil {
		return
	}
	if w.running || !w.ctl.HasData(strings.TrimSpace(w.dirEntry.Text)) {
		w.clearBtn.Disable()
		return
	}
	w.clearBtn.Enable()
}

// onClearData 弹出确认后清空目录中已生成的数据 (保留目录), 并刷新界面状态.
func (w *mainWindow) onClearData() {
	dir := strings.TrimSpace(w.dirEntry.Text)
	if dir == "" || !w.ctl.HasData(dir) {
		return
	}
	label := widget.NewLabel(w.mgr.T("msg.clearContent"))
	label.Wrapping = fyne.TextWrapWord
	spacer := canvas.NewRectangle(color.Transparent)
	spacer.SetMinSize(fyne.NewSize(460, 0))
	content := container.NewVBox(spacer, label)
	d := dialog.NewCustomConfirm(
		w.mgr.T("msg.clearTitle"),
		w.mgr.T("msg.clearConfirm"),
		w.mgr.T("msg.clearCancel"),
		content,
		func(ok bool) {
			if !ok {
				return
			}
			if err := w.ctl.ClearData(dir); err != nil {
				w.fail(err.Error())
				return
			}
			w.refreshResult()
			w.refreshCutoff()
			w.updateClearButton()
			w.info("msg.clearSuccess")
		},
		w.win,
	)
	d.Show()
}

// updateOpenButton 依据运行态与目录是否为空, 切换打开目录按钮可用性.
func (w *mainWindow) updateOpenButton() {
	if w.running || strings.TrimSpace(w.dirEntry.Text) == "" {
		w.openBtn.Disable()
		return
	}
	w.openBtn.Enable()
}

// showResult 记录并显示生成结果表格.
//   - tables, 各表名称/行数/大小.
func (w *mainWindow) showResult(tables []app.TableStat) {
	w.tables = tables
	w.hasResult = true
	w.renderResult()
	w.resultBox.Show()
}

// hideResult 隐藏结果区.
func (w *mainWindow) hideResult() {
	w.hasResult = false
	w.resultBox.Hide()
}

// renderResult 以当前语言重建结果表格 (三列等宽铺满, 内容居中, 表头加粗).
// 使用普通网格而非 widget.Table, 以免表格内部滚动拦截页面滚轮事件.
func (w *mainWindow) renderResult() {
	if !w.hasResult {
		return
	}
	rowsUnit := w.mgr.T("result.rows")
	head := func(key string) fyne.CanvasObject {
		return widget.NewLabelWithStyle(w.mgr.T(key), fyne.TextAlignCenter, fyne.TextStyle{Bold: true})
	}
	cell := func(s string) fyne.CanvasObject {
		return widget.NewLabelWithStyle(s, fyne.TextAlignCenter, fyne.TextStyle{})
	}
	header := container.NewGridWithColumns(3,
		head("result.colTable"), head("result.colRows"), head("result.colSize"),
	)
	cells := make([]fyne.CanvasObject, 0, len(w.tables)*3)
	for _, t := range w.tables {
		cells = append(cells,
			cell(t.Name),
			cell(fmt.Sprintf("%d %s", t.Rows, rowsUnit)),
			cell(formatSize(t.Size)),
		)
	}
	data := container.NewGridWithColumns(3, cells...)
	w.resultTable.Objects = []fyne.CanvasObject{header, widget.NewSeparator(), data}
	w.resultTable.Refresh()
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

// gap 返回一个用于拉开表单分组间距的透明竖向占位块.
// 返回值 fyne.CanvasObject, 占位块.
func gap() fyne.CanvasObject {
	spacer := canvas.NewRectangle(color.Transparent)
	spacer.SetMinSize(fyne.NewSize(0, 8))
	return spacer
}

// formatSize 将字节数格式化为便于阅读的单位 (B/KB/MB/GB).
//   - size, 字节数.
//
// 返回值 string, 可读文本.
func formatSize(size int64) string {
	const unit = 1024
	if size < unit {
		return fmt.Sprintf("%d B", size)
	}
	div, exp := int64(unit), 0
	for n := size / unit; n >= unit; n /= unit {
		div *= unit
		exp++
	}
	return fmt.Sprintf("%.1f %cB", float64(size)/float64(div), "KMGT"[exp])
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
