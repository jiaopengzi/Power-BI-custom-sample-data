// FilePath    : internal/i18n/messages.go
// Author      : jiaopengzi
// Blog        : https://jiaopengzi.com
// Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
// Description : 界面文案表 (简体中文 / 英文), 键名镜像原前端 i18n.

package i18n

// entry 单条文案的中英文取值.
//   - zh, 简体中文.
//   - en, 英文.
type entry struct {
	zh string
	en string
}

// brandName 应用品牌名 (中英文一致), 依据重构要求由 "Power BI Sample Data Generator" 更名而来.
const (
	brandName      = "Power BI Custom Sample Data"
	cancelTextEnUS = "Cancel"
)

// catalog 扁平点号键的文案表, 单表存放中英文以避免重复的镜像 map.
// 键名与原前端 frontend/src/i18n 完全一致; stages.* 键与 generator 发出的阶段标识对齐.
// 品牌名依据重构要求由 "Power BI Sample Data Generator" 更名为 "Power BI Custom Sample Data".
var catalog = map[string]entry{
	"title":    {brandName, brandName},
	"subtitle": {"Power BI 自定义示例数据生成器", brandName},

	// common.* 为 Fyne 对话框所需 (标题/按钮), 原前端用 Element Plus 轻提示无需此类文案.
	"common.ok":     {"确定", "OK"},
	"common.cancel": {"取消", cancelTextEnUS},
	"common.info":   {"提示", "Notice"},
	"common.error":  {"错误", "Error"},

	"form.startDate":            {"开始日期", "Start Date"},
	"form.endDate":              {"结束日期", "End Date"},
	"form.productCount":         {"产品数量", "Products"},
	"form.storeCount":           {"门店数量", "Stores"},
	"form.inventoryCycle":       {"入库周期", "Inventory Cycle"},
	"form.outputDir":            {"存放目录", "Output Directory"},
	"form.outputDirPlaceholder": {"请先指定数据存放目录", "Please choose an output directory first"},
	"form.localeLabel":          {"语言", "Language"},
	"form.dateInvalid":          {"日期格式错误, 应为 YYYY-MM-DD", "Invalid date, expected YYYY-MM-DD"},
	"form.dateOrder":            {"开始日期需早于结束日期", "Start date must be before end date"},

	"buttons.generate":    {"生成示例数据", "Generate Sample Data"},
	"buttons.incremental": {"增量更新", "Incremental Update"},
	"buttons.clearData":   {"清空数据", "Clear Data"},
	"buttons.chooseDir":   {"指定存放目录", "Choose Directory"},
	"buttons.openDir":     {"打开存放目录", "Open Directory"},
	"buttons.docs":        {"使用文档", "Documentation"},

	"stages.stageStart":      {"准备中...", "Preparing..."},
	"stages.stageDimensions": {"生成维度表 (大区/省/市/区县)...", "Generating dimensions (region/province/city/district)..."},
	"stages.stageProducts":   {"生成产品表...", "Generating products..."},
	"stages.stageStores":     {"生成门店表...", "Generating stores..."},
	"stages.stageCustomers":  {"生成客户表...", "Generating customers..."},
	"stages.stageOrders":     {"生成入库/订单数据...", "Generating inventory/orders..."},
	"stages.stageDone":       {"生成完成!", "Done!"},

	"result.title":     {"生成结果", "Result"},
	"result.products":  {"产品", "Products"},
	"result.stores":    {"门店", "Stores"},
	"result.customers": {"客户", "Customers"},
	"result.inventory": {"入库", "Inventory"},
	"result.orders":    {"订单主", "Orders"},
	"result.orderItem": {"订单子", "Order Items"},
	"result.rows":      {"行", "rows"},
	"result.colTable":  {"表名称", "Table"},
	"result.colRows":   {"行数", "Rows"},
	"result.colSize":   {"大小", "Size"},

	"date.today":      {"今日", "Today"},
	"date.yearStart":  {"年初", "Year Start"},
	"date.monthStart": {"月初", "Month Start"},
	"date.monthEnd":   {"月末", "Month End"},
	"date.yearEnd":    {"年末", "Year End"},

	"msg.chooseDirFirst":     {"请先指定数据存放目录", "Please choose an output directory first"},
	"msg.invalidRange":       {"结束日期必须晚于开始日期, 且窗口至少 60 天", "End date must be after start date, window at least 60 days"},
	"msg.generating":         {"正在生成, 请稍候...", "Generating, please wait..."},
	"msg.generateSuccess":    {"示例数据生成完成", "Sample data generated"},
	"msg.incrementalSuccess": {"增量更新完成", "Incremental update completed"},
	"msg.noBaseData":         {`目录中没有基础数据, 请先点击 "生成示例数据"`, `No base data in directory. Click "Generate Sample Data" first`},
	"msg.dateConflict":       {"增量日期区间与现有数据冲突: {range}, 请调整日期", "Incremental range conflicts with existing data: {range}. Adjust dates"},
	"msg.incStartMismatch":   {"增量更新的开始日期应为 {date} (事实表截止日 + 1)", "Incremental start date must be {date} (fact data cutoff + 1)"},
	"info.cutoff":            {"事实数据截止 {date}", "Data through {date}"},
	"msg.failed":             {"操作失败: {msg}", "Operation failed: {msg}"},
	"msg.productRange":       {"产品数量需在 1 - 10000 之间", "Products must be between 1 and 10000"},
	"msg.storeRange":         {"门店数量需在 1 - 10000 之间", "Stores must be between 1 and 10000"},
	"msg.inventoryRange":     {"入库周期需在 5 - 180 之间", "Inventory cycle must be between 5 and 180"},
	"msg.overwriteTitle":     {"确认覆盖现有数据?", "Overwrite existing data?"},
	"msg.overwriteContent":   {"当前目录已存在示例数据, 生成将覆盖原有数据, 是否继续?", "The output directory already contains sample data. Generating will overwrite it. Continue?"},
	"msg.overwriteConfirm":   {"覆盖生成", "Overwrite"},
	"msg.overwriteCancel":    {"取消", cancelTextEnUS},
	"msg.clearTitle":         {"确认清空数据?", "Clear generated data?"},
	"msg.clearContent":       {"将删除当前目录中已生成的全部数据 (data 与 pbip 子目录, 目录本身保留), 是否继续?", "This will delete all generated data in the directory (data and pbip subfolders; the folder itself is kept). Continue?"},
	"msg.clearConfirm":       {"清空", "Clear"},
	"msg.clearCancel":        {"取消", cancelTextEnUS},
	"msg.clearSuccess":       {"数据已清空", "Data cleared"},
}
