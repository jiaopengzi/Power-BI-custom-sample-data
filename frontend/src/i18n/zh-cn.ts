/**
 * FilePath    : frontend/src/i18n/zh-cn.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 简体中文文案.
 */

import type { Messages } from "./types"

// 简体中文文案。
const messages: Messages = {
    title: "Power BI Sample Data Generator",
    subtitle: "Power BI 示例数据生成器",
    form: {
        startDate: "开始日期",
        endDate: "结束日期",
        productCount: "产品数量",
        storeCount: "门店数量",
        inventoryCycle: "入库周期",
        outputDir: "存放目录",
        outputDirPlaceholder: "请先指定数据存放目录",
        localeLabel: "语言",
    },
    buttons: {
        generate: "生成示例数据",
        incremental: "增量更新",
        chooseDir: "指定存放目录",
        openDir: "打开存放目录",
        docs: "使用文档",
    },
    stages: {
        stageStart: "准备中...",
        stageDimensions: "生成维度表 (大区/省/市/区县)...",
        stageProducts: "生成产品表...",
        stageStores: "生成门店表...",
        stageCustomers: "生成客户表...",
        stageOrders: "生成入库/订单数据...",
        stageDone: "生成完成!",
    },
    result: {
        title: "生成结果",
        products: "产品",
        stores: "门店",
        customers: "客户",
        inventory: "入库",
        orders: "订单主",
        orderItem: "订单子",
        rows: "行",
    },
    msg: {
        chooseDirFirst: "请先指定数据存放目录",
        invalidRange: "结束日期必须晚于开始日期, 且窗口至少 60 天",
        generating: "正在生成, 请稍候...",
        generateSuccess: "示例数据生成完成",
        incrementalSuccess: "增量更新完成",
        noBaseData: '目录中没有基础数据, 请先点击 "生成示例数据"',
        dateConflict: "增量日期区间与现有数据冲突: {range}, 请调整日期",
        failed: "操作失败: {msg}",
        productRange: "产品数量需在 1 - 2000 之间",
        storeRange: "门店数量需在 1 - 400 之间",
        inventoryRange: "入库周期需在 5 - 20 之间",
    },
}

export default messages
