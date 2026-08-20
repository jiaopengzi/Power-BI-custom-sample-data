/**
 * FilePath    : frontend/src/i18n/types.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 国际化文案结构类型.
 */

/**
 * Messages 国际化文案结构, 用于约束各语言包字段一致。
 */
export interface Messages {
    title: string
    subtitle: string
    form: {
        startDate: string
        endDate: string
        productCount: string
        storeCount: string
        inventoryCycle: string
        outputDir: string
        outputDirPlaceholder: string
        localeLabel: string
    }
    buttons: {
        generate: string
        incremental: string
        chooseDir: string
        openDir: string
        docs: string
    }
    stages: {
        stageStart: string
        stageDimensions: string
        stageProducts: string
        stageStores: string
        stageCustomers: string
        stageOrders: string
        stageDone: string
    }
    result: {
        title: string
        products: string
        stores: string
        customers: string
        inventory: string
        orders: string
        orderItem: string
        rows: string
    }
    msg: {
        chooseDirFirst: string
        invalidRange: string
        generating: string
        generateSuccess: string
        incrementalSuccess: string
        noBaseData: string
        dateConflict: string
        failed: string
        productRange: string
        storeRange: string
        inventoryRange: string
    }
}

/** LocaleKey 支持的语言标识。 */
export type LocaleKey = "zh-cn" | "en-us"
