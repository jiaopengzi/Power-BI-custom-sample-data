/**
 * FilePath    : frontend/src/api/types.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 前后端交互的参数与结果类型定义.
 */

/**
 * GenParams 前端提交给后端的生成参数, 日期为 YYYY-MM-DD 字符串。
 */
export interface GenParams {
    /** outputDir, 数据存放目录。 */
    outputDir: string
    /** locale, 语言标识 zh-cn / en-us。 */
    locale: string
    /** productCount, 产品数量。 */
    productCount: number
    /** storeCount, 门店数量。 */
    storeCount: number
    /** inventoryCycle, 入库周期。 */
    inventoryCycle: number
    /** startDate, 数据窗口开始日期。 */
    startDate: string
    /** endDate, 数据窗口结束日期。 */
    endDate: string
}

/**
 * GenResult 生成结果各表行数统计。
 */
export interface GenResult {
    products: number
    stores: number
    customers: number
    inventory: number
    orders: number
    orderItem: number
}

/**
 * GenResponse 后端统一返回结构。
 */
export interface GenResponse {
    /** code, 结果码。 */
    code: "ok" | "no_base_data" | "date_conflict" | "error"
    /** message, 附加消息。 */
    message: string
    /** result, 行数统计。 */
    result: GenResult
}

/**
 * ProgressPayload 生成进度事件负载。
 */
export interface ProgressPayload {
    /** percent, 当前进度百分比 [0, 100]。 */
    percent: number
    /** stage, 当前阶段文案键。 */
    stage: string
}

/**
 * WailsApp Wails 注入到 window.go.main.App 的后端方法集合。
 */
export interface WailsApp {
    GenerateSample(params: GenParams): Promise<GenResponse>
    IncrementalUpdate(params: GenParams): Promise<GenResponse>
    SelectOutputDir(): Promise<string>
    OpenOutputDir(dir: string): Promise<void>
    HasData(dir: string): Promise<boolean>
    DefaultOutputDir(): Promise<string>
}
