/**
 * FilePath    : frontend/src/api/wails.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : Wails 后端绑定方法的封装.
 */
import type { GenParams, GenResponse, ProgressPayload, WailsApp } from "./types"

/**
 * app 返回 Wails 注入的后端绑定对象, 未运行在 Wails 中时为 undefined。
 * @returns 后端绑定对象或 undefined。
 */
function app(): WailsApp | undefined {
    return window.go?.main?.App
}

/**
 * isWails 判断是否运行在 Wails 桌面环境中。
 * @returns 处于 Wails 环境返回 true。
 */
export const isWails = (): boolean => !!app()

/**
 * generateSample 触发全量示例数据生成。
 * @param params - 生成参数。
 * @returns 后端返回的结果。
 */
export function generateSample(params: GenParams): Promise<GenResponse> {
    return app()!.GenerateSample(params)
}

/**
 * incrementalUpdate 触发增量更新。
 * @param params - 生成参数, 日期区间为增量窗口。
 * @returns 后端返回的结果。
 */
export function incrementalUpdate(params: GenParams): Promise<GenResponse> {
    return app()!.IncrementalUpdate(params)
}

/**
 * selectOutputDir 打开目录选择对话框。
 * @returns 选中的目录路径。
 */
export function selectOutputDir(): Promise<string> {
    return app()!.SelectOutputDir()
}

/**
 * openOutputDir 在系统文件管理器中打开目录。
 * @param dir - 目录路径。
 */
export function openOutputDir(dir: string): Promise<void> {
    return app()!.OpenOutputDir(dir)
}

/**
 * hasData 判断目录是否已存在基础数据。
 * @param dir - 目录路径。
 * @returns 存在返回 true。
 */
export function hasData(dir: string): Promise<boolean> {
    return app()!.HasData(dir)
}

/**
 * defaultOutputDir 获取默认数据存放目录。
 * @returns 默认目录路径。
 */
export function defaultOutputDir(): Promise<string> {
    return app()!.DefaultOutputDir()
}

/**
 * onProgress 订阅生成进度事件。
 * @param cb - 进度回调, 入参为进度负载。
 */
export function onProgress(cb: (payload: ProgressPayload) => void): void {
    window.runtime?.EventsOn("generate:progress", (data) => cb(data as ProgressPayload))
}
