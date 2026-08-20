/**
 * FilePath    : frontend/src/i18n/index.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : i18n 实例 (默认 zh-cn).
 */

import { createI18n } from "vue-i18n"

import zhCN from "./zh-cn"
import enUS from "./en-us"

// i18n 实例, 默认使用简体中文。
// 消息结构已由 zh-cn/en-us 的 Messages 类型约束, 此处对 vue-i18n 放宽类型以避免其递归消息类型冲突。
export const i18n = createI18n({
    legacy: false,
    locale: "zh-cn",
    fallbackLocale: "zh-cn",
    messages: {
        "zh-cn": zhCN,
        "en-us": enUS,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any,
})
