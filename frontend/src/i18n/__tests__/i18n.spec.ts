/**
 * FilePath    : frontend/src/i18n/__tests__/i18n.spec.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 多语言字段一致性测试.
 */

import { describe, expect, it } from "vitest"

import zhCN from "../zh-cn"
import enUS from "../en-us"

/**
 * collectKeys 递归收集对象的全部叶子键路径, 用于比对多语言字段一致性。
 * @param obj - 待收集的对象。
 * @param prefix - 当前路径前缀。
 * @returns 排序后的键路径数组。
 */
function collectKeys(obj: Record<string, unknown>, prefix = ""): string[] {
    const keys: string[] = []
    for (const [k, v] of Object.entries(obj)) {
        const path = prefix ? `${prefix}.${k}` : k
        if (v && typeof v === "object") {
            keys.push(...collectKeys(v as Record<string, unknown>, path))
        } else {
            keys.push(path)
        }
    }
    return keys.toSorted()
}

describe("i18n", () => {
    it("zh-cn 与 en-us 字段完全一致", () => {
        expect(collectKeys(zhCN as unknown as Record<string, unknown>)).toEqual(collectKeys(enUS as unknown as Record<string, unknown>))
    })
})
