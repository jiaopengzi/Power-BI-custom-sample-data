/**
 * FilePath    : frontend/vitest.config.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : Vitest 测试配置.
 */

import { fileURLToPath } from "node:url"

import type { ConfigEnv, UserConfig as ViteUserConfig, UserConfigFnObject } from "vite"
import { configDefaults, defineConfig, mergeConfig } from "vitest/config"

import viteConfig from "./vite.config.ts"

/**
 * resolveViteConfig 解析 vite.config 的导出, 兼容对象式与函数式配置, 供 Vitest 合并。
 * @param configEnv - Vite 配置环境。
 * @returns 可供 mergeConfig 使用的 Vite 配置对象。
 */
const resolveViteConfig = async (configEnv: ConfigEnv): Promise<ViteUserConfig> => {
    if (typeof viteConfig === "function") {
        return (viteConfig as UserConfigFnObject)(configEnv)
    }
    return viteConfig
}

export default defineConfig(async (configEnv) => {
    const baseConfig = await resolveViteConfig(configEnv)

    return mergeConfig(baseConfig, {
        test: {
            setupFiles: [fileURLToPath(new URL("./vitest.setup.ts", import.meta.url))],
            environment: "jsdom",
            exclude: [...configDefaults.exclude, "**/node_modules/**", "**/dist/**", "**/wailsjs/**"],
            root: fileURLToPath(new URL("./", import.meta.url)),
            server: {
                deps: {
                    inline: ["element-plus"],
                },
            },
        },
    })
})
