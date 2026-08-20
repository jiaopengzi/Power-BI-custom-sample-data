/**
 * FilePath    : frontend/env.d.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 全局类型声明 (Vue SFC 与 Wails window 注入).
 */

/// <reference types="vite/client" />

import type { WailsApp } from "./src/api/types"

declare module "*.vue" {
    import type { DefineComponent } from "vue"

    const component: DefineComponent<Record<string, never>, Record<string, never>, unknown>

    export default component
}

declare global {
    interface Window {
        /** go, Wails 注入的后端绑定命名空间。 */
        go?: { main?: { App?: WailsApp } }
        /** runtime, Wails 注入的运行时事件/系统能力。 */
        runtime?: {
            EventsOn: (event: string, cb: (data: unknown) => void) => void
        }
    }
}
