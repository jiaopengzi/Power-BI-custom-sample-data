/**
 * FilePath    : frontend/vitest.setup.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : Vitest 全局初始化, 为测试环境注入 Wails 运行时桩对象.
 */

import { vi } from "vitest"

// Wails 运行时在浏览器/jsdom 中不存在, 注入空桩避免测试报错。
vi.stubGlobal("runtime", { EventsOn: () => {} })
