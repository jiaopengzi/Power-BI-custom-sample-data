/**
 * FilePath    : frontend/src/main.ts
 * Author      : jiaopengzi
 * Blog        : https://jiaopengzi.com
 * Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
 * Description : 前端入口, 注册 Element Plus/图标/i18n 并挂载应用.
 */

import { createApp } from "vue"
import ElementPlus from "element-plus"
import "element-plus/dist/index.css"
import * as ElementPlusIconsVue from "@element-plus/icons-vue"

import { i18n } from "@/i18n"
import App from "@/App.vue"
import "@/style.scss"

const app = createApp(App)

// 注册全部 Element Plus 图标。
for (const [key, component] of Object.entries(ElementPlusIconsVue)) {
    app.component(key, component)
}

app.use(ElementPlus)
app.use(i18n)
app.mount("#app")
