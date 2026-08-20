<!--
  FilePath    : frontend/src/App.vue
  Author      : jiaopengzi
  Blog        : https://jiaopengzi.com
  Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
  Description : 主界面, 生成/增量/目录/语言 交互.
-->
<script setup lang="ts">
import { computed, onMounted, reactive, ref } from "vue"
import { useI18n } from "vue-i18n"
import { ElMessage } from "element-plus"
import zhCn from "element-plus/es/locale/lang/zh-cn"
import en from "element-plus/es/locale/lang/en"

import type { GenParams, GenResult } from "@/api/types"
import type { LocaleKey } from "@/i18n/types"
import { defaultOutputDir, generateSample, incrementalUpdate, isWails, onProgress, openOutputDir, selectOutputDir } from "@/api/wails"
import logoUrl from "@/assets/logo.svg"

const { t, locale } = useI18n()

/** docsUrl 使用文档地址, 后续可替换。 */
const docsUrl = "https://jiaopengzi.com/"

/**
 * formatDate 将日期格式化为 YYYY-MM-DD。
 * @param d - 日期对象。
 * @returns 格式化字符串。
 */
function formatDate(d: Date): string {
    const y = d.getFullYear()
    const m = String(d.getMonth() + 1).padStart(2, "0")
    const day = String(d.getDate()).padStart(2, "0")
    return `${y}-${m}-${day}`
}

const today = new Date()
const start = new Date(today.getTime() - 1600 * 24 * 3600 * 1000)

// 表单参数, 默认值沿用原 VBA 表单 (产品 200, 门店 5, 入库 14)。
const form = reactive<GenParams>({
    outputDir: "",
    locale: "zh-cn",
    productCount: 200,
    storeCount: 5,
    inventoryCycle: 14,
    startDate: formatDate(start),
    endDate: formatDate(today),
})

const running = ref(false)
const percent = ref(0)
const stageText = ref("")
const result = ref<GenResult | null>(null)

// Element Plus 组件语言随界面语言切换。
const elLocale = computed(() => (locale.value === "en-us" ? en : zhCn))

/**
 * changeLocale 切换界面与产物语言。
 * @param val - 目标语言标识。
 */
function changeLocale(val: LocaleKey): void {
    locale.value = val
    form.locale = val
}

/**
 * chooseDir 弹出目录选择对话框并写回存放目录。
 */
async function chooseDir(): Promise<void> {
    if (!isWails()) {
        ElMessage.warning(t("msg.chooseDirFirst"))
        return
    }
    const dir = await selectOutputDir()
    if (dir) {
        form.outputDir = dir
    }
}

/**
 * openDir 在文件管理器中打开当前存放目录。
 */
async function openDir(): Promise<void> {
    if (form.outputDir) {
        await openOutputDir(form.outputDir)
    }
}

/**
 * validate 校验参数, 不合法时提示并返回 false。
 * @returns 校验通过返回 true。
 */
function validate(): boolean {
    if (!form.outputDir) {
        ElMessage.warning(t("msg.chooseDirFirst"))
        return false
    }
    if (form.productCount < 1 || form.productCount > 2000) {
        ElMessage.error(t("msg.productRange"))
        return false
    }
    if (form.storeCount < 1 || form.storeCount > 400) {
        ElMessage.error(t("msg.storeRange"))
        return false
    }
    if (form.inventoryCycle < 5 || form.inventoryCycle > 20) {
        ElMessage.error(t("msg.inventoryRange"))
        return false
    }
    if (form.endDate <= form.startDate) {
        ElMessage.error(t("msg.invalidRange"))
        return false
    }
    return true
}

/**
 * run 执行全量生成或增量更新。
 * @param kind - full 全量, inc 增量。
 */
async function run(kind: "full" | "inc"): Promise<void> {
    if (running.value || !validate()) {
        return
    }
    if (!isWails()) {
        ElMessage.warning(t("msg.chooseDirFirst"))
        return
    }
    running.value = true
    percent.value = 0
    stageText.value = t("stages.stageStart")
    result.value = null
    try {
        const res = kind === "full" ? await generateSample({ ...form }) : await incrementalUpdate({ ...form })
        switch (res.code) {
            case "ok":
                result.value = res.result
                ElMessage.success(kind === "full" ? t("msg.generateSuccess") : t("msg.incrementalSuccess"))
                break
            case "no_base_data":
                ElMessage.warning(t("msg.noBaseData"))
                break
            case "date_conflict":
                ElMessage.error(t("msg.dateConflict", { range: res.message }))
                break
            default:
                ElMessage.error(t("msg.failed", { msg: res.message }))
        }
    } catch (e) {
        ElMessage.error(t("msg.failed", { msg: String(e) }))
    } finally {
        running.value = false
    }
}

onMounted(async () => {
    if (isWails()) {
        form.outputDir = await defaultOutputDir()
    }
    onProgress((payload) => {
        const p = payload as { percent: number; stage: string }
        percent.value = Math.round(p.percent)
        stageText.value = t(`stages.${p.stage}`)
    })
})
</script>

<template>
    <el-config-provider :locale="elLocale">
        <div class="pbi">
            <header class="pbi__header">
                <div class="pbi__brand">
                    <img :src="logoUrl" class="pbi__logo" alt="logo" />
                    <h1 class="pbi__title">{{ t("subtitle") }}</h1>
                </div>
                <el-select :model-value="locale" class="pbi__locale" @change="changeLocale">
                    <el-option label="简体中文" value="zh-cn" />
                    <el-option label="English" value="en-us" />
                </el-select>
            </header>

            <el-card class="pbi__card" shadow="never">
                <el-form label-position="top">
                    <div class="pbi__row">
                        <el-form-item :label="t('form.startDate')" class="pbi__col">
                            <el-date-picker v-model="form.startDate" type="date" value-format="YYYY-MM-DD" :disabled="running" />
                        </el-form-item>
                        <el-form-item :label="t('form.endDate')" class="pbi__col">
                            <el-date-picker v-model="form.endDate" type="date" value-format="YYYY-MM-DD" :disabled="running" />
                        </el-form-item>
                    </div>

                    <div class="pbi__row">
                        <el-form-item :label="t('form.productCount')" class="pbi__col">
                            <el-input-number v-model="form.productCount" :min="1" :max="2000" :disabled="running" />
                        </el-form-item>
                        <el-form-item :label="t('form.storeCount')" class="pbi__col">
                            <el-input-number v-model="form.storeCount" :min="1" :max="400" :disabled="running" />
                        </el-form-item>
                        <el-form-item :label="t('form.inventoryCycle')" class="pbi__col">
                            <el-input-number v-model="form.inventoryCycle" :min="5" :max="20" :disabled="running" />
                        </el-form-item>
                    </div>

                    <el-form-item :label="t('form.outputDir')">
                        <div class="pbi__dir">
                            <el-input v-model="form.outputDir" readonly :placeholder="t('form.outputDirPlaceholder')" />
                            <el-button :disabled="running" @click="chooseDir">{{ t("buttons.chooseDir") }}</el-button>
                            <el-button :disabled="running || !form.outputDir" @click="openDir">
                                {{ t("buttons.openDir") }}
                            </el-button>
                        </div>
                    </el-form-item>

                    <div class="pbi__actions">
                        <el-button type="primary" :loading="running" @click="run('full')">
                            {{ t("buttons.generate") }}
                        </el-button>
                        <el-button :loading="running" @click="run('inc')">
                            {{ t("buttons.incremental") }}
                        </el-button>
                    </div>
                </el-form>

                <div v-if="running || percent > 0" class="pbi__progress">
                    <el-progress :percentage="percent" :stroke-width="16" />
                    <span class="pbi__stage">{{ stageText }}</span>
                </div>

                <el-descriptions v-if="result" class="pbi__result" :title="t('result.title')" :column="3" border>
                    <el-descriptions-item :label="t('result.products')"> {{ result.products }} {{ t("result.rows") }} </el-descriptions-item>
                    <el-descriptions-item :label="t('result.stores')"> {{ result.stores }} {{ t("result.rows") }} </el-descriptions-item>
                    <el-descriptions-item :label="t('result.customers')"> {{ result.customers }} {{ t("result.rows") }} </el-descriptions-item>
                    <el-descriptions-item :label="t('result.inventory')"> {{ result.inventory }} {{ t("result.rows") }} </el-descriptions-item>
                    <el-descriptions-item :label="t('result.orders')"> {{ result.orders }} {{ t("result.rows") }} </el-descriptions-item>
                    <el-descriptions-item :label="t('result.orderItem')"> {{ result.orderItem }} {{ t("result.rows") }} </el-descriptions-item>
                </el-descriptions>
            </el-card>

            <footer class="pbi__footer">
                <el-link type="primary" :href="docsUrl" target="_blank">{{ t("buttons.docs") }}</el-link>
            </footer>
        </div>
    </el-config-provider>
</template>

<style scoped lang="scss">
.pbi {
    display: flex;
    flex-direction: column;
    height: 100%;
    padding: 20px 28px;
    gap: 16px;

    &__header {
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    &__brand {
        display: flex;
        align-items: center;
        gap: 12px;
    }

    &__logo {
        width: 36px;
        height: 36px;
        border-radius: 8px;
    }

    &__title {
        margin: 0;
        font-size: 20px;
        color: var(--pbi-primary);
    }

    &__locale {
        width: 140px;
    }

    &__card {
        flex: 1;
        overflow: auto;
        border-radius: 12px;
    }

    &__row {
        display: flex;
        gap: 20px;

        .pbi__col {
            flex: 1;
        }
    }

    &__dir {
        display: flex;
        gap: 10px;
        width: 100%;
    }

    &__actions {
        display: flex;
        gap: 12px;
        margin-top: 4px;
    }

    &__progress {
        display: flex;
        align-items: center;
        gap: 14px;
        margin-top: 20px;
    }

    &__stage {
        font-size: 13px;
        color: #606266;
        white-space: nowrap;
    }

    &__result {
        margin-top: 22px;
    }

    &__footer {
        display: flex;
        justify-content: center;
    }
}
</style>
