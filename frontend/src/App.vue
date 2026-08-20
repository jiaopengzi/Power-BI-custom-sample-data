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
import { ElMessage, ElMessageBox } from "element-plus"
import zhCn from "element-plus/es/locale/lang/zh-cn"
import en from "element-plus/es/locale/lang/en"

import type { GenParams, GenResult } from "@/api/types"
import type { LocaleKey } from "@/i18n/types"
import { defaultOutputDir, generateSample, hasData, incrementalUpdate, isWails, onProgress, openOutputDir, selectOutputDir } from "@/api/wails"
import logoUrl from "@/assets/logo.svg"

const { t, locale } = useI18n()

/** docsUrl 使用文档地址, 后续可替换。 */
const docsUrl = "https://jiaopengzi.com/?post_id=19051044919050241"

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
    // 全量生成前, 若目录已有基础数据则提示会覆盖, 确认后再继续。
    if (kind === "full" && (await hasData(form.outputDir))) {
        try {
            await ElMessageBox.confirm(t("msg.overwriteContent"), t("msg.overwriteTitle"), {
                type: "warning",
                confirmButtonText: t("msg.overwriteConfirm"),
                cancelButtonText: t("msg.overwriteCancel"),
            })
        } catch {
            return
        }
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
            <div class="pbi__aurora" aria-hidden="true"></div>

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

            <main class="pbi__panel">
                <el-form label-position="top" class="pbi__form">
                    <div class="pbi__grid pbi__grid--2">
                        <el-form-item :label="t('form.startDate')">
                            <el-date-picker v-model="form.startDate" type="date" value-format="YYYY-MM-DD" :disabled="running" class="pbi__full" />
                        </el-form-item>
                        <el-form-item :label="t('form.endDate')">
                            <el-date-picker v-model="form.endDate" type="date" value-format="YYYY-MM-DD" :disabled="running" class="pbi__full" />
                        </el-form-item>
                    </div>

                    <div class="pbi__grid pbi__grid--3">
                        <el-form-item :label="t('form.productCount')">
                            <el-input-number v-model="form.productCount" :min="1" :max="2000" :disabled="running" controls-position="right" class="pbi__full" />
                        </el-form-item>
                        <el-form-item :label="t('form.storeCount')">
                            <el-input-number v-model="form.storeCount" :min="1" :max="400" :disabled="running" controls-position="right" class="pbi__full" />
                        </el-form-item>
                        <el-form-item :label="t('form.inventoryCycle')">
                            <el-input-number v-model="form.inventoryCycle" :min="5" :max="20" :disabled="running" controls-position="right" class="pbi__full" />
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
                        <el-button type="primary" class="pbi__btn-main" :loading="running" @click="run('full')">
                            {{ t("buttons.generate") }}
                        </el-button>
                        <el-button class="pbi__btn-sub" :loading="running" @click="run('inc')">
                            {{ t("buttons.incremental") }}
                        </el-button>
                    </div>
                </el-form>

                <transition name="pbi-rise">
                    <div v-if="running || percent > 0" class="pbi-progress">
                        <div class="pbi-progress__head">
                            <span class="pbi-progress__stage">
                                <span class="pbi-progress__dot" :class="{ 'is-active': running }"></span>
                                {{ stageText }}
                            </span>
                            <span class="pbi-progress__percent">{{ percent }}<i>%</i></span>
                        </div>
                        <div class="pbi-progress__track">
                            <div class="pbi-progress__fill" :class="{ 'is-active': running }" :style="{ width: percent + '%' }"></div>
                        </div>
                    </div>
                </transition>

                <transition name="pbi-rise">
                    <el-descriptions v-if="result" class="pbi__result" :title="t('result.title')" :column="3" border>
                        <el-descriptions-item :label="t('result.products')"> {{ result.products }} {{ t("result.rows") }} </el-descriptions-item>
                        <el-descriptions-item :label="t('result.stores')"> {{ result.stores }} {{ t("result.rows") }} </el-descriptions-item>
                        <el-descriptions-item :label="t('result.customers')"> {{ result.customers }} {{ t("result.rows") }} </el-descriptions-item>
                        <el-descriptions-item :label="t('result.inventory')"> {{ result.inventory }} {{ t("result.rows") }} </el-descriptions-item>
                        <el-descriptions-item :label="t('result.orders')"> {{ result.orders }} {{ t("result.rows") }} </el-descriptions-item>
                        <el-descriptions-item :label="t('result.orderItem')"> {{ result.orderItem }} {{ t("result.rows") }} </el-descriptions-item>
                    </el-descriptions>
                </transition>
            </main>

            <footer class="pbi__footer">
                <el-link type="primary" :href="docsUrl" target="_blank" :underline="false">{{ t("buttons.docs") }}</el-link>
            </footer>
        </div>
    </el-config-provider>
</template>

<style scoped lang="scss">
.pbi {
    position: relative;
    display: flex;
    flex-direction: column;
    height: 100%;
    padding: 20px 34px 14px;
    gap: 14px;
    overflow: hidden;
    background:
        radial-gradient(120% 120% at 100% 0%, rgba(200, 152, 40, 0.08), transparent 45%),
        radial-gradient(120% 120% at 0% 100%, rgba(30, 40, 88, 0.09), transparent 50%), var(--pbi-bg);

    &__aurora {
        position: absolute;
        inset: 0;
        pointer-events: none;
        z-index: 0;

        &::before,
        &::after {
            content: "";
            position: absolute;
            border-radius: 50%;
            filter: blur(80px);
        }

        &::before {
            width: 320px;
            height: 320px;
            top: -130px;
            right: -70px;
            background: radial-gradient(circle, rgba(200, 152, 40, 0.3), transparent 70%);
        }

        &::after {
            width: 360px;
            height: 360px;
            bottom: -170px;
            left: -110px;
            background: radial-gradient(circle, rgba(30, 40, 88, 0.24), transparent 70%);
        }
    }

    &__header {
        position: relative;
        z-index: 1;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    &__brand {
        display: flex;
        align-items: center;
        gap: 14px;
    }

    &__logo {
        width: 44px;
        height: 44px;
        border-radius: 12px;
        box-shadow: 0 8px 20px -8px rgba(30, 40, 88, 0.5);
    }

    &__title {
        margin: 0;
        font-size: 21px;
        font-weight: 700;
        letter-spacing: 0.5px;
        color: var(--pbi-primary);
    }

    &__locale {
        width: 140px;
    }

    &__panel {
        position: relative;
        z-index: 1;
        flex: 1;
        overflow: auto;
        padding: 24px 34px 26px;
        background: var(--pbi-card);
        border: 1px solid rgba(30, 40, 88, 0.06);
        border-radius: 18px;
        box-shadow: 0 24px 60px -32px rgba(30, 40, 88, 0.35);

        &::before {
            content: "";
            position: absolute;
            inset: 0 0 auto 0;
            height: 3px;
            border-radius: 18px 18px 0 0;
            background: linear-gradient(90deg, var(--pbi-primary), var(--pbi-secondary));
        }
    }

    &__grid {
        display: grid;
        gap: 4px 48px;

        &--2 {
            grid-template-columns: repeat(2, 1fr);
        }

        &--3 {
            grid-template-columns: repeat(3, 1fr);
        }
    }

    &__full {
        width: 100%;
    }

    &__dir {
        display: flex;
        gap: 10px;
        width: 100%;
    }

    &__actions {
        display: flex;
        gap: 14px;
        margin-top: 10px;
    }

    &__btn-main {
        min-width: 150px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }

    &__btn-sub {
        min-width: 120px;
    }

    &__result {
        margin-top: 18px;
    }

    &__footer {
        position: relative;
        z-index: 1;
        display: flex;
        justify-content: center;
    }
}

// 自定义进度区域, 主色 -> 金色渐变填充 + 流光动效。
.pbi-progress {
    margin-top: 18px;
    padding: 14px 20px;
    border-radius: 14px;
    background: linear-gradient(180deg, rgba(30, 40, 88, 0.045), rgba(30, 40, 88, 0.015));
    border: 1px solid rgba(30, 40, 88, 0.07);

    &__head {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        margin-bottom: 12px;
    }

    &__stage {
        display: inline-flex;
        align-items: center;
        gap: 9px;
        font-size: 13.5px;
        font-weight: 500;
        color: var(--pbi-primary);
    }

    &__dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: var(--pbi-secondary);

        &.is-active {
            animation: pbi-pulse 1.2s ease-in-out infinite;
        }
    }

    &__percent {
        font-size: 22px;
        font-weight: 700;
        color: var(--pbi-primary);
        font-variant-numeric: tabular-nums;

        i {
            margin-left: 2px;
            font-size: 13px;
            font-style: normal;
            color: var(--pbi-secondary);
        }
    }

    &__track {
        height: 10px;
        border-radius: 999px;
        background: rgba(30, 40, 88, 0.1);
        overflow: hidden;
    }

    &__fill {
        position: relative;
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, var(--pbi-primary), var(--pbi-secondary));
        transition: width 0.45s cubic-bezier(0.22, 1, 0.36, 1);

        &.is-active::after {
            content: "";
            position: absolute;
            inset: 0;
            background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.5), transparent);
            background-size: 200% 100%;
            animation: pbi-shine 1.4s linear infinite;
        }
    }
}

@keyframes pbi-shine {
    from {
        background-position: 200% 0;
    }

    to {
        background-position: -200% 0;
    }
}

@keyframes pbi-pulse {
    0%,
    100% {
        transform: scale(1);
        box-shadow: 0 0 0 0 rgba(200, 152, 40, 0.5);
    }

    50% {
        transform: scale(1.25);
        box-shadow: 0 0 0 5px rgba(200, 152, 40, 0);
    }
}

// 进度与结果的入场过渡。
.pbi-rise-enter-active {
    transition: all 0.4s cubic-bezier(0.22, 1, 0.36, 1);
}

.pbi-rise-enter-from {
    opacity: 0;
    transform: translateY(12px);
}

// Element Plus 细节微调, 使表单更轻盈统一。
:deep(.el-form-item__label) {
    padding-bottom: 4px !important;
    font-size: 12.5px;
    letter-spacing: 0.3px;
    color: #8a90a6;
}

:deep(.el-form-item) {
    margin-bottom: 16px;
}
</style>
