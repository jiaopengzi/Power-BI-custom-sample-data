# FilePath    : run.ps1
# Author      : jiaopengzi
# Blog        : https://jiaopengzi.com
# Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
# Description : 该脚本用于本地开发与 CI/CD, 统一管理 Go 后端与 Vue3 前端的格式化, 静态检查,
#               测试, 以及 Wails 桌面应用的开发与打包. 交互式运行时根据菜单选择编号;
#               CI 场景通过 -Choice 参数非交互执行, 例如 .\run.ps1 -Choice 1.

param(
    [string]$Choice
)

$ErrorActionPreference = "Stop"

# 可执行文件名称 (与 wails.json 的 outputfilename 保持一致)
$BINARY = "PowerBISampleDataGenerator"

# 前端目录
$FRONTEND = ".\frontend"

# 显示菜单
Write-Host ""
Write-Host "请选择需要执行的命令："
Write-Host "  0 - 安装/更新 Wails CLI"
Write-Host "  1 - CI 全流程 (Go 检查+测试, 前端检查+测试, Wails 打包)"
Write-Host "  2 - 格式化 Go 代码"
Write-Host "  3 - Go 静态检查 (go vet + golangci-lint)"
Write-Host "  4 - Go 单元测试"
Write-Host "  5 - 安装前端依赖 (自动同步锁文件)"
Write-Host "  6 - lint 前端代码"
Write-Host "  7 - 前端单元测试"
Write-Host "  8 - 格式化前端代码"
Write-Host "  9 - 构建前端 (pnpm build)"
Write-Host " 10 - Wails 开发模式 (wails dev)"
Write-Host " 11 - Wails 打包 (Windows amd64)"
Write-Host " 12 - 清理构建产物"
Write-Host ""

# 非交互 (CI) 场景使用 -Choice, 否则交互式读取
if ($PSBoundParameters.ContainsKey("Choice") -and $Choice -ne "") {
    $selected = $Choice
} else {
    $selected = Read-Host "请输入编号选择对应的操作"
}
Write-Host ""

# ensureFrontendDist 确保 frontend/dist 存在, 避免 go vet/test 时 go:embed 因缺少目录而失败.
# main.go 使用 //go:embed all:frontend/dist, 未构建前端时该目录不存在, 会导致 Go 命令报错.
function ensureFrontendDist {
    $distPath = "$FRONTEND\dist"
    if (-not (Test-Path "$distPath\index.html")) {
        Write-Host "📁 frontend/dist 不存在, 创建占位文件以满足 go:embed..." -ForegroundColor Yellow
        New-Item -ItemType Directory -Force -Path $distPath | Out-Null
        Set-Content -Path "$distPath\index.html" -Value "<!-- placeholder for go:embed, run frontend build to replace -->" -Encoding UTF8
    }
}

# installWails 安装或更新 Wails CLI
function installWails {
    Write-Host "📥 安装 Wails CLI..." -ForegroundColor Cyan
    go install github.com/wailsapp/wails/v2/cmd/wails@latest
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ Wails CLI 安装失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ Wails CLI 安装完成"
}

# formatGoCode 格式化 Go 代码
function formatGoCode {
    Write-Host "🔨 格式化 Go 代码..." -ForegroundColor Cyan
    gofmt -w .
    go mod tidy
    Write-Host "✅ Go 代码格式化完成"
}

# goLint Go 静态检查
function goLint {
    Write-Host "🔍 Go 静态检查..." -ForegroundColor Cyan
    ensureFrontendDist
    go vet ./...
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ go vet 失败" -ForegroundColor Red; exit 1 }
    golangci-lint run ./...
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ golangci-lint 失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ Go 静态检查完成"
}

# testGo Go 单元测试
function testGo {
    Write-Host "🧪 Go 单元测试..." -ForegroundColor Cyan
    ensureFrontendDist
    go test ./...
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ Go 测试失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ Go 测试完成"
}

# installFrontendWithAutoLockfileSync 安装前端依赖, 并在锁文件过期时自动同步后重试.
function installFrontendWithAutoLockfileSync {
    $isCI = $env:CI -eq "true"
    $installOutput = & pnpm -C $FRONTEND install --frozen-lockfile 2>&1
    $installExitCode = $LASTEXITCODE

    if ($installOutput) {
        $installOutput | Out-Host
    }

    if ($installExitCode -eq 0) {
        return
    }

    $installOutputText = $installOutput | Out-String
    if ($installOutputText -notmatch "ERR_PNPM_OUTDATED_LOCKFILE") {
        Write-Host "❌ 前端依赖安装失败" -ForegroundColor Red
        exit 1
    }

    if ($isCI) {
        Write-Host "❌ 检测到 pnpm-lock.yaml 已过期, CI 环境不会自动同步, 请先更新锁文件并提交." -ForegroundColor Red
        exit 1
    }

    Write-Host "📝 检测到 pnpm-lock.yaml 已过期, 正在自动同步锁文件..." -ForegroundColor Yellow
    pnpm -C $FRONTEND install --lockfile-only
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ 前端锁文件同步失败" -ForegroundColor Red; exit 1 }

    pnpm -C $FRONTEND install --frozen-lockfile
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ 前端依赖安装失败" -ForegroundColor Red; exit 1 }
}

# installFrontend 安装前端依赖
function installFrontend {
    Write-Host "📥 安装前端依赖..." -ForegroundColor Cyan
    installFrontendWithAutoLockfileSync
    Write-Host "✅ 前端依赖安装完成"
}

# lintFrontend lint 前端代码
function lintFrontend {
    Write-Host "🔍 lint 前端代码..." -ForegroundColor Cyan
    pnpm -C $FRONTEND lint
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ 前端 lint 失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ 前端 lint 完成"
}

# testFrontend 前端单元测试
function testFrontend {
    Write-Host "🧪 前端单元测试..." -ForegroundColor Cyan
    pnpm -C $FRONTEND test
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ 前端测试失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ 前端测试完成"
}

# formatFrontend 格式化前端代码
function formatFrontend {
    Write-Host "🔨 格式化前端代码..." -ForegroundColor Cyan
    pnpm -C $FRONTEND fmt
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ 前端格式化失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ 前端格式化完成"
}

# buildFrontend 构建前端
function buildFrontend {
    Write-Host "🔨 构建前端..." -ForegroundColor Cyan
    pnpm -C $FRONTEND build
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ 前端构建失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ 前端构建完成"
}

# wailsDev Wails 开发模式
function wailsDev {
    Write-Host "🚀 启动 Wails 开发模式..." -ForegroundColor Cyan
    wails dev
}

# wailsBuild Wails 打包 (Windows amd64)
function wailsBuild {
    Write-Host "📦 Wails 打包 (Windows amd64)..." -ForegroundColor Cyan
    wails build -platform windows/amd64 -tags production -o "$BINARY.exe"
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ Wails 打包失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ 打包完成, 产物位于 build\bin\$BINARY.exe"
}

# clean 清理构建产物
function clean {
    Write-Host "🧹 清理构建产物..." -ForegroundColor Cyan
    if (Test-Path ".\build\bin") { Remove-Item -Recurse -Force ".\build\bin" }
    if (Test-Path "$FRONTEND\dist") { Remove-Item -Recurse -Force "$FRONTEND\dist" }
    go clean
    Write-Host "✅ 清理完成"
}

# ci CI 全流程: 后端检查测试 + 前端检查测试 + 打包
function ci {
    goLint
    testGo
    installFrontend
    lintFrontend
    testFrontend
    wailsBuild
    Write-Host "✅ CI 全流程执行完毕" -ForegroundColor Green
}

# switch 放到最后, 执行用户选择的操作
switch ($selected) {
    0 { installWails }
    1 { ci }
    2 { formatGoCode }
    3 { goLint }
    4 { testGo }
    5 { installFrontend }
    6 { lintFrontend }
    7 { testFrontend }
    8 { formatFrontend }
    9 { buildFrontend }
    10 { wailsDev }
    11 { wailsBuild }
    12 { clean }
    default { Write-Host "❌ 无效的选择" -ForegroundColor Red }
}
