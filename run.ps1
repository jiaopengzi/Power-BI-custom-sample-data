# FilePath    : run.ps1
# Author      : jiaopengzi
# Blog        : https://jiaopengzi.com
# Copyright   : Copyright (c) 2026 by jiaopengzi, All Rights Reserved.
# Description : 本地开发与 CI/CD 脚本, 统一管理原生 Go + Fyne 桌面应用的格式化, 静态检查,
#               测试与打包. 交互式运行时按菜单编号选择; CI 场景通过 -Choice 参数非交互执行,
#               例如 .\run.ps1 -Choice 1. 依赖 Go 工具链 + gcc (CGO) + fyne CLI.

param(
    [string]$Choice
)

$ErrorActionPreference = "Stop"

# 应用名 (与 cmd/pbicsd/FyneApp.toml 的 Name 保持一致, fyne package 以此命名 exe)
$APP_NAME = "Power BI Custom Sample Data"

# 可执行入口包
$ENTRY = ".\cmd\pbicsd"

# Fyne 依赖 CGO (底层 OpenGL/系统 GUI 绑定), 需本机具备 gcc; 显式开启避免默认关闭导致链接失败
$env:CGO_ENABLED = "1"

# 显示菜单
Write-Host ""
Write-Host "请选择需要执行的命令："
Write-Host "  0 - 安装/更新开发工具 (fyne CLI + golangci-lint)"
Write-Host "  1 - CI 全流程 (静态检查 + 测试 + 打包)"
Write-Host "  2 - 格式化 Go 代码 (gofmt + go mod tidy)"
Write-Host "  3 - Go 静态检查 (go vet + golangci-lint)"
Write-Host "  4 - Go 单元测试 (go test)"
Write-Host "  5 - 打包 Windows 桌面应用 (fyne package)"
Write-Host "  6 - 本地运行 (go run)"
Write-Host "  7 - 清理构建产物"
Write-Host ""

# 非交互 (CI) 场景使用 -Choice, 否则交互式读取
if ($PSBoundParameters.ContainsKey("Choice") -and $Choice -ne "") {
    $selected = $Choice
} else {
    $selected = Read-Host "请输入编号选择对应的操作"
}
Write-Host ""

# ensureGcc 确保 gcc 可用 (Fyne/CGO 必需):
#   1. 当前进程 PATH 已能解析 gcc -> 直接返回;
#   2. 否则自动探测常见安装位置 (scoop mingw / 常见 mingw64), 命中则注入本次进程 PATH
#      (会传递给子进程 go/fyne/gcc, 故后续 CGO 构建可正常链接);
#   3. 仍找不到才报错退出.
# 说明: scoop 的 mingw 不生成 shim, gcc 仅靠 `scoop\apps\mingw\current\bin` 挂到 PATH;
#       若运行本脚本的终端 PATH 过期或不含该目录, 直接用 Get-Command 会误报 "未找到".
function ensureGcc {
    if (Get-Command gcc -ErrorAction SilentlyContinue) { return }

    $candidates = @(
        (Join-Path $env:USERPROFILE 'scoop\apps\mingw\current\bin'),
        (Join-Path $env:USERPROFILE 'scoop\apps\mingw\current\mingw64\bin'),
        'C:\mingw64\bin',
        'C:\ProgramData\mingw64\mingw64\bin'
    )
    if ($env:SCOOP) { $candidates = @((Join-Path $env:SCOOP 'apps\mingw\current\bin')) + $candidates }

    foreach ($dir in $candidates) {
        if ($dir -and (Test-Path (Join-Path $dir 'gcc.exe'))) {
            $env:Path = "$dir;$env:Path"
            Write-Host "ℹ️  已自动定位 gcc 并加入本次 PATH: $dir" -ForegroundColor DarkGray
            return
        }
    }

    Write-Host "❌ 未找到 gcc, Fyne 依赖 CGO. 请安装 MinGW-w64 (如 scoop install mingw), 或重启终端使 PATH 生效." -ForegroundColor Red
    exit 1
}

# assertFyne 确认 fyne CLI 可用, 缺失则提示先执行选项 0 安装.
function assertFyne {
    if ($null -eq (Get-Command fyne -ErrorAction SilentlyContinue)) {
        Write-Host "❌ 未找到 fyne CLI, 请先执行 .\run.ps1 -Choice 0 安装 (并确保 GOPATH\bin 在 PATH 中)." -ForegroundColor Red
        exit 1
    }
}

# installTools 安装或更新开发工具 (fyne CLI + golangci-lint)
function installTools {
    Write-Host "📥 安装 fyne CLI..." -ForegroundColor Cyan
    go install fyne.io/tools/cmd/fyne@latest
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ fyne CLI 安装失败" -ForegroundColor Red; exit 1 }
    Write-Host "📥 安装 golangci-lint..." -ForegroundColor Cyan
    go install github.com/golangci/golangci-lint/v2/cmd/golangci-lint@latest
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ golangci-lint 安装失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ 开发工具安装完成 (请确保 $(go env GOPATH)\bin 已加入 PATH)"
}

# formatGoCode 格式化 Go 代码
function formatGoCode {
    Write-Host "🔨 格式化 Go 代码..." -ForegroundColor Cyan
    gofmt -w .
    go mod tidy
    Write-Host "✅ Go 代码格式化完成"
}

# goLint Go 静态检查 (go vet + golangci-lint)
function goLint {
    Write-Host "🔍 Go 静态检查..." -ForegroundColor Cyan
    ensureGcc
    go vet ./...
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ go vet 失败" -ForegroundColor Red; exit 1 }
    golangci-lint run ./...
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ golangci-lint 失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ Go 静态检查完成"
}

# testGo Go 单元测试
function testGo {
    Write-Host "🧪 Go 单元测试..." -ForegroundColor Cyan
    ensureGcc
    go test ./...
    if ($LASTEXITCODE -ne 0) { Write-Host "❌ Go 测试失败" -ForegroundColor Red; exit 1 }
    Write-Host "✅ Go 测试完成"
}

# packageApp 打包 Windows 桌面应用 (fyne package: GUI 子系统, 无控制台黑框, 内嵌图标)
function packageApp {
    Write-Host "📦 打包 Windows 桌面应用..." -ForegroundColor Cyan
    ensureGcc
    assertFyne
    if (-not (Test-Path ".\build\bin")) { New-Item -ItemType Directory -Force -Path ".\build\bin" | Out-Null }
    # 清理旧产物, 避免打包/发布携带过期 exe
    Get-ChildItem ".\build\bin\*.exe" -ErrorAction SilentlyContinue | Remove-Item -Force
    Push-Location $ENTRY
    try {
        fyne package -os windows -icon ..\..\build\appicon.png -release
        $code = $LASTEXITCODE
    } finally {
        Pop-Location
    }
    if ($code -ne 0) { Write-Host "❌ fyne package 失败" -ForegroundColor Red; exit 1 }
    Move-Item -Force "$ENTRY\*.exe" ".\build\bin\"
    Write-Host "✅ 打包完成, 产物位于 build\bin\$APP_NAME.exe"
}

# runApp 本地运行 (源码直跑, 便于开发调试)
function runApp {
    Write-Host "🚀 运行 $APP_NAME..." -ForegroundColor Cyan
    ensureGcc
    go run $ENTRY
}

# clean 清理构建产物
function clean {
    Write-Host "🧹 清理构建产物..." -ForegroundColor Cyan
    if (Test-Path ".\build\bin") { Remove-Item -Recurse -Force ".\build\bin" }
    go clean
    Write-Host "✅ 清理完成"
}

# ci CI 全流程: 静态检查 + 测试 + 打包
function ci {
    goLint
    testGo
    packageApp
    Write-Host "✅ CI 全流程执行完毕" -ForegroundColor Green
}

# switch 放到最后, 执行用户选择的操作
switch ($selected) {
    0 { installTools }
    1 { ci }
    2 { formatGoCode }
    3 { goLint }
    4 { testGo }
    5 { packageApp }
    6 { runApp }
    7 { clean }
    default { Write-Host "❌ 无效的选择" -ForegroundColor Red }
}
