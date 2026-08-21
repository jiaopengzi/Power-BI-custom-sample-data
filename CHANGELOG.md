# Changelog

本文件将记录本项目的所有重要变更。

该格式基于 [Keep a Changelog](https://keepachangelog.com),
本项目遵循 [语义化版本控制](https://semver.org/spec/v2.0.0.html)。

## [Unreleased]

### Changed

- ♻️ 重构: GUI 由 Wails + Vue3 迁移到**原生 Go + Fyne**, 业务逻辑层 (`internal/{config,data,generator,model,csvw,util}`) 完全保持不变。
- 🏗️ 采用标准 Go 布局: 入口移至 `cmd/pbicsd`, 控制器 `internal/app`, 界面 `internal/ui`, 文案 `internal/i18n`。
- 🔤 内嵌子集化的 Noto Sans SC (OFL) 中文字体, 界面中文不再显示为方块 (tofu)。
- 🏷️ 窗口标题更名为 `Power BI Custom Sample Data`。
- 🧰 重写 `run.ps1` 与 `.github/workflows/build.yaml`: 去除 pnpm/Node/Wails, 改用 `fyne package` (CGO) 打包。

### Removed

- 🗑️ 移除前端工程 (`frontend/`, Vue3 + Element Plus) 及 Wails 相关配置与依赖。

## [v0.1.0-rc] - 2026-08-20

### Initial release

- 生产环境首次发布.
