# Changelog

本文件将记录本项目的所有重要变更。

该格式基于 [Keep a Changelog](https://keepachangelog.com),
本项目遵循 [语义化版本控制](https://semver.org/spec/v2.0.0.html)。

## [v0.2.0] - 2026-09-25

### Feat

- 调整随机因子

## [v0.2.0-rc] - 2026-09-24

### Test

- 测试 ci/cd 流程(01)

## [v0.1.0-rc1] - 2026-08-21

### Added

- ✨ 新增原生 Fyne 日期输入控件, 支持日历弹窗、年份切换, 以及今日/年初/月初/月末/年末快捷选择。
- ✨ 新增整数输入控件与错误边框反馈, 支持数值步进、区间校验与非法输入即时提示。
- ✨ 新增生成结果表格、事实数据截止日期提示与清空已生成数据入口, 完善桌面端主窗口交互闭环。
- 🧪 新增输入框滚轮透传相关回归测试, 覆盖日期、数字与目录输入框的真实主窗口场景。

### Changed

- ♻️ 重构: GUI 由 Wails + Vue3 迁移到**原生 Go + Fyne**, 业务逻辑层 (`internal/{config,data,generator,model,csvw,util}`) 完全保持不变。
- 🏗️ 采用标准 Go 布局: 入口移至 `cmd/pbicsd`, 控制器 `internal/app`, 界面 `internal/ui`, 文案 `internal/i18n`。
- 🎨 重构主窗口布局与主题细节, 调整间距、进度条强调色、结果区排版与控件状态反馈, 更贴近桌面应用体验。
- 📁 目录选择改为系统原生选择器, Windows 下通过原生对话框选择输出目录。
- 📊 结果展示由简单统计区升级为表格化展示, 同时展示表名、行数与文件大小。
- 🔤 内嵌子集化的 Noto Sans SC (OFL) 中文字体与粗体字重资源, 界面中文不再显示为方块 (tofu), 加粗文本也能正确渲染。
- 🏷️ 窗口标题更名为 `Power BI Custom Sample Data`。
- 🧰 重写 `run.ps1` 与 `.github/workflows/build.yaml`: 去除 pnpm/Node/Wails, 改用 `fyne package` (CGO) 打包。

### Fixed

- 🛞 修复日期、数字和目录输入框在鼠标悬停时吞掉滚轮事件, 导致页面无法继续滚动的问题。
- ✅ 修复表单联动校验体验, 开始/结束日期先后关系、增量起始日期与目录非空状态现在会即时反馈并同步影响按钮可用性。
- 🌐 修复部分增量更新与范围冲突提示的本地化表现, 错误文案改为由界面层统一提供中英文包装。

### Removed

- 🗑️ 移除前端工程 (`frontend/`, Vue3 + Element Plus) 及 Wails 相关配置与依赖。

## [v0.1.0-rc] - 2026-08-20

### Initial release

- 生产环境首次发布.
