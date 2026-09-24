# AGENTS.md

面向 AI 编程助手的项目约定。修改任何代码前请先读完本文件。

## 项目概览

- 原生 Go + Fyne 的 Windows 桌面应用: 一键生成 Power BI 零售示例数据 (11 张相互关联的 CSV) 并释放可直接用 Power BI Desktop 打开的 PBIP 示例工程。
- 由 Access/VBA 版本重构而来, **表结构、字段名、行数规则与业务分布必须与原版保持一致** (随机序列除外, 逐字节不可复现是预期行为)。
- 界面与产物均为双语: `zh-cn` (默认) / `en-us`。
- 详细功能与使用说明见 [README.md](./README.md)。

## 环境与命令 (硬性要求)

- Windows 开发机; Go 版本以 `go.mod` 为准 (当前 1.27.1), 不手动指定其他版本。
- Fyne 依赖 CGO: **编译/运行/测试一律需要 `CGO_ENABLED=1` + gcc (MinGW-w64)**。`run.ps1` 会自动开启 CGO 并探测 scoop/mingw 常见安装路径; 直接跑 `go test` 报 CGO 链接错误时, 先经 `run.ps1` 或自行确保 gcc 在 PATH。
- 统一入口 `run.ps1` (交互菜单或 `.\run.ps1 -Choice <编号>`):

| 编号 | 作用 |
| --- | --- |
| 0 | 安装/更新开发工具 (fyne CLI + golangci-lint) |
| 1 | CI 全流程 (静态检查 + 测试 + 打包) |
| 2 | 格式化 (`gofmt` + `go mod tidy`) |
| 3 | 静态检查 (`go vet` + `golangci-lint`) |
| 4 | 单元测试 (`go test ./...`) |
| 5 | 打包 Windows 应用 (`fyne package`) |
| 6 | 本地运行 (`go run`) |
| 7 | 清理构建产物 |

- **提交前最低要求**: `Choice 3` 与 `Choice 4` 全部通过; 改动 UI 时建议再跑 `Choice 6` 人工过一遍界面。

## 架构与数据流

```text
cmd/pbicsd (入口 main.go)
  └── internal/ui        Fyne 界面: 只做渲染与事件, 不含业务逻辑
        └── internal/app       GUI 无关的编排层: 参数校验/生成/增量/目录探测/清空 (不 import Fyne, 可独立单测)
              └── internal/generator   各表生成逻辑 + 全量编排 + 增量更新
                    ├── internal/csvw      带 UTF-8 BOM 的 CSV 写入器 (产物一律经它写出)
                    ├── internal/data      内嵌地理/姓名基础数据 + PBIP 模板释放
                    ├── internal/model     表名与表头唯一定义处
                    └── internal/util      随机/银行家舍入等 VBA 移植辅助
internal/config  参数范围与校验; internal/i18n  双语文案表
```

- 响应码 `CodeOK` / `CodeError` / `CodeNoBaseData` / `CodeDateConflict` (`internal/app/types.go`) 与 `internal/ui/mainwindow.go` 的 `renderResponse` 分支一一对应, 新增码需两侧同步。
- 产物目录约定: CSV → `<存放目录>/data`, PBIP → `<存放目录>/pbip`; 由 `internal/app/controller.go` 的 `dataDir` / `pbipDir` 统一拼接, 不要在别处硬编码。

## 代码风格 (强约束)

- 每个源文件头部固定注释块: `FilePath / Author / Blog / Copyright / Description`。
- 注释一律**中文**; 包/类型/函数 doc 注释遵循现有格式, 函数注释需列明参数与返回值含义。
- import 按 **标准库 / 三方 / 本地 (`jiaopengzi/Power-BI-custom-sample-data`)** 三段分组 (goimports 已配置 local-prefixes)。
- 静态检查以 `.golangci.yaml` (v2) 为准: 圈复杂度与认知复杂度阈值 15, 另启用 gosec / dupl(150) / goconst / errcheck(含 `_` 忽略) 等; 新增代码不得引入告警; 故意忽略错误时用项目惯用的 `#nosec` / `//nolint:errcheck` 并写明理由。
- lint 不检查 `_test.go` (`run.tests: false`), 但测试代码必须可编译且通过 `go test`。

## 修改时的同步点 (易错)

1. **i18n 双语**: 所有界面文案集中在 `internal/i18n/messages.go` 扁平点号键表, zh/en 两列必须同时补齐; 界面层不得硬编码用户可见文本。
2. **参数范围双份维护**: 取值范围同时存在于 `internal/config/config.go` 常量与 i18n 的 `msg.productRange` / `msg.storeRange` / `msg.inventoryRange` 提示文案, 改一侧必须同步另一侧。
3. **阶段键对齐**: `stages.*` 文案键与 generator 进度回调发出的阶段标识 (`stageStart` / `stageDimensions` / ...) 逐字对应。
4. **表定义单点**: 表名与表头只在 `internal/model/tables.go` 定义; `AllFiles` 顺序被增量前置检查、结果统计与清空逻辑复用。
5. **PBIP 模板 embed 清单**: 模板在 `internal/data/assets/pbip/`, `internal/data/pbip.go` 的 `//go:embed` 是**逐条列出**的 (防止 `.pbi` 本地缓存入库, 又保留点开头的必需文件);
   增删模板文件必须同步该清单, `TestPbipEmbedComplete` 会校验磁盘与内嵌一致。`**/.pbi/localSettings.json` 与 `**/.pbi/cache.abf` 永不入库、不入 embed。
6. **PBIP 动态替换**: 释放时 `expressions.tmdl` 的 `Path` 改指 data 目录, `01_Calendar.tmdl` 年份改写为生成窗口年份; 增量后日历年份**只扩张不缩小**。调整模板时注意保持这两处正则可匹配。
7. **BOM**: CSV 产物必须经 `internal/csvw` 写出 (自动带 UTF-8 BOM), 不要用 `encoding/csv` 直接写产物文件。
8. **内嵌资源位置**: `go:embed` 不能跨目录, 字体/图标等资源必须放在引用包自己的 `assets/` 下。

## 测试

- `go test ./...` (需 CGO)。`internal/generator` 含全量/增量端到端测试; `internal/data` 含 PBIP 内嵌完整性测试; `internal/ui` 含滚轮透传回归测试; `internal/app` 有控制器测试。
- 改动 generator 业务规则时, 优先在对应 `_test.go` 补断言, 而不是只靠手工运行验证。

## 发版 (不要随意执行)

- `CHANGELOG.md` 手动维护 (Keep a Changelog + SemVer, 版本以小写 `v` 开头)。
- `git savetag` (别名指向 `.gitalias/savetag.sh`): 校验版本 → 提交 CHANGELOG → 打 tag 并推送; 推 `v*` tag 触发 `build.yaml` 自动构建发版, 随后 `sync_gitee.yaml` 强制镜像到 Gitee。
- **除非用户明确要求发版, 不要打 tag、不要运行 `git savetag`、不要修改 CHANGELOG 的既有版本章节。**

## 其他

- `build/` 目录入库 (图标等), 仅 `build/bin/` 忽略。
- `.feat/` / `.debug/` / `.mimosa/` 等目录是本地过程记录, 已被 `.gitignore` 忽略, 不要提交也不要引用其中的路径。
- Markdown 遵循 `.markdownlint.yaml`: 行宽 ≤ 200, 列表缩进 4 空格。
- 提交信息沿用仓库既有风格 (`✨ feat:` / `🐞 fix:` / `📝 docs:` / `♻️ refactor:` / `⚙️ ci:` 等 emoji + type 前缀)。
