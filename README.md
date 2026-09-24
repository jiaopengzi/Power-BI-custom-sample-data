# Power BI 自定义示例数据生成器

> Power BI Custom Sample Data — 使用**原生 Go + [Fyne](https://fyne.io/)** 构建的跨平台桌面应用, 一键生成符合真实零售业务场景的 Power BI 学习/演示示例数据, 产物为 **UTF-8 (含 BOM) 的 CSV** 与 **开箱即用的 PBIP 示例工程**, 任何工具皆可直接读取。

本项目由早期的 Access/VBA 版本 [《赠送 300 家门店 260 亿销售额的零售企业 Power BI 实战示例数据》](https://jiaopengzi.com/p/1035912423145473) 重构而来。原版数据存放在 Access 中不便分发与识别, 现改为跨平台桌面应用, 产物统一为 CSV 与 PBIP。

- 使用文档: <https://jiaopengzi.com/p/19051044919050241>
- 作者: 焦棚子 · <jiaopengzi@qq.com>

---

## 一、功能特性

- **一键生成** 11 张相互关联的业务/维度表 (见下表), 覆盖产品、门店、客户、入库、订单、销售目标、行政区划等。
- **PBIP 示例工程 (新功能)**: 生成时同步释放内嵌的 PBIP 模板 (语义模型 + 报表), 模型的 `Path` 参数自动指向 CSV 数据目录, 用 Power BI Desktop 直接打开即可, 无需手工建表与建立关系。
- **参数化时间窗口**: 通过 `开始日期` / `结束日期` 控制数据的时间范围 (替代原版写死的 `Now()-1500` 基线)。
- **增量更新**: 在已有数据基础上, 按新的日期区间追加订单/入库数据; 开始日期须为现有事实截止日的次日, 日期区间与现有数据冲突时会提示。
- **清空数据**: 一键删除已生成的 `data` 与 `pbip` 产物 (存放目录本身保留), 删除前有确认弹窗。
- **双语界面**: 简体中文 (默认) / English, 运行时可切换 (界面语言同时决定产物语言)。
- **CSV 友好**: 输出带 UTF-8 BOM, Excel 双击打开中文不乱码。
- **进度可视化与结果回显**: 生成过程实时显示进度与阶段, 完成后以表格展示各表行数与大小; 重新打开软件时会自动回显目录中已有的产物统计与事实数据截止日期。
- **单文件绿色运行**: 中文字体与图标已内嵌进 exe, 无需安装运行时或额外依赖。

### 数据表一览

| 文件名 | 说明 | 行数规模 |
| --- | --- | --- |
| `T00_产品表.csv` | 产品 (SKU) | 产品数量 (原 N0) |
| `T01_门店表.csv` | 门店 (含开/关店日期、经纬度) | 门店数量 (原 N1) |
| `T02_客户表.csv` | 客户 (按门店规模注册) | 与门店营业期相关 |
| `T03_入库信息表.csv` | 入库流水 | 与订单/入库周期相关 |
| `T04_订单主表.csv` | 订单主 (下单/送货/渠道/客户) | 数据量最大 |
| `T05_订单子表.csv` | 订单明细 (价格/折扣/数量/金额) | 数据量最大 |
| `T06_销售目标表.csv` | 按省份的月度销售目标 | 省份数 x 目标月份数 |
| `D00_大区表.csv` | 六大区 | 6 |
| `D01_省份表.csv` | 省份 | 34 |
| `D02_城市表.csv` | 城市 | ~370 |
| `D03_区县表.csv` | 区县 | ~2700 |

> 说明: 由于 VBA 的 `Rnd()` 与 Go 的随机序列无法逐字节复现, 本重构**保持字段、表结构一致**, 行数规则与业务分布与原版基本一致 (T00 产品名称与 T06 销售目标规则相对原版有增强, 见下文差异)。

### 产物目录结构

```text
<存放目录>
├── data/                    # 11 张 CSV 产物 (UTF-8 BOM)
└── pbip/                    # PBIP 示例工程
    ├── demo.pbip            # 工程入口, Power BI Desktop 直接打开
    ├── demo.Report/         # 报表 (页面/视觉对象/主题)
    └── demo.SemanticModel/  # 语义模型 (表结构/关系/度量值, Path 参数已指向 ../data)
```

---

## 二、使用手册

### 1. 下载与运行

从 [Releases](https://github.com/jiaopengzi/Power-BI-custom-sample-data/releases) 下载 Windows 压缩包, 解压后双击运行 `Power BI Custom Sample Data.exe`。程序为单文件绿色应用, 无需安装。
发版产物会同步镜像到 [Gitee Releases](https://gitee.com/jiaopengzi/power-bi-custom-sample-data/releases)。

### 2. 界面参数

| 参数 | 默认值 | 说明 | 取值范围 |
| --- | --- | --- | --- |
| 开始日期 | 今天 - 1600 天 | 数据时间窗口起点 | 早于结束日期, 窗口至少 60 天 |
| 结束日期 | 今天 | 数据时间窗口终点 | 晚于开始日期 |
| 产品数量 | 666 | 生成的 SKU 数量 (原 N0) | 1 – 10000 |
| 门店数量 | 55 | 生成的门店数量 (原 N1) | 1 – 10000 |
| 入库周期 | 22 | 入库间隔最大天数 (原 N3) | 5 – 180 |
| 存放目录 | 用户主目录下 `PowerBISampleData` | 产物输出目录 | 通过 `指定存放目录` 选择 |

> 日期可手动输入 `YYYY-MM-DD`, 也可点击右侧按钮用日历选择 (支持年份切换与今日/年初/月初/月末/年末快捷选择); 数量可直接输入或用+-微调。

### 3. 操作流程

1. 点击 **指定存放目录**, 选择产物输出目录 (默认在用户目录下的 `PowerBISampleData`)。
2. 设置日期区间与三个数量参数。
3. 点击 **生成示例数据**, 等待进度完成 (数据量越大耗时越长)。若目录已有数据会弹窗确认覆盖。
4. 产物落在 `<存放目录>/data` (11 张 CSV) 与 `<存放目录>/pbip` (PBIP 工程), 点击 **打开存放目录** 查看, 结果表格会列出各表行数与大小。
5. 用 Power BI Desktop (需支持 Power BI 项目的版本) 打开 `<存放目录>\pbip\demo.pbip` 即可继续分析; 也可通过 `获取数据 → 文本/CSV` 自行导入 `data` 目录下的 CSV, 按表名建立关系。

### 4. 增量更新

- 若目录中**没有基础数据**, 会提示先执行 `生成示例数据`。
- **开始日期必须等于 "事实数据截止日期 + 1 天"**: 界面会显示事实数据的截止日期 (增量按钮旁), 起始日期不符合要求时会提示应填的日期。
- 点击 **增量更新**, 系统会复用已有的产品/门店/客户维度, 追加该区间的订单与入库数据, 并续接各表的 `F_00_自动编号` 主键。
- 增量完成后会依据全部实绩**重建 `T06_销售目标表`**: 目标区间随事实区间扩张, 增量前已生成且实绩未变的月份目标值保持不变。
- PBIP 工程的日历表年份范围会**只扩张不缩小** (起年取较小值, 止年取较大值), 保证已有数据的日期不被裁掉; 其余 PBIP 内容不变。
- 若日期区间与现有订单区间**冲突**, 会提示现有数据的日期范围, 请调整后重试。

### 5. 清空数据

点击 **清空数据** (红色危险按钮, 仅在目录已有数据时可用), 确认后删除 `<存放目录>/data` 下的全部 CSV 与整个 `<存放目录>/pbip` 目录; 存放目录本身及其中的其他文件保留。

### 6. 语言切换

界面右上角下拉框可在 `简体中文` / `English` 间切换; 切换会同时改变界面语言与后续生成产物的语言。

---

## 三、开发手册

### 1. 环境要求

| 工具 | 版本 / 说明 |
| --- | --- |
| Go | 1.27.x (见 `go.mod`) |
| gcc (MinGW-w64) | **CGO 必需** — Fyne 依赖系统 GUI/OpenGL 绑定; Windows 推荐 `scoop install mingw` |
| fyne CLI | `fyne.io/tools` (打包用), `go install fyne.io/tools/cmd/fyne@latest` |
| golangci-lint | v2.x (静态检查, 可选) |

> 本项目使用 [Fyne](https://fyne.io/) 作为 GUI 框架, 因此**编译/运行/测试都需要 CGO** (`CGO_ENABLED=1`) 与本机 C 编译器 (gcc)。`run.ps1` 已自动开启 CGO, 并在 PATH 找不到 gcc 时自动探测 scoop/mingw 常见安装位置。

### 2. 目录结构

```text
.
├── cmd/
│   └── pbicsd/
│       ├── main.go          # 可执行入口: ui.Run()
│       └── FyneApp.toml     # fyne package 打包元数据 (Name/ID/Version/Icon)
├── internal/
│   ├── app/                 # 与 GUI 无关的控制器 (生成/增量/目录/数据探测/清空)
│   ├── ui/                  # Fyne 界面 (主题/主窗口/日期与数量控件/内嵌资源)
│   │   └── assets/          # 内嵌图标与中文字体 (go:embed, 含 OFL 许可)
│   ├── i18n/                # zh-cn / en-us 文案表与运行时语言切换
│   ├── config/              # 配置与参数校验
│   ├── data/                # 内嵌基础数据 (go:embed)
│   │   └── assets/          # province/city/district.csv, first/last_name.txt
│   │       └── pbip/        # PBIP 模板 (demo.pbip + demo.Report + demo.SemanticModel)
│   ├── model/               # 表名与表头定义
│   ├── util/                # 随机/银行家舍入/格式化辅助
│   ├── csvw/                # 带 UTF-8 BOM 的 CSV 写入器
│   └── generator/           # 各表生成逻辑 + 全量编排 + 增量更新
├── build/
│   ├── appicon.png          # 应用图标源 (1024×1024 PNG)
│   └── windows/icon.ico     # Windows 图标
├── .chglog/                 # git-chglog 配置与模板
├── .gitalias/savetag.sh     # git savetag 别名脚本 (提交 CHANGELOG 并打 tag)
├── .github/workflows/       # build.yaml (构建发版) + sync_gitee.yaml (Gitee 镜像)
├── CHANGELOG.md             # 版本变更记录 (Keep a Changelog)
├── run.ps1                  # 本地开发 / CI 脚本
└── go.mod
```

### 3. 常用命令

推荐使用 `run.ps1` (交互式菜单, 或 `-Choice <编号>` 非交互):

| 编号 | 作用 |
| --- | --- |
| 0 | 安装/更新开发工具 (fyne CLI + golangci-lint) |
| 1 | CI 全流程 (静态检查 + 测试 + 打包) |
| 2 | 格式化 Go 代码 (`gofmt` + `go mod tidy`) |
| 3 | Go 静态检查 (`go vet` + `golangci-lint`) |
| 4 | Go 单元测试 (`go test`) |
| 5 | 打包 Windows 桌面应用 (`fyne package`) |
| 6 | 本地运行 (`go run`) |
| 7 | 清理构建产物 |

```powershell
.\run.ps1 -Choice 6      # 运行
.\run.ps1 -Choice 1      # 静态检查 + 测试 + 打包
```

等价的原生命令 (需 `CGO_ENABLED=1` 且 gcc 在 PATH):

```bash
go run ./cmd/pbicsd            # 运行
go build ./...                 # 编译
go test ./...                  # 测试 (含全量/增量端到端测试)
golangci-lint run ./...        # 静态检查 (配置见 .golangci.yaml)

# 打包 (在 cmd/pbicsd 下执行, 产出 GUI 子系统 exe, 无控制台黑框)
cd cmd/pbicsd
fyne package -os windows -icon ../../build/appicon.png -release
```

### 4. PBIP 模板维护

- 模板源文件位于 `internal/data/assets/pbip/` (`demo.pbip` + `demo.Report` + `demo.SemanticModel`), 由 `internal/data/pbip.go` 在生成时释放到 `<存放目录>/pbip`。
- 释放时完成两处**动态替换**: `expressions.tmdl` 的 `Path` 参数改指 `<存放目录>/data`; `01_Calendar.tmdl` 的 `date_start`/`date_end` 年份改写为生成窗口起止年份。增量更新后仅**扩张**日历年份范围 (`ExpandPbipCalendar`)。
- `//go:embed` 清单在 `internal/data/pbip.go` 中**逐条列出**而非整树内嵌: `.gitignore` 忽略的 `**/.pbi/localSettings.json` 与 `**/.pbi/cache.abf` (本地缓存可达上百 MB) 不能进入产物,
  而点开头的隐藏文件又是模板必需。增删模板文件时必须同步调整 embed 清单, 对齐关系由 `TestPbipEmbedComplete` 守护。

### 5. 界面文案 (i18n)

- 所有文案集中在 `internal/i18n/messages.go` 的扁平点号键表 (`title/subtitle`、`form.*`、`buttons.*`、`stages.*`、`result.*`、`msg.*`、`common.*`)。
- `internal/i18n/i18n.go` 的 `Manager` 提供 `T(key)` / `Tf(key, vars)` 与运行时 `SetLocale`; `stages.*` 键与生成器发出的阶段标识对齐。
- 新增/修改文案时同时补齐中英文两列即可, 界面会在切换语言时整体刷新。
- **注意同步**: 参数取值范围同时存在于 `internal/config/config.go` 的常量与 `msg.productRange` / `msg.storeRange` / `msg.inventoryRange` 提示文案中, 修改任一侧必须同步另一侧。

### 6. 中文字体 (内嵌)

- 默认 Fyne 字体只含拉丁字形, 中文会渲染为方块 (tofu), 故内嵌了子集化的 **Noto Sans SC (思源黑体, OFL 许可)**。
- 位置: `internal/ui/assets/NotoSansSC.ttf`; 许可与来源见同目录 `OFL.txt` 与 `README.md`。
- 由可变字体固定到 `wght=400` 后, 子集化为 ASCII + 常用标点 + 完整 GB2312 汉字 (保证任意简体中文路径可显示) + 文案表实际用字, 约 2.4 MB。
- `go:embed` 无法跨目录引用, 故字体/图标资源都放在 `internal/ui/assets/` 内, 由 `internal/ui/resources.go` 内嵌。

### 7. 图标

| 用途 | 位置 | 说明 |
| --- | --- | --- |
| 界面内 Logo / exe 图标 | `build/appicon.png` | 1024×1024 PNG, 打包时经 `-icon` 传给 `fyne package` 生成 exe 图标; 界面头部也复用它 (内嵌副本在 `internal/ui/assets/appicon.png`) |
| Windows 图标 | `build/windows/icon.ico` | 备用 `.ico` |

> 替换图标: 覆盖 `build/appicon.png` 并同步 `internal/ui/assets/appicon.png` (二者保持一致), 重新打包即可。`build/` 目录需入库, 仅 `build/bin/` (编译产物) 被 git 忽略。

### 8. 主题色

主色使用深藏蓝 `#1e2858` + 金 `#c89828`, 在 `internal/ui/theme.go` 的 `brandTheme` 中实现: 强制浅色变体, 覆盖 `ColorNamePrimary` (按钮/进度条) 与 `ColorNameHyperlink` (文档链接), 并返回内嵌中文字体。修改该文件即可换肤。

### 9. 发版流程

1. 在 `CHANGELOG.md` 顶部新增版本章节 (`## [vX.Y.Z] - YYYY-MM-DD`, 遵循 Keep a Changelog + SemVer, 版本号以小写 `v` 开头)。
2. 执行 `git savetag` (别名指向 `.gitalias/savetag.sh`, 一次性配置 `git config --global alias.savetag '!bash ./.gitalias/savetag.sh'`): 脚本校验版本号合法且 tag 未被占用后, 以 `Release:  vX.Y.Z` 提交 CHANGELOG 并推送, 再打同名 tag 推送。
3. 推送 `v*` tag 触发 `.github/workflows/build.yaml`: 按 `go.mod` 版本装配 Go → `run.ps1 -Choice 1` (静态检查 + 测试 + 打包) → 产物压缩为 `<tag>-windows.zip` → 从 CHANGELOG 提取该 tag 的章节生成发版说明 → 创建 GitHub Release。
4. 随后 `.github/workflows/sync_gitee.yaml` 将 main 分支与标签**强制镜像**到 [Gitee](https://gitee.com/jiaopengzi/power-bi-custom-sample-data),
   并复用 GitHub Release 附件覆盖式发布到 Gitee Release (依赖仓库 secrets `GITEE_SSH_PRIVATE_KEY` / `GITEE_TOKEN`)。

### 10. 与原 VBA 版本的差异

- 产物由 Access 表改为 **CSV (UTF-8 BOM) + PBIP 示例工程**, CSV 落在 `<存放目录>/data`, PBIP 落在 `<存放目录>/pbip`。
- 时间基线由 `Now()-1500` **参数化**为 `开始日期 / 结束日期` (默认起始为今天 - 1600 天)。
- 新增 **增量更新** 功能, 起始日期限定为事实截止日 + 1 天; 增量完成后按全部实绩重建 T06 销售目标表 (未变化月份的目标值保持不变), 并扩张 PBIP 日历表年份。
- 新增 **清空数据** 功能与启动时的历史产物回显。
- 参数上限放宽: 产品/门店数量 1 – 10000, 入库周期 5 – 180。
- 随机序列不同 (逐字节不可复现), 但业务逻辑/字段/行数规则一致。
- **产品名称**缩短为 产品 + 3 位字母数字组合 (如 `产品A7X`, 原为 `产品B0122`), 组合全局唯一, 产品编号不变 (如 `SKU_000122`)。
- **产品售价**改为对数正态分布并截断在 [1000, 30000] (家具品类, 中位数约 6000, 多数偏低少数高价; 原为与分类强绑定的均匀分布 [5000, 10000]), 产品分类仍按价格分位取 A-J。
- **销售目标**覆盖区间为 首个销售事实年度次年的 1 月 至 最后事实年度的年末 (如 事实 2022-08 起 2026-08 止 → 目标 2023-01 至 2026-12, 原仅最近两年);
  事实期内的月份以各省当月实绩为锚, 完成率约落在 90%-120% (原方案整体完成率仅约 80%); 事实期之后的计划月份按近 12 个月月均叠加行业淡旺季外推;
  目标带**年度偏置** (部分年度偏乐观/偏保守, 年度完成率围绕 100% 分化)。
- 新增**随机因子** (时间/产品维度): 年景系数、月度噪声、星期系数、门店客流系数 (钟形近似正态)、产品人气系数 (钟形近似正态), 打破不同年份/门店/产品之间 "几乎一致" 的形态。
- GUI 由 Access 窗体 → (曾用 Wails + Vue3) → 现为**原生 Go + Fyne**, 业务逻辑层保持不变。

---

## License

- 代码: [MIT](LICENSE) © 焦棚子
- 内嵌字体 Noto Sans SC: [SIL Open Font License 1.1](internal/ui/assets/OFL.txt)
