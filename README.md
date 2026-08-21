# Power BI 自定义示例数据生成器

> Power BI Custom Sample Data — 使用**原生 Go + [Fyne](https://fyne.io/)** 构建的跨平台桌面应用, 一键生成符合真实零售业务场景的 Power BI 学习/演示示例数据, 产物为 **UTF-8 (含 BOM) 的 CSV**, 任何工具皆可直接读取。

本项目由早期的 Access/VBA 版本 [《赠送 300 家门店 260 亿销售额的零售企业 Power BI 实战示例数据》](https://jiaopengzi.com/?post_id=1035912423145473) 重构而来。原版数据存放在 Access 中不便分发与识别, 现改为跨平台桌面应用, 产物统一为 CSV。

- 使用文档: <https://jiaopengzi.com/?post_id=19051044919050241>
- 作者: 焦棚子 · <jiaopengzi@qq.com>

---

## 一、功能特性

- **一键生成** 11 张相互关联的业务/维度表 (见下表), 覆盖产品、门店、客户、入库、订单、销售目标、行政区划等。
- **参数化时间窗口**: 通过 `开始日期` / `结束日期` 控制数据的时间范围 (替代原版写死的 `Now()-1500` 基线)。
- **增量更新 (新功能)**: 在已有数据基础上, 按新的日期区间追加订单/入库数据; 日期区间与现有数据冲突时会提示。
- **双语界面**: 简体中文 (默认) / English, 运行时可切换 (界面语言同时决定产物语言)。
- **CSV 友好**: 输出带 UTF-8 BOM, Excel 双击打开中文不乱码。
- **进度可视化**: 生成过程实时显示进度与阶段。
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
| `T06_销售目标表.csv` | 按省份的月度销售目标 | 省份数 × 24 |
| `D00_大区表.csv` | 六大区 | 6 |
| `D01_省份表.csv` | 省份 | 34 |
| `D02_城市表.csv` | 城市 | ~370 |
| `D03_区县表.csv` | 区县 | ~2700 |

> 说明: 由于 VBA 的 `Rnd()` 与 Go 的随机序列无法逐字节复现, 本重构**保持字段、表结构、行数规则与业务分布一致**, 但具体随机值与原版不同。

---

## 二、使用手册

### 1. 下载与运行

从 [Releases](https://github.com/jiaopengzi/Power-BI-custom-sample-data/releases) 下载 Windows 压缩包, 解压后双击运行 `Power BI Custom Sample Data.exe`。程序为单文件绿色应用, 无需安装。

### 2. 界面参数

| 参数 | 说明 | 取值范围 |
| --- | --- | --- |
| 开始日期 | 数据时间窗口起点 | 早于结束日期, 窗口至少 60 天 |
| 结束日期 | 数据时间窗口终点 | 晚于开始日期 |
| 产品数量 | 生成的 SKU 数量 (原 N0) | 1 – 2000 |
| 门店数量 | 生成的门店数量 (原 N1) | 1 – 400 |
| 入库周期 | 入库间隔最大天数 (原 N3) | 5 – 20 |
| 存放目录 | CSV 产物输出目录 | 通过 `指定存放目录` 选择 |

> 日期可手动输入 `YYYY-MM-DD`, 也可点击右侧按钮用日历选择; 数量可直接输入或用+-微调。

### 3. 操作流程

1. 点击 **指定存放目录**, 选择 CSV 输出目录 (默认在用户目录下的 `PowerBISampleData`)。
2. 设置日期区间与三个数量参数。
3. 点击 **生成示例数据**, 等待进度完成 (数据量越大耗时越长)。若目录已有数据会弹窗确认覆盖。
4. 点击 **打开存放目录** 查看产物。
5. 将 CSV 导入 Power BI (`获取数据 → 文本/CSV` 或 `文件夹`), 按表名建立关系即可。

### 4. 增量更新

- 若目录中**没有基础数据**, 会提示先执行 `生成示例数据`。
- 设置一个**与现有数据不重叠**的日期区间, 点击 **增量更新**, 系统会复用已有的产品/门店/客户维度, 追加该区间的订单与入库数据。
- 若日期区间与现有订单区间**冲突**, 会提示现有数据的日期范围, 请调整后重试。

### 5. 语言切换

界面右上角下拉框可在 `简体中文` / `English` 间切换; 切换会同时改变界面语言与后续生成产物的语言。

---

## 三、开发手册

### 1. 环境要求

| 工具 | 版本 / 说明 |
| --- | --- |
| Go | 1.26.x (见 `go.mod`) |
| gcc (MinGW-w64) | **CGO 必需** — Fyne 依赖系统 GUI/OpenGL 绑定; Windows 推荐 `scoop install mingw` |
| fyne CLI | `fyne.io/tools` (打包用), `go install fyne.io/tools/cmd/fyne@latest` |
| golangci-lint | v2.x (静态检查, 可选) |

> 本项目使用 [Fyne](https://fyne.io/) 作为 GUI 框架, 因此**编译/运行/测试都需要 CGO** (`CGO_ENABLED=1`) 与本机 C 编译器 (gcc)。`run.ps1` 已自动开启 CGO 并在缺少 gcc 时给出提示。

### 2. 目录结构

```text
.
├── cmd/
│   └── pbicsd/
│       ├── main.go          # 可执行入口: ui.Run()
│       └── FyneApp.toml     # fyne package 打包元数据 (Name/ID/Version/Icon)
├── internal/
│   ├── app/                 # 与 GUI 无关的控制器 (生成/增量/目录/数据探测)
│   ├── ui/                  # Fyne 界面 (主题/主窗口/日期与数量控件/内嵌资源)
│   │   └── assets/          # 内嵌图标与中文字体 (go:embed, 含 OFL 许可)
│   ├── i18n/                # zh-cn / en-us 文案表与运行时语言切换
│   ├── config/              # 配置与参数校验
│   ├── data/                # 内嵌地理/姓名基础数据 (go:embed)
│   │   └── assets/          # province/city/district.csv, first/last_name.txt
│   ├── model/               # 表名与表头定义
│   ├── util/                # 随机/银行家舍入/格式化辅助
│   ├── csvw/                # 带 UTF-8 BOM 的 CSV 写入器
│   └── generator/           # 各表生成逻辑 + 全量编排 + 增量更新
├── build/
│   ├── appicon.png          # 应用图标源 (1024×1024 PNG)
│   └── windows/icon.ico     # Windows 图标
├── run.ps1                  # 本地开发 / CI 脚本
├── go.mod
└── .github/workflows/build.yaml
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

### 4. 界面文案 (i18n)

- 所有文案集中在 `internal/i18n/messages.go` 的扁平点号键表 (`title/subtitle`、`form.*`、`buttons.*`、`stages.*`、`result.*`、`msg.*`、`common.*`)。
- `internal/i18n/i18n.go` 的 `Manager` 提供 `T(key)` / `Tf(key, vars)` 与运行时 `SetLocale`; `stages.*` 键与生成器发出的阶段标识对齐。
- 新增/修改文案时同时补齐中英文两列即可, 界面会在切换语言时整体刷新。

### 5. 中文字体 (内嵌)

- 默认 Fyne 字体只含拉丁字形, 中文会渲染为方块 (tofu), 故内嵌了子集化的 **Noto Sans SC (思源黑体, OFL 许可)**。
- 位置: `internal/ui/assets/NotoSansSC.ttf`; 许可与来源见同目录 `OFL.txt` 与 `README.md`。
- 由可变字体固定到 `wght=400` 后, 子集化为 ASCII + 常用标点 + 完整 GB2312 汉字 (保证任意简体中文路径可显示) + 文案表实际用字, 约 2.4 MB。
- `go:embed` 无法跨目录引用, 故字体/图标资源都放在 `internal/ui/assets/` 内, 由 `internal/ui/resources.go` 内嵌。

### 6. 图标

| 用途 | 位置 | 说明 |
| --- | --- | --- |
| 界面内 Logo / exe 图标 | `build/appicon.png` | 1024×1024 PNG, 打包时经 `-icon` 传给 `fyne package` 生成 exe 图标; 界面头部也复用它 (内嵌副本在 `internal/ui/assets/appicon.png`) |
| Windows 图标 | `build/windows/icon.ico` | 备用 `.ico` |

> 替换图标: 覆盖 `build/appicon.png` 并同步 `internal/ui/assets/appicon.png` (二者保持一致), 重新打包即可。`build/` 目录需入库, 仅 `build/bin/` (编译产物) 被 git 忽略。

### 7. 主题色

主色使用深藏蓝 `#1e2858` + 金 `#c89828`, 在 `internal/ui/theme.go` 的 `brandTheme` 中实现: 强制浅色变体, 覆盖 `ColorNamePrimary` (按钮/进度条) 与 `ColorNameHyperlink` (文档链接), 并返回内嵌中文字体。修改该文件即可换肤。

### 8. 与原 VBA 版本的差异

- 产物由 Access 表改为 **CSV (UTF-8 BOM)**。
- 时间基线由 `Now()-1500` **参数化**为 `开始日期 / 结束日期`。
- 新增 **增量更新** 功能。
- 随机序列不同 (逐字节不可复现), 但业务逻辑/字段/行数规则一致。
- GUI 由 Access 窗体 → (曾用 Wails + Vue3) → 现为**原生 Go + Fyne**, 业务逻辑层保持不变。

---

## License

- 代码: [MIT](LICENSE) © 焦棚子
- 内嵌字体 Noto Sans SC: [SIL Open Font License 1.1](internal/ui/assets/OFL.txt)
