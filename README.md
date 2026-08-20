# Power BI 自定义示例数据生成器

> Power BI Custom Sample Data Generator — 使用 Go + Wails + Vue3 重构的桌面应用, 一键生成符合真实零售业务场景的 Power BI 学习/演示示例数据, 产物为 **UTF-8 (含 BOM) 的 CSV**, 任何工具皆可直接读取。

本项目由早期的 Access/VBA 版本 [《赠送 300 家门店 260 亿销售额的零售企业 Power BI 实战示例数据》](https://jiaopengzi.com/?post_id=1035912423145473) 重构而来。原版数据存放在 Access 中不便分发与识别, 现改为跨平台桌面应用, 产物统一为 CSV。

- 使用文档: <https://jiaopengzi.com/?post_id=19051044919050241>
- 作者: 焦棚子 · <jiaopengzi@qq.com>

---

## 一、功能特性

- **一键生成** 11 张相互关联的业务/维度表 (见下表), 覆盖产品、门店、客户、入库、订单、销售目标、行政区划等。
- **参数化时间窗口**: 通过 `开始日期` / `结束日期` 控制数据的时间范围 (替代原版写死的 `Now()-1500` 基线)。
- **增量更新 (新功能)**: 在已有数据基础上, 按新的日期区间追加订单/入库数据; 日期区间与现有数据冲突时会提示。
- **双语界面**: 简体中文 (默认) / English, 运行时可切换。
- **CSV 友好**: 输出带 UTF-8 BOM, Excel 双击打开中文不乱码。
- **进度可视化**: 生成过程实时显示进度与阶段。

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

从 [Releases](https://github.com/jiaopengzi/Power-BI-custom-sample-data/releases) 下载对应平台的压缩包, 解压后运行 `PowerBISampleDataGenerator.exe` (Windows)。

### 2. 界面参数

| 参数 | 说明 | 取值范围 |
| --- | --- | --- |
| 开始日期 | 数据时间窗口起点 | 早于结束日期, 窗口至少 60 天 |
| 结束日期 | 数据时间窗口终点 | 晚于开始日期 |
| 产品数量 | 生成的 SKU 数量 (原 N0) | 1 – 2000 |
| 门店数量 | 生成的门店数量 (原 N1) | 1 – 400 |
| 入库周期 | 入库间隔最大天数 (原 N3) | 5 – 20 |
| 存放目录 | CSV 产物输出目录 | 通过 `指定存放目录` 选择 |

### 3. 操作流程

1. 点击 **指定存放目录**, 选择 CSV 输出目录。
2. 设置日期区间与三个数量参数。
3. 点击 **生成示例数据**, 等待进度完成 (数据量越大耗时越长)。
4. 点击 **打开存放目录** 查看产物。
5. 将 CSV 导入 Power BI (`获取数据 → 文本/CSV` 或 `文件夹`), 按表名建立关系即可。

### 4. 增量更新

- 若目录中**没有基础数据**, 会提示先执行 `生成示例数据`。
- 设置一个**与现有数据不重叠**的日期区间, 点击 **增量更新**, 系统会复用已有的产品/门店/客户维度, 追加该区间的订单与入库数据。
- 若日期区间与现有订单区间**冲突**, 会提示现有数据的日期范围, 请调整后重试。

### 5. 语言切换

界面右上角下拉框可在 `简体中文` / `English` 间切换。

---

## 三、开发手册

### 1. 环境要求

| 工具 | 版本 |
| --- | --- |
| Go | 1.26.x (见 `go.mod`) |
| Node.js | 24.x |
| pnpm | 11.x |
| Wails CLI | v2.10.x |

安装 Wails CLI:

```bash
go install github.com/wailsapp/wails/v2/cmd/wails@latest
wails doctor
```

### 2. 目录结构

```text
.
├── main.go                 # Wails 应用入口
├── app.go                  # 绑定给前端的 API (生成/增量/目录/语言)
├── wails.json              # Wails 配置 (使用 pnpm)
├── go.mod
├── internal/
│   ├── config/             # 配置与参数校验
│   ├── data/               # 内嵌地理/姓名基础数据 (go:embed)
│   │   └── assets/         # province/city/district.csv, first/last_name.txt
│   ├── model/              # 表名与表头定义
│   ├── util/               # 随机/银行家舍入/格式化辅助
│   ├── csvw/               # 带 UTF-8 BOM 的 CSV 写入器
│   └── generator/          # 各表生成逻辑 + 全量编排 + 增量更新
├── frontend/               # Vue3 + TypeScript + Element Plus
│   ├── src/
│   │   ├── api/            # Wails 绑定封装与类型
│   │   ├── i18n/           # zh-cn / en-us 文案
│   │   ├── App.vue         # 主界面
│   │   └── main.ts
│   ├── oxlint.config.ts    # oxlint 静态检查
│   ├── .oxfmtrc.json       # oxfmt 格式化
│   ├── vitest.config.ts    # vitest 测试
│   └── tsconfig.*.json
└── .github/workflows/build.yaml
```

### 3. 后端 (Go)

```bash
go build ./...                 # 编译
go test ./...                  # 测试 (含全量/增量端到端测试)
golangci-lint run ./...        # 静态检查 (配置见 .golangci.yaml)
```

### 4. 前端 (frontend/)

```bash
pnpm install                   # 安装依赖
pnpm dev                       # Vite 开发服务器 (通常经 wails dev 启动)
pnpm build                     # 类型检查 + 构建到 dist
pnpm type-check                # vue-tsc 类型检查
pnpm lint / pnpm lint:fix      # oxlint
pnpm fmt                       # oxfmt 格式化
pnpm test                      # vitest
```

### 5. 桌面应用开发/打包

```bash
wails dev                      # 热重载开发 (自动运行前端 pnpm dev)
wails build -platform windows/amd64   # 打包, 产物在 build/bin/
```

### 6. 基础数据来源

`internal/data/assets/` 下的省市区与姓名数据由从旧版 `demo_jiaopengzi_data_vba.bas` 抽取而来如需更新原始数据.

### 7. Logo 与图标

| 用途 | 位置 | 说明 |
| --- | --- | --- |
| 界面内 Logo | `frontend/src/assets/logo.svg` | 显示在应用头部, 直接替换该 SVG 即可 (当前为占位) |
| 应用/窗口/exe 图标 | `build/appicon.png` | 1024×1024 PNG, Wails 由此自动生成各平台图标 (Windows `.ico` 等) |
| 网页 favicon | `frontend/public/` + `frontend/index.html` 的 `<link rel="icon">` | 可选 |

> 首次运行 `wails dev` 或 `wails build` 会生成 `build/` 目录及默认 `build/appicon.png`, 用你的图标替换后重新 `wails build` 即可。`build/` 目录 (图标与平台配置) **需要入库**, 仅 `build/bin/` (编译产物) 被 git 忽略。

### 8. 主题色

主色使用深藏蓝 `#1e2858` + 金 `#c89828`, 在 `frontend/src/style.scss` 中通过覆盖 Element Plus 的 `--el-color-primary` 系列变量统一按钮/进度条/日期选择器等控件配色, 修改该文件即可换肤。

### 9. 与原 VBA 版本的差异

- 产物由 Access 表改为 **CSV (UTF-8 BOM)**。
- 时间基线由 `Now()-1500` **参数化**为 `开始日期 / 结束日期`。
- 新增 **增量更新** 功能。
- 随机序列不同 (逐字节不可复现), 但业务逻辑/字段/行数规则一致。

---

## License

[MIT](LICENSE) © 焦棚子
