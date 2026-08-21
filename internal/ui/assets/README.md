# 内嵌资源说明 (Embedded assets)

本目录下的资源通过 `//go:embed`（见 `../resources.go`）编入可执行文件。
`go:embed` 无法跨目录引用（不能 `..`），故资源必须放在本包内。

## appicon.png

应用图标，复制自仓库根的 `build/appicon.png`（同一图标，二者保持一致）。

## NotoSansSC.ttf

界面中文字体，用于避免默认字体将中文渲染为方块（tofu）。

- 字体：Noto Sans SC（思源黑体 简体中文，基于 Adobe Source Han Sans）。
- 许可：SIL Open Font License 1.1，见同目录 `OFL.txt`。
- 来源：Google Fonts 可变字体 `ofl/notosanssc/NotoSansSC[wght].ttf`。
- 处理：先将可变字体固定到 `wght=400`（Regular），再子集化以缩小体积。
  - 覆盖范围：ASCII、常用标点、CJK 符号/标点、完整 GB2312 汉字（保证任意简体中文路径可显示），以及文案表中实际出现的全部 CJK 字符。
  - 结果：约 2.4 MB，含约 7700 个字形。

复现子集化（需 Python + fonttools）：

```python
# 1. 下载可变字体 NotoSansSC[wght].ttf
# 2. 固定字重: instantiateVariableFont(font, {"wght": 400})
# 3. 子集化: Subsetter().populate(unicodes=<ASCII+标点+GB2312+文案CJK>)
# 4. 保存为 NotoSansSC.ttf
```

> 注：OFL 保留字体名（RFN）为 “Source”/“Noto”，如需再分发请遵守 OFL 1.1，勿以保留名命名衍生字体。
