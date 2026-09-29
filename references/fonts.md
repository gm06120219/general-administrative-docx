# 字体策略与授权

## 默认字体映射

| 用途 | 首选字体 | 可分发兜底 |
|---|---|---|
| 公文标题 | 方正小标宋简体 | Noto Serif SC SemiBold |
| 正文、三级、四级 | 仿宋_GB2312 / 仿宋 | Noto Serif SC |
| 一级标题 | 黑体 | Noto Sans SC |
| 二级标题 | 楷体_GB2312 / 楷体 | Noto Serif SC |

生成器默认把首选字体名称写入 DOCX，以便安装了这些字体的电脑准确显示。可通过输入 JSON 的 `fonts` 字段切换为 Noto 字体，以获得更好的跨机器一致性。

## 本机检查结果（2026-09-28）

- 已发现：仿宋_GB2312、楷体_GB2312、仿宋、楷体、黑体。
- 未发现已注册的方正小标宋简体；发现的 `FZSTK.TTF` 和 `FZYTK.TTF` 分别是方正舒体、方正姚体，不能替代小标宋。
- 已发现：Noto Sans SC、Noto Serif SC。

## 公开仓库规则

- `assets/fonts/NotoSansSC-VF.ttf` 与 `NotoSerifSC-VF.ttf` 按 SIL Open Font License 1.1 随包分发；许可证见 `assets/fonts/OFL-1.1.txt`。
- 不随包分发 Windows 系统字体、方正商业字体、来源不明或经过修改但没有授权文件的 GB2312 字体。
- 不要因为电脑里“能用”就推断字体可公开上传。
- 字体文件随 Skill 提供不代表已经嵌入 Word。需要跨机器完全一致时，可让使用者安装 `assets/fonts/` 中的 Noto 字体，或使用已获授权的正式单位模板。
