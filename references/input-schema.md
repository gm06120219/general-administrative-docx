# DOCX 生成器输入

## 最小示例

```json
{
  "title": "关于开展业务培训的通知",
  "unit_name": "示例单位办公室",
  "date": "2026年9月28日",
  "recipient": "各部门：",
  "blocks": [
    {"type": "paragraph", "text": "为进一步提升工作质量，现就有关事项通知如下。"},
    {"type": "heading1", "text": "培训安排"},
    {"type": "paragraph", "text": "培训时间、地点和参加人员另行通知。"}
  ]
}
```

## 顶层字段

| 字段 | 必需 | 说明 |
|---|---:|---|
| `title` | 是 | 文档标题，也用于默认文件名。 |
| `mode` | 否 | `ordinary`（默认）或 `redhead`。 |
| `document_type` | 否 | 通知、报告、请示、函、纪要、方案等。 |
| `recipient` | 否 | 主送对象；生成器不会自动补冒号。 |
| `unit_name` | 否 | 落款单位。也可在 `signature.unit_name` 单独指定。 |
| `date` | 否 | 落款日期；缺省使用当天中文日期。 |
| `version` | 否 | 默认 `1.0`，输出显示为 `V1.0`。 |
| `page_numbers` | 否 | 普通模式默认 `false`，红头模式默认 `true`。 |
| `blocks` | 是 | 正文块数组。 |
| `signature` | 否 | `{"unit_name": "...", "date": "..."}`；设为 `false` 可取消落款。 |
| `fonts` | 否 | 字体覆盖，键为 `title/body/h1/h2/h3/h4/page_number`。 |

## 正文块

- `paragraph`：普通正文；可用 `first_line_indent_chars` 覆盖默认 2 字符。
- `heading1`、`heading2`、`heading3`、`heading4`：标题。默认自动加规范序号；若文本已含对应序号则不重复。
- `names`：`names` 为姓名数组；可用 `per_line` 覆盖默认每行 5 人。
- `attachment`：附件说明，正文内容放在 `text`。
- `page_break`：插入分页符。
- `blank`：插入一个空行，仅在版面确有需要时使用。

块可设置 `auto_number: false` 禁用自动序号。标题计数会在进入新的上级标题后重置下级计数。

## 文件命名

默认输出为 `标题-YYYYMMDD-VN.N.docx`。非法文件名字符会替换为全角或安全字符；若文件已存在，自动递增小版本，如 `V1.0` → `V1.1`，除非显式使用 `--overwrite`。
