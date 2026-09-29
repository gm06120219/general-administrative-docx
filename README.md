# 通用行政公文编制 Skill

这是一个可直接放入 GitHub 仓库并安装到 Codex/Agent 环境的 Skill。它把行政公文内容生成、版式约束和 DOCX 导出封装为可重复流程，默认生成非红头 Word 文件。

## 安装

```powershell
npx skills add https://github.com/<owner>/<repo> --skill general-administrative-docx
```

本地开发时，也可以把整个 `general-administrative-docx` 目录复制到个人 skills 目录。不要只复制 `SKILL.md`。

## 运行生成器

```powershell
python -m pip install -r requirements.txt
python scripts/generate_docx.py --input examples/sample-input.json --output-dir examples
python scripts/validate_docx.py examples/<生成的文件名>.docx
```

输入字段见 `references/input-schema.md`。最终交付物是 `.docx`，JSON 只用于生成过程。

## 字体

`assets/fonts/` 附带可公开再分发的 Noto Sans SC / Noto Serif SC 兜底字体和 OFL 1.1 许可证。方正、Windows 系统字体及来源不明的 GB2312 字体没有被打包。详情见 `references/fonts.md`。

## 发布前

- 运行 `scripts/validate_docx.py`。
- 用 Word 或 LibreOffice 打开并逐页检查版面。
- 根据你的仓库策略补充项目级许可证；字体始终受 `assets/fonts/OFL-1.1.txt` 约束。
