# Translation Reviewer

A Claude Code skill for reviewing translated content across **10 quality dimensions** and **7 document formats**.

## Supported Formats

| Format | Single-file | Comparative (source↔target) |
|--------|-------------|---------------------------|
| Excel (.xlsx) | ✅ cell-level | ✅ cell-by-cell alignment |
| Word (.docx) | ✅ paragraph/table | ✅ section alignment |
| PowerPoint (.pptx) | ✅ slide-level | ✅ section alignment |
| PDF (.pdf) | ✅ page/line | ✅ section alignment |
| Markdown (.md) | ✅ line-level | ✅ section alignment |
| HTML (.html) | ✅ tag text | ✅ section alignment |
| Plain Text (.txt) | ✅ line-level | ✅ section alignment |

## Review Modes

### Single-File Review
Quick quality check of translated file — spelling, grammar, style, idiomaticity, sensitive words.

### Comparative Review (source↔target)
Full accuracy review with source-target alignment. Detects omissions, additions, mistranslations, and consistency issues.

## 10 Quality Dimensions

1. 漏翻 (Missing) - automated CJK detection
2. 错译 (Mistranslation)
3. 一致性 (Consistency)
4. 术语错误 (Terminology)
5. 拼写错误 (Spelling)
6. 语法错误 (Grammar)
7. 敏感词 (Sensitive) - automated
8. 技术原理 (Technical Accuracy)
9. 翻译规范 (Style)
10. 地道性 (Idiomaticity)

## Output

Excel report with color-coded severity (P0🔴 P1🟠 P2🟡 P3🔵):
- **审查结果**: Detailed findings
- **审查总结**: Per-dimension summary

## Usage

```bash
# Single-file review
python scripts/extract.py translated.pdf

# Comparative review
python scripts/compare.py source-cn.pdf target-en.pdf

# Export findings to Excel
python scripts/export_review.py findings.json report.xlsx
```

## Requirements

- Python 3.8+
- Libraries: openpyxl, python-docx, python-pptx, pdfplumber, beautifulsoup4, lxml
- Install: `pip install openpyxl python-docx python-pptx pdfplumber beautifulsoup4 lxml`

## Related

- [excel-multilingual-translator](https://github.com/Mayten510/excel-multilingual-translator) — translate Chinese Excel to 6 languages
