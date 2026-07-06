# Translation Reviewer

A Claude Code skill for reviewing translated Excel files across 10 quality dimensions.

## Review Dimensions

| # | Dimension | Automated | Description |
|---|-----------|-----------|-------------|
| 1 | 漏翻 (Missing) | ✅ | Detects untranslated Chinese characters (CJK) |
| 2 | 错译 (Mistranslation) | — | Flags semantically incorrect translations |
| 3 | 一致性 (Consistency) | — | Same source term → same translation everywhere |
| 4 | 术语错误 (Terminology) | — | Non-standard terms vs Microsoft/industry conventions |
| 5 | 拼写错误 (Spelling) | — | Misspelled words in target language |
| 6 | 语法错误 (Grammar) | — | Agreement, tense, word order, particles |
| 7 | 敏感词 (Sensitive) | ✅ | Political, legal, or offensive terms |
| 8 | 技术原理 (Technical) | — | Descriptions must be technically sound |
| 9 | 翻译规范 (Style) | — | Microsoft UI style, conciseness, imperative mood |
| 10 | 地道性 (Idiomaticity) | — | Reads naturally, not translationese |

## Output

Excel report with two sheets:
- **审查结果**: Color-coded findings table (位置 | 描述 | 严重程度 | 类别)
- **审查总结**: Summary counts per dimension

## Severity Levels

| Level | Label | Color |
|-------|-------|-------|
| P0 | Critical | 🔴 Red |
| P1 | High | 🟠 Orange |
| P2 | Medium | 🟡 Yellow |
| P3 | Low/Style | 🔵 Blue |

## Usage

```
Review the translation quality of this Excel file
```

Or programmatically:

```bash
# Step 1: Automated scan
python scripts/scan.py translated.xlsx

# Step 2: Manual review by Claude (follows SKILL.md)

# Step 3: Export findings to Excel
python scripts/export_review.py findings.json review_report.xlsx
```

## Requirements

- Python 3.8+
- openpyxl (`pip install openpyxl`)

## Related

- [excel-multilingual-translator](https://github.com/Mayten510/excel-multilingual-translator) — translate Chinese Excel to 6 languages
