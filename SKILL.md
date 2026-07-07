---
name: translation-reviewer
description: Review translated content for quality across 10 dimensions. Supports single-file review and source↔target comparative review. Works with XLSX, DOCX, PPTX, PDF, MD, HTML, TXT formats. Use whenever the user wants to review, audit, or check translation quality. Triggers on "review translation", "检查翻译", "翻译审校", "audit translation", "翻译审查", "check translation quality", "翻译质量检查", "对比原文和译文", "compare translation".
---

# Translation Reviewer

Two review modes across 7 document formats:

| Mode | Input | Best For |
|------|-------|----------|
| **Single-file** | Translated file only | Quick quality check, spelling/grammar, style, idiomaticity |
| **Comparative** | Source (CN) + Translated file | Full accuracy: mistranslations, omissions, additions, meaning shifts |

## Supported Formats

| Format | Extensions | Single-file | Comparative |
|--------|-----------|-------------|-------------|
| Excel | `.xlsx` `.xls` | ✅ cell-level | ✅ cell-by-cell alignment |
| Word | `.docx` `.doc` | ✅ paragraph/table | ✅ section alignment |
| PowerPoint | `.pptx` `.ppt` | ✅ slide-level | ✅ section alignment |
| PDF | `.pdf` | ✅ page/line | ✅ section alignment |
| Markdown | `.md` `.markdown` | ✅ line-level | ✅ section alignment |
| HTML | `.html` `.htm` | ✅ tag text | ✅ section alignment |
| Plain Text | `.txt` `.csv` | ✅ line-level | ✅ section alignment |

## Mode A: Single-File Review

### Step 1: Extract

```bash
python <skill-path>/scripts/extract.py <file>
```

Auto-detects format, extracts all text sections with location metadata, runs automated checks (CJK residual, sensitive words).

### Step 2: Manual Review

Claude reviews extracted texts against 10 dimensions.

### Step 3: Export

```bash
python <skill-path>/scripts/export_review.py <findings.json> <output.xlsx>
```

## Mode B: Comparative Review

### Step 1: Align and Compare

```bash
python <skill-path>/scripts/compare.py <source-cn> <target-translated>
```

Auto-detects format:
- **Excel**: cell-by-cell alignment with omission/addition/consistency detection
- **Other formats**: parallel extraction with size ratio warning

### Step 2: Side-by-Side Review

For Excel: review aligned pairs (source→target). Flag `missing`/`addition`/`modified` pairs.

For PDF/DOCX/PPTX: extract both, read source sections first, then review target. Pay attention to:
- Added content not in source (增译)
- Missing content from source (漏译)
- Meaning shifts in translated sections
- Size mismatch warning (e.g., "Target is 2.4x larger")

### Step 3: Export

Same as single-file mode.

---

## 10 Review Dimensions

| # | Dimension | Auto | What to Check |
|---|-----------|------|--------------|
| 1 | **漏翻** Missing | ✅ CJK regex | Chinese characters remaining in target |
| 2 | **错译** Mistranslation | — | Source vs target meaning divergence |
| 3 | **一致性** Consistency | ✅ Excel only | Same source → same translation everywhere |
| 4 | **术语错误** Terminology | — | Non-standard vs Microsoft/industry terms |
| 5 | **拼写错误** Spelling | — | Misspelled words in target language |
| 6 | **语法错误** Grammar | — | Agreement, tense, articles, word order |
| 7 | **敏感词** Sensitive | ✅ regex | Political/legal/offensive terms |
| 8 | **技术原理** Technical | — | Technical descriptions must be correct |
| 9 | **翻译规范** Style | — | Microsoft UI style, conciseness, case |
| 10 | **地道性** Idiomaticity | — | Reads naturally, not translationese |

Comparative mode adds:
| # | Dimension | What to Check |
|---|-----------|--------------|
| 11 | **增译** Addition | Target text with no source counterpart |
| 12 | **漏译** Omission | Source content missing from target |

## Severity Levels

| Level | Color | Criteria |
|-------|-------|----------|
| **P0** | 🔴 Red | Harmful/offensive, security, legal exposure |
| **P1** | 🟠 Orange | Meaning wrong, technical inaccuracy, critical omission |
| **P2** | 🟡 Yellow | Awkward phrasing, minor terminology, style deviation |
| **P3** | 🔵 Blue | Idiomaticity, minor preference |

## Output Format

Excel report (via `export_review.py`) with two sheets:
- **审查结果**: findings table with color-coded severity (位置 | 描述 | 严重程度 | 类别)
- **审查总结**: per-dimension counts + overall assessment

Location format by document type:
- Excel: `Sheet!CellCoord`
- PDF: `Page N, Line M`
- DOCX: `Paragraph N` or `Table N, Row M`
- PPTX: `Slide N, Paragraph M`
- MD/HTML/TXT: `Line N`

Also output markdown summary in conversation.

## Language-Specific Notes

**English**: Chinese punctuation (，。→ ,.) ・ Chinglish (according to operation → as follows) ・ sentence case labels

**Japanese**: Katakana for loanwords ・ です・ます調 consistency ・ avoid 漢語 overload

**Korean**: 합쇼체/해요체 consistency ・ 띄어쓰기 spacing ・ English loanword conventions

**German**: Sie/du consistency ・ compound nouns ・ article/case agreement

**French**: vous/tu consistency ・ gender agreement ・ accent marks

**Spanish**: usted/tú consistency ・ gender/number agreement ・ accents/ñ

## Bundled Scripts

| Script | Purpose |
|--------|---------|
| `scripts/extract.py` | Universal extractor: xlsx/docx/pptx/pdf/md/html/txt → JSON with auto-checks |
| `scripts/compare.py` | Multi-format source↔target alignment: Excel cell-level, generic section-level |
| `scripts/export_review.py` | Findings JSON → formatted Excel report (P0-P3 color-coded) |
