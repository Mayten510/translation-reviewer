---
name: translation-reviewer
description: Review translated content for quality across 10 dimensions. Supports two modes — single-file review (check translated output alone) and comparative review (source↔target side-by-side). Use whenever the user wants to review, audit, or check translation quality. Triggers on "review translation", "检查翻译", "翻译审校", "audit translation", "翻译审查", "check translation quality", "翻译质量检查", "对比原文和译文", "compare translation".
---

# Translation Reviewer

Two review modes:

| Mode | Input | Best For |
|------|-------|----------|
| **Single-file** | Translated file only | Quick quality check, spelling/grammar, style, idiomaticity |
| **Comparative** | Source (Chinese) + Translated file | Full accuracy review: catch mistranslations, omissions, additions, meaning shifts |

## Mode A: Single-File Review (默认)

Review a translated file without the original. See [Single-File Workflow](#single-file-workflow).

## Mode B: Comparative Review (推荐)

When both source (Chinese) and translated files are available, gets source↔target sentence-level alignment and check every segment.

### Comparative Dimensions

In addition to all 10 single-file dimensions, comparative mode adds:

| # | Dimension | What to Check |
|---|-----------|--------------|
| 11 | **增译** Addition | Text in translation that has no corresponding source text |
| 12 | **漏译** Omission | Source content missing from the translation |

And strengthens:

| # | Dimension | Enhanced Check |
|---|-----------|---------------|
| 2 | **错译** | Now verifiable: compare source meaning vs translation meaning directly |
| 3 | **一致性** | Check that same Chinese term → same English term across the document |
| 8 | **技术原理** | Verify that technical facts are preserved correctly from source |

### Comparative Workflow

#### Step 1: Align Source and Target

For **Excel files**, run:
```bash
python <skill-path>/scripts/compare.py <source-zh.xlsx> <target-en.xlsx>
```

This outputs a JSON with aligned cell pairs:
```json
{
  "sheet_name": "首页",
  "pairs": [
    {"coordinate": "B2", "source": "保存", "target": "Save", "match": "exact"},
    {"coordinate": "B5", "source": "操作失败，请重试", "target": "Operation failed. Please try again.", "match": "exact"},
    {"coordinate": "B8", "source": "确定要删除所选项目吗？此操作不可撤销。", "target": null, "match": "missing"}
  ]
}
```

For **PDF/text files**, extract text from both and align by paragraph/section manually.

#### Step 2: Review Each Pair

For each aligned pair, check:

1. **漏译 (Omission)**: `target` is null/empty but `source` has content → P1
2. **增译 (Addition)**: `target` has content but `source` is null (unless header/metadata) → P2
3. **错译 (Mistranslation)**: Source and target meanings diverge → P1/P2 depending on severity
4. **术语错误**: Compare source terminology against target — must match Microsoft/industry standard
5. **一致性**: Same source term → same translation throughout. The comparison script groups identical source texts, making this easy to verify
6. **All single-file dimensions**: Grammar, spelling, style, idiomaticity on the target text

#### Step 3: Compile and Export

Same as single-file mode — compile findings and export via `export_review.py`.

---

## Single-File Workflow

### Step 1: Extract and Scan

```bash
python <skill-path>/scripts/scan.py <translated-file.xlsx>
```

Outputs:
- All unique translated texts grouped by sheet and language
- Automated checks: CJK residual (漏翻) + sensitive words (敏感词)

### Step 2: Manual Review

Review extracted texts against dimensions 2–10.

### Step 3: Export

```bash
python <skill-path>/scripts/export_review.py <findings.json> <output.xlsx>
```

---

## 10 Review Dimensions

### Pass A — Automated (script-based)

| # | Dimension | How |
|---|-----------|-----|
| 1 | **漏翻** Missing | Regex scan for `[一-鿿]` in target |
| 7 | **敏感词** Sensitive | Regex patterns for political/legal/offensive terms |

### Pass B — Manual (Claude reviews)

| # | Dimension | What to Check |
|---|-----------|--------------|
| 2 | **错译** Mistranslation | Source meaning vs target meaning (comparative mode makes this precise) |
| 3 | **一致性** Consistency | Same source term → same translation everywhere |
| 4 | **术语错误** Terminology | Non-standard terms vs Microsoft/industry conventions |
| 5 | **拼写错误** Spelling | Misspelled words in target language |
| 6 | **语法错误** Grammar | Agreement, tense, word order, particles, articles |
| 8 | **技术原理** Technical Accuracy | Technical descriptions must be correct |
| 9 | **翻译规范** Style | Microsoft UI style, conciseness, imperative mood, sentence case |
| 10 | **地道性** Idiomaticity | Reads naturally — not translationese |

### Review Strategy

1. **Group and compare** — review each sheet's content together
2. **Read in context** — adjacent cells form logical groups
3. **Parallel structures** — if rows 2-10 are similar, check all consistently
4. **Comparative mode**: Flag when `match: "missing"` (漏译) or `source: null, target: "text"` (增译)

## Severity Levels

| Level | Label | Criteria |
|-------|-------|----------|
| **P0** | Critical | Harmful/offensive, security risk, legal exposure |
| **P1** | High | Meaning wrong, technical inaccuracy, critical omission |
| **P2** | Medium | Awkward phrasing, minor terminology, style deviation |
| **P3** | Low | Could be more idiomatic, minor preference |

## Output Format

Markdown table + Excel report (via `export_review.py`):

```
| # | 问题位置 | 问题描述 | 严重程度 | 问题类别 |
|---|---------|---------|---------|---------|
| 1 | Cover!A15 | "excerpted, reproduced" → redundant | P3 | 地道性 |
```

**问题位置**: File-format specific:
- Excel: `SheetName!CellCoord` or `SheetName!Range`
- PDF/Text: `Page L{line}` or `Section "title"`

Include a summary section with per-dimension counts and overall assessment.

## Language-Specific Review Notes

### English
- Chinese punctuation (，。"" → ,."")
- Chinglish: "according to the operation" → "as follows"
- Sentence case for labels, title case for proper nouns only

### Japanese (JA)
- Technical terms → katakana (サーバー, プロトコル)
- Politeness: です・ます調 for instructions; avoid 漢語 overload

### Korean (KO)
- Honorific: 합쇼체 (formal) vs 해요체 (semi-formal); 띄어쓰기 spacing

### German (DE)
- Formal "Sie" vs "du"; compound noun correctness; article/case agreement

### French (FR)
- Formal "vous" vs "tu"; gender agreement; accent marks (éèêëàâùûîïôç)

### Spanish (ES)
- Formal "usted" vs "tú"; gender/number agreement; accents and ñ

## Bundled Scripts

| Script | Purpose |
|--------|---------|
| `scripts/scan.py` | Single-file: CJK残留检测, 敏感词扫描, 文本提取 |
| `scripts/compare.py` | Comparative: align source↔target cells, detect omissions/additions |
| `scripts/export_review.py` | Export findings JSON → formatted Excel report (both modes) |
