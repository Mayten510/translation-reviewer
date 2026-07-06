---
name: translation-reviewer
description: Review translated Excel files for quality across 10 dimensions (missing translations, mistranslations, consistency, terminology, spelling, grammar, sensitive words, technical accuracy, style compliance, idiomaticity). Use this whenever the user wants to review, audit, or check the quality of translated content, especially after running the excel-multilingual-translator. Triggers on "review translation", "检查翻译", "翻译审校", "audit translation", "翻译审查", "check translation quality", "翻译质量检查".
---

# Translation Reviewer

Review translated content (Excel files) across 10 quality dimensions. Outputs a structured audit report with issue location, description, severity, and category.

## Input

- A translated Excel file (`.xlsx`), typically the output of `excel-multilingual-translator`
- Works for any target language, auto-detected from cell content
- Can also review plain text, CSV, or any text-based format

## 10 Review Dimensions

Perform the review in two passes to avoid interference between dimensions:

### Pass A — Automated (script-based)

Run `scripts/scan.py <file.xlsx>` to detect:

| # | Dimension | What It Checks | How |
|---|-----------|---------------|-----|
| 1 | **漏翻** Missing | Cells still containing Chinese characters (CJK) | Regex scan for `[一-鿿]` |
| 5 | **拼写错误** Spelling | Misspelled words in each target language | `enchant`/`pyspellchecker` for EN; rule-based for JA/KO/DE/FR/ES |

### Pass B — Manual (Claude reviews each text)

For each unique translated text, evaluate against dimensions 2–10:

| # | Dimension | What to Check |
|---|-----------|--------------|
| 2 | **错译** Mistranslation | Does the meaning match what you'd expect from a Chinese→LANG translation? Even without source, flag obviously wrong meanings (e.g., "Save" where context suggests "Delete") |
| 3 | **一致性** Consistency | Same Chinese term → same translation everywhere. Flag when the same concept uses different English terms (e.g., "User" vs "Account" for the same Chinese source concept) |
| 4 | **术语错误** Terminology | Translation uses non-standard terms. Check against Microsoft product terminology: OK not Confirm, Save not Store, Settings not Configuration, Cancel not Abort, File not Document |
| 6 | **语法错误** Grammar | Subject-verb agreement, tense errors, article usage, word order. For JA: particle errors (は/が/を); KO: honorific level consistency; DE: article/case/gender; FR/ES: gender/number agreement |
| 7 | **敏感词** Sensitive | Flag terms that could cause legal/political/cultural issues: Taiwan/China references, disputed territories, politically charged terms, offensive language, non-inclusive terms |
| 8 | **技术原理** Technical Accuracy | Technical descriptions must be technically sound. Flag: impossible operations, incorrect protocol descriptions, misleading security claims, wrong technical cause-effect |
| 9 | **翻译规范** Style | Does the translation follow Microsoft UI style? Concise? Imperative mood for actions? Sentence case for labels? No unnecessary punctuation? Variables/placeholders preserved? |
| 10 | **地道性** Idiomaticity | Does it read naturally, or does it feel like translationese? Flag: 逐字翻译 (word-for-word), 欧化中文 patterns carried over, unnatural word order, stilted phrasing |

### Review Strategy

1. **Group by sheet and language** — review each sheet's content together to catch cross-cell consistency issues
2. **Read in context** — adjacent cells often form sentences/logical groups; review them as a unit
3. **Prioritize high-density areas** — sheets with many cells need more scrutiny than sparse sheets
4. **Compare parallel structures** — if rows 2-10 are similar patterns, check that all are translated consistently

## Severity Levels

| Level | Label | Criteria |
|-------|-------|----------|
| **P0** | Critical | Misleading/harmful translation, security risk, legal exposure, offensive content |
| **P1** | High | Meaning is wrong or confusing, technical inaccuracy, broken consistency |
| **P2** | Medium | Awkward phrasing, minor terminology issue, style deviation |
| **P3** | Low | Could be more idiomatic, minor style preference, optional improvement |

## Output Format

After completing both passes, present results as a **markdown table**:

```
| # | 问题位置 | 问题描述 | 严重程度 | 问题类别 |
|---|---------|---------|---------|---------|
| 1 | Cover!A15 | "excerpted, reproduced" → redundant; "extracted" is more idiomatic | P3 | 地道性 |
| 2 | Content!H14 | "BMC has no service on this port" → technically: "BMC does not listen on this port" | P1 | 技术原理 |
| 3 | Content!J4 | "N/A" used in row 4, "None" used in row 6 for same context — inconsistent | P2 | 一致性 |
```

**问题位置 format**: `SheetName!CellCoord` (e.g., `Cover!A15`, `Content!H14`) or `SheetName!Range` for merged cells.

Also provide a **summary section** at the end:

```
## 审查总结

| 维度 | P0 | P1 | P2 | P3 | 合计 |
|------|----|----|----|----|------|
| 漏翻 | 0 | 0 | 0 | 0 | 0 |
| 错译 | 0 | 1 | 0 | 0 | 1 |
| 一致性 | 0 | 0 | 1 | 0 | 1 |
| ... | ... | ... | ... | ... | ... |
| **总计** | **0** | **2** | **3** | **1** | **6** |

**整体评估**: [一句话总结]
**是否建议发布**: [是/否/修改后发布]
```

## Workflow

### Step 1: Extract and Scan

```bash
python <skill-path>/scripts/scan.py <translated-file.xlsx>
```

This outputs:
- All unique translated texts grouped by sheet and language
- Automated check results (CJK residual → 漏翻, spellcheck → 拼写)

### Step 2: Review Each Unique Text

For each sheet and language group, review the extracted texts against dimensions 2–10. Think aloud as you review — mention what you're checking and why.

### Step 3: Compile Report

Combine automated findings with manual review results. Format as the markdown table above. Sort by severity (P0 first), then by dimension.

### Step 4: Export to Excel

After compiling findings as a JSON array, generate the Excel report:

```bash
python <skill-path>/scripts/export_review.py <findings.json> <output.xlsx>
```

The JSON format:
```json
[
  {"location": "Cover!A15", "description": "...", "severity": "P3", "dimension": "地道性"},
  {"location": "Content!H11", "description": "...", "severity": "P1", "dimension": "技术原理"}
]
```

The output Excel contains two sheets:
- **审查结果**: Detailed findings table (问题位置 | 问题描述 | 严重程度 | 问题类别) with color-coded severity
- **审查总结**: Summary counts per dimension with overall assessment

### Step 5: Provide Recommendations

For each P0/P1 issue, suggest a corrected translation. For P2/P3, optionally suggest improvements.

## Language-Specific Review Notes

### English
- Check for Chinese punctuation (，。"" instead of ,."")
- Check for Chinglish patterns: "according to the operation" → "as follows", "please operate" → "click/tap"
- Capitalization: sentence case for labels, title case for proper nouns only

### Japanese
- Katakana vs hiragana for loanwords: technical terms → katakana (サーバー, プロトコル)
- Politeness level: use です・ます調 for instructional text consistently
- Avoid 漢語 (kango) overload — too many Chinese-derived compounds is unnatural

### Korean
- Honorific consistency: 합쇼체 for formal, 해요체 for semi-formal
- Spacing rules (띄어쓰기): Korean uses spaces between words
- English loanword spelling: follow 표준국어대사전 conventions

### German
- Formal "Sie" vs informal "du" — must be consistent
- Compound noun correctness: Benutzerhandbuch not Benutzer Handbuch
- Article/case agreement

### French
- Formal "vous" vs informal "tu" — must be consistent
- Gender agreement for adjectives and past participles
- Accent marks: ensure all required accents are present (é, è, ê, ë, à, â, ù, û, î, ï, ô, ç)

### Spanish
- Formal "usted" vs informal "tú" — must be consistent
- Gender/number agreement for adjectives
- Accent marks and ñ: ensure present where required

## Bundled Scripts

| Script | Purpose |
|--------|---------|
| `scripts/scan.py` | Automated first-pass: CJK残留检测, 敏感词扫描, 文本提取分组 |
| `scripts/export_review.py` | Export findings JSON to formatted Excel report with color-coded severity |

Use `scan.py` first to extract texts and run automated checks. Then Claude performs manual review. Finally, `export_review.py` generates the formatted Excel report.
