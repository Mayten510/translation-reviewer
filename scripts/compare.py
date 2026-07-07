#!/usr/bin/env python3
"""
Align source (Chinese) and target (translated) Excel files cell-by-cell.
Detects omissions, additions, and groups identical source texts for consistency review.

Usage:
    python compare.py <source-zh.xlsx> <target-en.xlsx>
"""

import json
import re
import sys
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is required. Install with: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


def has_chinese(text):
    """Check if text contains Chinese characters."""
    if not text or not isinstance(text, str):
        return False
    return bool(re.search(r'[一-鿿]', text))


def extract_cells(ws):
    """Extract all non-empty cells from a sheet."""
    cells = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None:
                val = str(cell.value).strip()
                if val:
                    cells[cell.coordinate] = val
    return cells


def determine_match(source_val, target_val):
    """
    Classify the source→target relationship.
    Returns match type string.
    """
    source_has_cn = has_chinese(source_val) if source_val else False
    source_empty = not source_val
    target_empty = not target_val

    if source_empty and target_empty:
        return "both_empty"
    if source_empty and not target_empty:
        return "addition"      # Text in target with no source
    if not source_empty and target_empty:
        if source_has_cn:
            return "missing"   # Chinese source not translated
        else:
            return "source_only"  # Non-Chinese source not carried over
    # Both have content
    if source_has_cn:
        return "translated"    # Chinese → translated
    else:
        if source_val == target_val:
            return "unchanged"  # Same non-Chinese content
        else:
            return "modified"   # Non-Chinese content changed (unusual)


def compare_sheets(ws_source, ws_target):
    """Compare two sheets cell by cell."""
    source_cells = extract_cells(ws_source)
    target_cells = extract_cells(ws_target)

    all_coords = set(source_cells.keys()) | set(target_cells.keys())
    pairs = []

    stats = defaultdict(int)

    def col_sort_key(coord):
        """Convert coordinate like 'A15' to (col_letter_value, row_number)."""
        m = re.match(r'([A-Z]+)(\d+)', coord)
        if not m:
            return (0, 0)
        col_letters = m.group(1)
        # Convert column letters to number: A=1, B=2, ..., Z=26, AA=27, ...
        col_num = 0
        for ch in col_letters:
            col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
        return (col_num, int(m.group(2)))

    for coord in sorted(all_coords, key=col_sort_key):
        source_val = source_cells.get(coord)
        target_val = target_cells.get(coord)
        match_type = determine_match(source_val, target_val)

        pair = {
            "coordinate": coord,
            "source": source_val,
            "target": target_val,
            "match": match_type,
        }
        pairs.append(pair)
        stats[match_type] += 1

    return pairs, dict(stats)


def group_by_source(pairs):
    """
    Group identical source texts for consistency checking.
    Returns {source_text: [{coordinate, target, match}, ...]}
    """
    groups = defaultdict(list)
    for p in pairs:
        if p["source"] and has_chinese(p["source"]):
            groups[p["source"]].append({
                "coordinate": p["coordinate"],
                "target": p["target"],
                "match": p["match"],
            })

    # Only keep groups with multiple occurrences (relevant for consistency)
    consistency_issues = []
    for source_text, occurrences in groups.items():
        if len(occurrences) > 1:
            targets = [o["target"] for o in occurrences if o["target"]]
            unique_targets = set(targets)
            if len(unique_targets) > 1:
                consistency_issues.append({
                    "source": source_text,
                    "occurrences": occurrences,
                    "unique_translations": list(unique_targets),
                    "issue": f"Same source translated {len(unique_targets)} different ways",
                })

    return {
        "total_unique_sources": len(groups),
        "consistency_issues": consistency_issues,
    }


def main():
    if len(sys.argv) < 3:
        print("Usage: python compare.py <source-zh.xlsx> <target-en.xlsx>", file=sys.stderr)
        sys.exit(1)

    source_path = Path(sys.argv[1])
    target_path = Path(sys.argv[2])

    for p in [source_path, target_path]:
        if not p.exists():
            print(f"Error: file not found: {p}", file=sys.stderr)
            sys.exit(1)

    wb_src = openpyxl.load_workbook(source_path, data_only=True)
    wb_tgt = openpyxl.load_workbook(target_path, data_only=True)

    result = {
        "source_file": str(source_path.absolute()),
        "target_file": str(target_path.absolute()),
        "sheets": [],
        "global_stats": {},
        "consistency": {},
        "automated_findings": [],
    }

    global_stats = defaultdict(int)
    all_pairs = []

    for sheet_name in wb_src.sheetnames:
        # Find matching sheet in target (may have translated name)
        ws_src = wb_src[sheet_name]
        ws_tgt = None

        # Try exact match first
        if sheet_name in wb_tgt.sheetnames:
            ws_tgt = wb_tgt[sheet_name]
        else:
            # Try fuzzy: target sheet may be suffixed or translated
            for tgt_name in wb_tgt.sheetnames:
                if tgt_name.startswith(sheet_name) or sheet_name.startswith(tgt_name.rstrip('_EN_JA_KO_DE_FR_ES')):
                    ws_tgt = wb_tgt[tgt_name]
                    break
            # Fallback: use first sheet or by position
            if ws_tgt is None:
                src_idx = wb_src.sheetnames.index(sheet_name)
                if src_idx < len(wb_tgt.sheetnames):
                    ws_tgt = wb_tgt[wb_tgt.sheetnames[src_idx]]

        if ws_tgt is None:
            result["automated_findings"].append({
                "dimension": "漏翻",
                "severity": "P1",
                "location": f"Sheet: {sheet_name}",
                "description": f"No matching sheet found in target for '{sheet_name}'",
            })
            continue

        sheet_pairs, sheet_stats = compare_sheets(ws_src, ws_tgt)
        all_pairs.extend(sheet_pairs)

        for k, v in sheet_stats.items():
            global_stats[k] += v

        # Generate automated findings for this sheet
        for p in sheet_pairs:
            if p["match"] == "missing":
                result["automated_findings"].append({
                    "dimension": "漏译",
                    "severity": "P1" if has_chinese(p["source"]) and len(p["source"]) > 3 else "P2",
                    "location": f"{sheet_name}!{p['coordinate']}",
                    "description": f"Source '{p['source'][:40]}' not translated (target is empty)",
                })
            elif p["match"] == "addition":
                if len(p["target"]) > 3:  # Skip trivial additions
                    result["automated_findings"].append({
                        "dimension": "增译",
                        "severity": "P2",
                        "location": f"{sheet_name}!{p['coordinate']}",
                        "description": f"Added text '{p['target'][:40]}' with no source counterpart",
                    })

        result["sheets"].append({
            "source_name": sheet_name,
            "target_name": ws_tgt.title if ws_tgt else "N/A",
            "stats": sheet_stats,
            "pair_count": len(sheet_pairs),
        })

    wb_src.close()
    wb_tgt.close()

    # Consistency analysis
    consistency = group_by_source(all_pairs)
    result["consistency"] = consistency

    # Report consistency issues
    for issue in consistency["consistency_issues"]:
        translations = issue["unique_translations"]
        coords = [o["coordinate"] for o in issue["occurrences"]]
        result["automated_findings"].append({
            "dimension": "一致性",
            "severity": "P1" if len(translations) > 2 else "P2",
            "location": ", ".join(coords[:5]),
            "description": f"'{issue['source'][:30]}' translated inconsistently: {translations}",
        })

    result["global_stats"] = dict(global_stats)
    result["total_pairs"] = len(all_pairs)
    result["total_automated_findings"] = len(result["automated_findings"])

    # Group findings by dimension
    by_dim = defaultdict(list)
    for f in result["automated_findings"]:
        by_dim[f["dimension"]].append(f)
    result["automated_findings_by_dimension"] = {k: v for k, v in by_dim.items()}

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
