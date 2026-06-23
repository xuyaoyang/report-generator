#!/usr/bin/env python
"""
Material certificate params cleaner — extract params from filenames,
normalize grades, and update the database.

Usage:
  python tools/clean_material_params.py --dry-run     # preview only
  python tools/clean_material_params.py               # apply updates
  python tools/clean_material_params.py --category 钢板 # single category
"""

import sqlite3
import json
import re
import os
import sys
import argparse
from datetime import datetime
from collections import defaultdict


# ====================================================================
# Extraction rules per category
# Each rule: (priority, regex, extract_fn)
# extract_fn(match, filename) -> dict of param updates
# Returns None if the match doesn't actually yield valid params
# ====================================================================

def _num(s):
    """Normalize a number string."""
    try:
        n = float(s)
        if n == int(n):
            return str(int(n))
        return str(n)
    except (ValueError, TypeError):
        return s.strip()

def _grade_normalize(g):
    """Normalize common grade aliases to canonical forms."""
    g = str(g).strip().upper()
    # Steel plate grades
    if g in ('Q355', 'Q355BB', 'Q355C', 'Q355D'):
        return 'Q355B'
    if g in ('Q235', 'Q235A', 'Q235C', 'Q235D'):
        return 'Q235B'
    if g in ('235', '235B'):
        return 'Q235B'
    if g in ('355', '355B'):
        return 'Q355B'
    if g in ('LY',):
        return None  # ambiguous, need subtype
    # Rebar grades
    g_upper = g
    if re.match(r'^HRB\d+E?$', g_upper):
        return g_upper
    if g_upper in ('三级', '三级钢', 'HRB400'):
        return 'HRB400E'
    # Bolt grades
    if re.match(r'^\d+\.\d+$', g_upper):
        return g_upper
    return g_upper


# ── 钢板 (Steel Plate) ─────────────────────────────────────────────

def _plate_extract(filename):
    """Try to extract (thickness, material_grade) from a steel plate filename.
    Returns dict or None."""
    f = filename
    result = {}

    # Priority 1: "20厚钢板235B" or "20厚钢板 235B"
    m = re.search(r'(\d+\.?\d*)\s*厚钢板\s*(\d+[A-Za-z]*)', f)
    if m:
        return {'thickness': _num(m.group(1)),
                'material_grade': _grade_normalize(m.group(2))}

    # Priority 2: "25-Q355B", "25_Q355B", "20_Q355B_202502" etc.
    m = re.search(r'(\d+\.?\d*)[_\-]((?:Q\d+[A-Za-z]*|LY\d+|Z\d+))\b', f)
    if m:
        result = {'thickness': _num(m.group(1)),
                  'material_grade': _grade_normalize(m.group(2))}
        try:
            t = float(m.group(1))
            if 0.3 <= t <= 200:
                return result
        except ValueError:
            pass

    # Priority 2b: "235b-20x2000" — grade first
    # Priority 2c: "Q235B-20-..."  — grade first with hyphen
    # Grade patterns: Q-prefixed, or known grades like 45#, 40Cr, 235B that have ≤6 chars
    m = re.search(r'\b(Q\d+[A-Za-z]*|LY\d+|Z\d+|[A-Za-z]{2,4}\d{1,3}[A-Za-z]?)\s*[-_]\s*(\d+\.?\d*)[x×X_-]', f)
    if m:
        gr = _grade_normalize(m.group(1))
        t = _num(m.group(2))
        try:
            if gr and 0.3 <= float(t) <= 200:
                return {'thickness': t, 'material_grade': gr}
        except (ValueError, TypeError):
            pass

    # Priority 2d: "Q355B-20" or "Q235B_20" (need word boundary \b at end to not
    # match "Q235B-20-2000-Q355B" wrongly, but also need to not require \b at start
    # since filename often has a prefix like uuid_ before the grade)
    m = re.search(r'(Q\d+[A-Za-z]*)[-_](\d+\.?\d*)\b', f)
    if m:
        gr = _grade_normalize(m.group(1))
        t = _num(m.group(2))
        try:
            if 0.3 <= float(t) <= 200:
                return {'thickness': t, 'material_grade': gr}
        except ValueError:
            pass

    # Priority 2e: "20-2200-Q355B" or "20-2200-235B" — thick-width-grade (3-part)
    m = re.search(r'(\d+\.?\d*)[-_](\d{3,5})[-_](\d+[A-Za-z]+|[A-Za-z]+\d+[A-Za-z]*|[A-Za-z]+[A-Za-z]*\d+)', f)
    if m:
        gr = _grade_normalize(m.group(3))
        if gr:
            return {'thickness': _num(m.group(1)),
                    'material_grade': gr}

    # Priority 2f: "Q355B16-2200" — grade directly followed by thickness, then hyphen
    # Require \b before Q to avoid matching within heat numbers like "2211MQ40419"
    m = re.search(r'\b(Q\d+[A-Za-z]*)(\d+\.?\d*)[-_]', f)
    if m:
        gr = _grade_normalize(m.group(1))
        t = _num(m.group(2))
        try:
            if 0.3 <= float(t) <= 200:
                return {'thickness': t, 'material_grade': gr}
        except ValueError:
            pass

    # Priority 3: "Q235B 20 2000" — grade then thickness then width
    m = re.search(r'(?:^|[_])?(Q\d+[A-Za-z]*)\s+(\d+\.?\d*)\s+\d+', f)
    if m:
        t = float(m.group(2))
        if 0.3 <= t <= 200:
            return {'thickness': _num(m.group(2)),
                    'material_grade': _grade_normalize(m.group(1))}

    # Priority 3b: "Q355B 10mm"
    m = re.search(r'(?:^|[_])?(Q\d+[A-Za-z]*)\s+(\d+\.?\d*)\s*mm', f)
    if m:
        t = float(m.group(2))
        if 0.3 <= t <= 200:
            return {'thickness': _num(m.group(2)),
                    'material_grade': _grade_normalize(m.group(1))}

    # Priority 3c: "Q355B 20+2000" — grade thickness+width
    m = re.search(r'(Q\d+[A-Za-z]*)\s+(\d+\.?\d*)\+', f)
    if m:
        t = float(m.group(2))
        if 0.3 <= t <= 200:
            return {'thickness': _num(m.group(2)),
                    'material_grade': _grade_normalize(m.group(1))}

    # Priority 3d: REMOVED — regex was malformed. Covered by P2d.

    # Priority 3e: "Q235B-16-20-25-30" — grade with multiple thicknesses, take the first
    m = re.search(r'(Q\d+[A-Za-z]*)[-_](\d+\.?\d*)[-_](\d+\.?\d*)', f)
    if m:
        t = float(m.group(2))
        if 0.3 <= t <= 200:
            return {'thickness': _num(m.group(2)),
                    'material_grade': _grade_normalize(m.group(1))}

    # Priority 3f: "HRB400-14-18-20" — rebar grade used in钢板 context, extract thickness
    m = re.search(r'(HRB\d+[A-Za-z]*)[-_](\d+\.?\d*)[-_](\d+\.?\d*)', f)
    if m:
        t = float(m.group(2))
        if 0.3 <= t <= 200:
            return {'thickness': _num(m.group(2))}

    # Priority 4: "低合金板25" or "普板80" or "低合金卷25" or "普卷3.5"
    m = re.search(r'(低合金板|低合金卷|普板|普卷)\s*(\d+\.?\d*)', f)
    if m:
        gr = 'Q355B' if '低合金' in m.group(1) else 'Q235B'
        return {'thickness': _num(m.group(2)), 'material_grade': gr}

    # Priority 5: "C板12" — C plate
    m = re.search(r'C板\s*(\d+\.?\d*)', f)
    if m:
        return {'thickness': _num(m.group(1))}

    # Priority 6: "10x1510" pattern (plate thickness x width)
    m = re.search(r'(?:^|[^0-9])(\d+\.?\d*)[x×X](\d{3,5})(?:[^0-9]|$)', f)
    if m:
        t = float(m.group(1))
        w = float(m.group(2))
        if 0.3 <= t <= 200 and 900 <= w <= 5000:
            return {'thickness': _num(m.group(1))}

    # Priority 7: "25-2200", "40-2200" — bare thickness-width at end of filename
    m = re.search(r'(?:^|[^0-9])(\d+\.?\d*)[-_](\d{3,5})(?:\.png|\.jpg|$|）|\()', f)
    if m:
        t = float(m.group(1))
        w = float(m.group(2))
        if 0.3 <= t <= 200 and 900 <= w <= 5000:
            return {'thickness': _num(m.group(1))}

    return None


# ── 钢筋 (Rebar) ────────────────────────────────────────────────────

def _rebar_extract(filename):
    """Extract (diameter, material_grade) from rebar filename."""
    f = filename
    result = {}

    # Priority 1: "12_HRB400E_202410" or "6_HRB400E"
    m = re.search(r'(\d+)[_\-](HRB\d+[A-Za-z]*)\b', f)
    if m:
        return {'diameter': _num(m.group(1)),
                'material_grade': _grade_normalize(m.group(2))}

    # Priority 1b: "14第三方报告" — digits before known keywords
    m = re.search(r'(?:^|[^0-9])(\d+)\s*(?:mm|毫米|钢筋材质单|材质单|第三方报告|吊牌)', f)
    if m:
        dia = int(m.group(1))
        if 4 <= dia <= 40:
            result['diameter'] = _num(m.group(1))
            if not result.get('diameter'):
                result['diameter'] = _num(m.group(1))
            return result

    # Priority 1c: "螺纹钢20" — 螺纹(钢) followed by digits
    m = re.search(r'螺纹(?:钢)?\s*(\d+)', f)
    if m:
        dia = int(m.group(1))
        if 4 <= dia <= 40:
            result['diameter'] = _num(m.group(1))
            return result

    # Priority 2: "HRB400E 12 202410"
    m = re.search(r'\b(HRB\d+[A-Za-z]*)\s+(\d+)\s+', f)
    if m:
        dia = int(m.group(2))
        if 4 <= dia <= 40:
            return {'diameter': _num(m.group(2)),
                    'material_grade': _grade_normalize(m.group(1))}

    # Priority 3: "12mm钢筋" or "12钢筋" or "12mm" or "钢筋12"
    m = re.search(r'(\d+)\s*mm?\s*(?:钢筋|螺纹)', f)
    if m:
        dia = int(m.group(1))
        if 4 <= dia <= 40:
            result['diameter'] = _num(m.group(1))
    if not result:
        m = re.search(r'(?:钢筋|螺纹)\s*(\d+)', f)
        if m:
            dia = int(m.group(1))
            if 4 <= dia <= 40:
                result['diameter'] = _num(m.group(1))
    if not result:
        # "25E三级" — 25mm 抗震
        m = re.search(r'(\d+)E?(?:三级|四级|级)', f)
        if m:
            dia = int(m.group(1))
            if 4 <= dia <= 40:
                result['diameter'] = _num(m.group(1))
                result['material_grade'] = 'HRB400E'

    # Priority 4: "32钢筋"
    if not result:
        m = re.search(r'(?:^|[\s_])(?:(\d+)\s*(?:钢筋|螺纹))', f)
        if m:
            dia = int(m.group(1))
            if 4 <= dia <= 40:
                result['diameter'] = _num(m.group(1))
    if not result:
        m = re.search(r'(?:^|[\s_])(\d+)\s*(?:钢筋)', f)
        if m:
            dia = int(m.group(1))
            if 4 <= dia <= 40:
                result['diameter'] = _num(m.group(1))

    return result if result else None


# ── 螺栓 (Bolt) ─────────────────────────────────────────────────────

def _bolt_extract(filename):
    """Extract (specification, grade) from bolt filename.
    Grade = strength grade like 10.9, not surface treatment."""
    f = filename
    result = {}

    # Priority 1: Specification M27×80, M42×100
    m = re.search(r'M(\d+)\s*[×xX×]\s*(\d+)', f)
    if m:
        result['specification'] = f'M{m.group(1)}×{m.group(2)}'

    # Priority 2: Strength grade — look for patterns like "10.9", "8.8"
    # Must be a standalone decimal number, not part of a date or UUID
    # Common context: "10.9淬黑", "10.9S", "螺栓10.9"
    m = re.search(r'(?:^|[^0-9])(\d+\.\d+)\s*[淬发镀S]?', f)
    if m:
        g = m.group(1)
        # Validate it looks like a bolt grade (4.6, 4.8, 5.6, 5.8, 6.8, 8.8, 9.8, 10.9, 12.9)
        valid_grades = {'4.6', '4.8', '5.6', '5.8', '6.8', '8.8', '9.8', '10.9', '12.9'}
        if g in valid_grades:
            result['grade'] = g

    return result if result else None


# ── 商混 (Commercial Concrete) ──────────────────────────────────────

def _concrete_extract(filename):
    """Extract strength_grade from concrete cert filename."""
    f = filename
    # "C30", "C35", "C40" etc.
    m = re.search(r'\b(C\d{2})\b', f)
    if m:
        return {'strength_grade': m.group(1)}
    # "C34灌浆料"
    m = re.search(r'\b(C\d{2})\D', f)
    if m:
        return {'strength_grade': m.group(1)}
    return None


# ── 铅 (Lead) ───────────────────────────────────────────────────────

def _lead_extract(filename):
    """Extract purity from lead cert filename."""
    f = filename
    # "99.99" or "99.9%"
    m = re.search(r'(\d{2}\.\d+)%?', f)
    if m:
        purity = float(m.group(1))
        if 90 <= purity <= 100:
            return {'purity': m.group(1)}
    return None


# ── 无缝管 (Seamless Pipe) ──────────────────────────────────────────

def _pipe_extract(filename):
    """Try to extract pipe params from filename."""
    f = filename
    result = {}
    # "133×20" or "180x22"
    m = re.search(r'(\d+)[×xX](\d+)\s*(?:mm|无缝|管)?', f)
    if m:
        od = int(m.group(1))
        wt = int(m.group(2))
        if 20 <= od <= 1000 and 2 <= wt <= 100 and wt < od:
            result['outer_diameter'] = str(od)
            result['wall_thickness'] = str(wt)

    m2 = re.search(r'\b(Q\d+[A-Za-z]*|20#?|45#?|40Cr|16Mn)\b', f)
    if m2:
        result['material_grade'] = _grade_normalize(m2.group(1))

    return result if result else None


# ── 日期提取 (Date extraction from filename) ───────────────────────

def _date_extract(filename):
    """Try to extract a date from filename.
    Returns (cert_date, cert_date) in YYYYMM format or None."""
    f = filename
    # "20211014" — full date
    m = re.search(r'(20\d{2})(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])', f)
    if m:
        full = m.group(0)
        return full[:6]  # YYYYMM
    # "202410" — year+month
    m = re.search(r'(20\d{2})(0[1-9]|1[0-2])(?:\D|$)', f)
    if m:
        return m.group(0)[:6]
    # "2025年2月"
    m = re.search(r'(20\d{2})\s*年\s*(\d{1,2})\s*月', f)
    if m:
        return f'{m.group(1)}{int(m.group(2)):02d}'
    return None


# ── Categorization ──────────────────────────────────────────────────

EXTRACTORS = {
    '钢板': _plate_extract,
    '钢筋': _rebar_extract,
    '螺栓': _bolt_extract,
    '商混': _concrete_extract,
    '铅': _lead_extract,
    '无缝管': _pipe_extract,
    # 圆钢, 油漆 — filenames already well-extracted, no pattern issues
}


def extract_from_filename(category, image_filename, original_filename=''):
    """Try to extract params from filenames.
    Returns (updates_dict, source_description) or (None, reason)."""
    ext = EXTRACTORS.get(category)
    if not ext:
        return None, f'no extractor for category "{category}"'

    # Check both filenames
    for src, fname in [('image_filename', image_filename),
                        ('original_filename', original_filename)]:
        if not fname:
            continue
        result = ext(fname)
        if result:
            # Also try to extract date
            return result, src

    return None, 'no pattern matched'


def extract_date_from_filename(image_filename, original_filename=''):
    """Try to extract cert_date from filename."""
    for fname in [image_filename, original_filename]:
        if not fname:
            continue
        d = _date_extract(fname)
        if d:
            return d
    return None


# ====================================================================
# Grade normalization (applied to ALL records, not just new extracts)
# ====================================================================

def normalize_params(category, params):
    """Normalize params values: grade aliases, number formatting, bolt grade cleaning."""
    result = dict(params or {})

    for key in list(result.keys()):
        val = result.get(key)

        # Clean bolt grade — remove surface treatment values
        if category == '螺栓' and key == 'grade':
            if val in ('淬黑', '发黑', '镀锌', '达克罗', '热镀锌', ''):
                del result[key]
            elif val and re.match(r'^\d+\.\d+$', str(val)):
                result[key] = str(val)
            continue

        # Number-type params
        if key in ('thickness', 'diameter', 'outer_diameter', 'wall_thickness'):
            if val:
                result[key] = _num(val)
            continue

        # Grade normalization
        if key in ('material_grade', 'grade'):
            if val:
                normalized = _grade_normalize(val)
                if normalized:
                    result[key] = normalized
            continue

    # Remove keys with empty values
    result = {k: v for k, v in result.items() if v not in (None, '')}

    return result


# ====================================================================
# Main logic
# ====================================================================

def load_all_records(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute('SELECT * FROM material_certificates ORDER BY id').fetchall()
    conn.close()
    return [dict(r) for r in rows]


def analyze(db_path):
    """Dry-run: analyze all records and report what would change."""
    records = load_all_records(db_path)

    stats = defaultdict(lambda: {'total': 0, 'had_params': 0, 'empty_params': 0,
                                  'extracted': 0, 'merged': 0, 'grade_norm': 0,
                                  'date_extracted': 0, 'conflicts': 0})
    changes = []

    for rec in records:
        cat = rec['category']
        stats[cat]['total'] += 1

        old_params = {}
        try:
            old_params = json.loads(rec['params'] or '{}')
        except (json.JSONDecodeError, TypeError):
            pass

        if old_params:
            stats[cat]['had_params'] += 1
        else:
            stats[cat]['empty_params'] += 1

        # Step A: Normalize existing params first (cleans bolt "淬黑"→removed etc.)
        old_normalized = normalize_params(cat, old_params)

        # Step B: Try filename extraction
        new_params = None
        source = ''
        extracted, source = extract_from_filename(
            cat, rec['image_filename'], rec.get('original_filename', ''))
        if extracted:
            new_params = dict(extracted)

        # Step C: Merge — start with normalized old, fill gaps from filename
        merged = dict(old_normalized)
        if new_params:
            for k, v in new_params.items():
                if k not in merged or not merged[k]:
                    merged[k] = v
                elif str(merged[k]) != str(v):
                    # Conflict: existing value differs from filename extraction
                    # Keep existing (it may have been manually set)
                    stats[cat]['conflicts'] += 1

        # Step D: Final normalization (in case merge introduced unnormalized values)
        normalized = normalize_params(cat, merged)

        # Step D: Date extraction
        new_date = None
        old_date = (rec.get('cert_date') or '').strip()
        if not old_date or len(old_date) < 6:
            new_date = extract_date_from_filename(
                rec['image_filename'], rec.get('original_filename', ''))

        # Track what changed
        old_params_str = json.dumps(old_params, ensure_ascii=False, sort_keys=True)
        new_params_str = json.dumps(normalized, ensure_ascii=False, sort_keys=True)
        params_changed = old_params_str != new_params_str
        date_changed = new_date and new_date != old_date
        grade_normalized_any = False
        for k in old_params:
            if old_params[k] != normalized.get(k):
                if k in ('material_grade', 'grade'):
                    grade_normalized_any = True
                    break

        if params_changed:
            if not old_params:
                stats[cat]['extracted'] += 1
            else:
                stats[cat]['merged'] += 1

        if grade_normalized_any:
            stats[cat]['grade_norm'] += 1

        if date_changed:
            stats[cat]['date_extracted'] += 1

        if params_changed or date_changed:
            changes.append({
                'id': rec['id'],
                'category': cat,
                'image_filename': rec['image_filename'],
                'old_params': old_params,
                'new_params': normalized,
                'old_date': old_date,
                'new_date': new_date,
                'source': source,
            })

    return stats, changes


def apply_updates(db_path, changes):
    """Apply the computed changes to the database."""
    conn = sqlite3.connect(db_path)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    updated = 0

    for ch in changes:
        new_params_str = json.dumps(ch['new_params'], ensure_ascii=False)
        updates = {'params': new_params_str}
        if ch['new_date']:
            updates['cert_date'] = ch['new_date']
        updates['updated_at'] = now

        set_clause = ', '.join(f'{k} = ?' for k in updates)
        values = list(updates.values()) + [ch['id']]
        conn.execute(
            f'UPDATE material_certificates SET {set_clause} WHERE id = ?',
            values)
        updated += 1

    conn.commit()
    conn.close()
    return updated


def print_stats(stats, changes):
    print('\n' + '=' * 70)
    print('CLEANUP DRY-RUN REPORT')
    print('=' * 70)

    grand = {'total': 0, 'had_params': 0, 'empty_params': 0,
             'extracted': 0, 'merged': 0, 'grade_norm': 0,
             'date_extracted': 0, 'conflicts': 0}

    for cat in sorted(stats.keys()):
        s = stats[cat]
        for k in grand:
            grand[k] += s[k]
        if s['extracted'] or s['merged'] or s['date_extracted']:
            print(f'\n  {cat}:')
            print(f'    Total: {s["total"]}, Had params: {s["had_params"]}, '
                  f'Empty: {s["empty_params"]}')
            if s['extracted']:
                print(f'    New extraction (from filename): {s["extracted"]}')
            if s['merged']:
                print(f'    Merged into existing: {s["merged"]}')
            if s['grade_norm']:
                print(f'    Grade normalized: {s["grade_norm"]}')
            if s['date_extracted']:
                print(f'    Date extracted: {s["date_extracted"]}')
            if s['conflicts']:
                print(f'    [!] Conflicts (new vs old): {s["conflicts"]}')

    print(f'\n  --- TOTALS ---')
    print(f'  Records: {grand["total"]}')
    print(f'  Empty params before: {grand["empty_params"]}')
    print(f'  Records changed: {len(changes)}')
    print(f'  New extractions: {grand["extracted"]}')
    print(f'  Merges: {grand["merged"]}')
    print(f'  Grade normalizations: {grand["grade_norm"]}')
    print(f'  Date extractions: {grand["date_extracted"]}')
    print(f'  Conflicts: {grand["conflicts"]}')

    # Show detail for a sample of changes
    print(f'\n{"=" * 70}')
    print('SAMPLE CHANGES (first 40):')
    print('=' * 70)
    for ch in changes[:40]:
        print(f'\n  [{ch["category"]}] id={ch["id"]} source={ch["source"]}')
        print(f'    File: {ch["image_filename"][:80]}')
        if ch['old_params'] != ch['new_params']:
            print(f'    Params: {json.dumps(ch["old_params"], ensure_ascii=False)}')
            print(f'         →  {json.dumps(ch["new_params"], ensure_ascii=False)}')
        if ch['new_date']:
            print(f'    Date:   "{ch["old_date"]}" → "{ch["new_date"]}"')


def main():
    import io
    # Use UTF-8 output on Windows
    if hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                       errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8',
                                       errors='replace')

    parser = argparse.ArgumentParser(description='Clean material cert params')
    parser.add_argument('--dry-run', action='store_true',
                        help='Preview changes without applying')
    parser.add_argument('--category', type=str,
                        help='Process only a single category')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Show all changes')
    args = parser.parse_args()

    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    db_path = os.path.join(project_root, 'image_lib', 'material_certs.db')

    if not os.path.exists(db_path):
        print(f'ERROR: Database not found: {db_path}')
        sys.exit(1)

    print(f'Database: {db_path}')
    stats, changes = analyze(db_path)

    if args.category:
        # Filter to single category
        changes = [c for c in changes if c['category'] == args.category]
        stats = {k: v for k, v in stats.items() if k == args.category}

    print_stats(stats, changes)

    if args.verbose:
        print(f'\n{"=" * 70}')
        print('ALL CHANGES:')
        print('=' * 70)
        for ch in changes:
            print(f'\n  [{ch["category"]}] id={ch["id"]}')
            print(f'    File: {ch["image_filename"]}')
            print(f'    Params: {json.dumps(ch["old_params"], ensure_ascii=False)}')
            print(f'         →  {json.dumps(ch["new_params"], ensure_ascii=False)}')
            if ch['new_date']:
                print(f'    Date:   "{ch["old_date"]}" → "{ch["new_date"]}"')

    if args.dry_run:
        print(f'\n>>> DRY RUN — no changes applied. '
              f'Run without --dry-run to apply {len(changes)} updates.')
        return

    if not changes:
        print('\nNo changes to apply.')
        return

    print(f'\n>>> Applying {len(changes)} updates...')
    updated = apply_updates(db_path, changes)
    print(f'>>> Done: {updated} records updated.')


if __name__ == '__main__':
    main()
