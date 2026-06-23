#!/usr/bin/env python
"""
Batch OCR material certificate images using RapidOCR (offline).
Extracts: manufacturer, cert_date, heat_batch_no, certificate_no, standard.
Supports checkpoint/resume.

Usage:
  python tools/batch_ocr_extract.py                    # process all pending
  python tools/batch_ocr_extract.py --limit 20         # process 20
  python tools/batch_ocr_extract.py --category 螺栓    # single category
  python tools/batch_ocr_extract.py --dry-run          # preview only
"""

import json
import os
import re
import sys
import io
import sqlite3
import time
from datetime import datetime
from collections import OrderedDict

if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                   errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8',
                                   errors='replace')

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(PROJECT_ROOT, 'image_lib', 'material_certs.db')
IMAGE_LIB = os.path.join(PROJECT_ROOT, 'image_lib')
CHECKPOINT_PATH = os.path.join(PROJECT_ROOT, '临时文件', 'ocr_extract_progress.json')


# ====================================================================
# Field extractors — regex patterns applied to OCR'd text
# ====================================================================

def extract_manufacturer(text, category):
    """Extract manufacturer name from OCR text."""
    # Common patterns for manufacturer in Chinese certificates
    patterns = [
        # "生产单位：攀钢集团..."
        r'(?:生产单位|生产厂家|制造商|钢厂)[：:]\s*(.{2,30}?(?:公司|集团|钢厂|钢铁|有限|实业|股份))',
        # "攀钢集团西昌钢钒有限公司" — standalone company name
        r'(?:质量证明书|产品质量证明书|INSPECTION\s*CERTIFICATE|检验证书)\s*\n*\s*(.{2,30}?(?:公司|集团|钢铁|钢厂))',
        # "BAOWU 产品质量证明书 重庆钢铁股份有限公司"
        r'(?:证明书|CERTIFICATE)\s*\n*\s*(?:BAOWU\s*)?(.{2,30}?(?:公司|集团|钢铁|钢厂|实业|有限))',
        # At top of cert: company name before "产品质量证明书"
        r'^\s*(.{2,30}?(?:公司|集团|钢铁|钢厂))\s*\n\s*(?:产品质量证明书|INSPECTION)',
        # "攀钢集团国际经济贸易有限公司" — any 公司 near "订货单位" or "生产"
        r'([一-鿿]{4,20}(?:公司|集团|钢铁|钢厂|实业|有限))',
        # "赤峰山金银铅有限公司" — standalone from top
        r'^(.{6,30}(?:有限公司|有限责任公司))',
    ]

    for pat in patterns:
        m = re.search(pat, text, re.MULTILINE | re.IGNORECASE)
        if m:
            name = m.group(1).strip()
            # Filter false positives
            if len(name) < 4 or name in ('Date', 'CERTIFICATE', 'INSPECTION'):
                continue
            if any(w in name for w in ('标准', '技术', '检验', '证书', '合格', '试验',
                                         '检测', '质量', '证明', '合同', '收货', '订货',
                                         '车轮', '车号')):
                continue
            return name

    # Fallback: search for specific manufacturer names
    known_mfr = re.findall(r'(攀钢\S{0,15}(?:公司|集团))|'
                           r'(重庆钢铁[^\n]{0,10})|'
                           r'(陕西略阳钢铁\S{0,10})|'
                           r'(四川德胜\S{0,10})|'
                           r'(冷水江天宝\S{0,10})|'
                           r'(山西晋南\S{0,10})|'
                           r'(赤峰山金银铅\S{0,10})|'
                           r'(四川汉邦石化\S{0,10})', text)
    if known_mfr:
        for groups in known_mfr:
            for g in groups:
                if g:
                    return g

    return None


def extract_cert_date(text):
    """Extract certification/release date from OCR text.
    Key: 发货日期(DATE TO PULL) is the closest to actual shipping date.
    The date may be on the SAME line or the NEXT line after the label."""
    lines = text.split('\n')

    # Try same-line extraction first
    for pat_label, pat_date in [
        # "发货日期：2021-10-14" — same line
        (r'发货日期|DATE\s*TO\s*PULL', r'[：:]\s*(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})'),
        (r'出厂日期', r'[：:]\s*(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})'),
        (r'(?:制造日期|签发日期|检测日期)', r'[：:]\s*(\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2})'),
    ]:
        for i, line in enumerate(lines):
            if re.search(pat_label, line, re.IGNORECASE):
                # Check same line
                m = re.search(pat_date, line, re.IGNORECASE)
                if m:
                    d_clean = re.sub(r'[-/\.]', '', m.group(1))
                    if len(d_clean) >= 8:
                        return d_clean[:6]
                # Check next 1-2 lines
                for j in range(i+1, min(i+3, len(lines))):
                    next_line = lines[j].strip()
                    m2 = re.search(r'(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})', next_line)
                    if m2:
                        try:
                            return f'{int(m2.group(1)):04d}{int(m2.group(2)):02d}'
                        except ValueError:
                            pass
                    m3 = re.search(r'(\d{8,14})', next_line)  # 20211014223231
                    if m3:
                        return m3.group(1)[:6]

    # Fallback: scan for any "日期" label with a date nearby (within 2 lines)
    for i, line in enumerate(lines):
        if re.search(r'(?:日期|DATE)\s*[：:]?', line, re.IGNORECASE):
            for j in range(i, min(i+3, len(lines))):
                target = lines[j].strip()
                m = re.search(r'(\d{4})[-/\.年](\d{1,2})[-/\.月](\d{1,2})', target)
                if m:
                    try:
                        y = int(m.group(1))
                        mo = int(m.group(2))
                        if 2020 <= y <= 2026 and 1 <= mo <= 12:
                            return f'{y:04d}{mo:02d}'
                    except ValueError:
                        pass
                m2 = re.search(r'(\d{8,14})', target)
                if m2:
                    ds = m2.group(1)
                    if ds.startswith('20'):
                        return ds[:6]

    return None


def extract_heat_batch_no(text, category):
    """Extract batch/heat number from OCR text.

    重钢: PLATE NO | BATCH No | HEAT No | DIMENSION
      Data: 6P23729353200  13180659RLL3500  23106654  20.00*2000*10500
      Batch = token after PLATE NO (pattern: digits+letter+digits, 10+ chars)

    攀钢: 序号 | 熔炼号 | 钢卷号 | 检验批号
      Data: X21106932  X11015105000  X11015105000  Q235B
      检验批号 = 3rd X-code after header

    Other: SAMPLEID/LOT NO. column
    """
    tokens = text.split()
    n = len(tokens)

    if category == '钢板':
        # Method A: consecutive pair "6P2372... 13180659RLL3500"
        for ti in range(n - 1):
            a = tokens[ti]
            b = tokens[ti+1]
            if re.match(r'^\d+[A-Z]\d{6,}', a) and len(a) >= 10:
                if re.match(r'^\d{6,}[A-Z]{2,4}\d{0,}', b) and len(b) >= 10:
                    return b

        # Method B: 攀钢 — header: 序号 熔炼号 钢卷号 检验批号 ...
        # 检验批号 = 3rd X-code after '检验批号' header
        for ti, tok in enumerate(tokens):
            if re.search(r'检验批号|Check\s*No', tok, re.IGNORECASE):
                x_codes = []
                for dk in range(ti+1, min(ti+30, n)):
                    if re.match(r'^X[A-Z0-9]{8,}$', tokens[dk]):
                        x_codes.append(tokens[dk])
                    if len(x_codes) >= 3:
                        return x_codes[2]
                if x_codes:
                    return x_codes[-1]
                break

        # Method C: SAMPLEID/LOT NO. column
        for ti, tok in enumerate(tokens):
            if re.search(r'SAMPLEID|LOT\s*NO|试片编号', tok, re.IGNORECASE):
                for dj in range(ti+1, min(ti+10, n)):
                    if re.match(r'^\d{10,}$', tokens[dj]) and len(tokens[dj]) >= 10:
                        return tokens[dj]

        # Method D: standalone batch-like codes
        m = re.search(r'\b(\d{8,14}[A-Z]{2,4}\d{0,6})\b', text)
        if m:
            return m.group(1)

    # Non-钢板 fallback — use batch number pattern matching
    if category in ('无缝管', '圆钢'):
        # Batch numbers: 2203D929, 22B02225, 2302C454, 2502C549
        # Pattern: 2-4 digits + 1-2 uppercase letters + 2-5 digits
        tokens = text.split()
        for ti, tok in enumerate(tokens):
            if re.search(r'批号|Batch\s*No', tok, re.IGNORECASE):
                for dj in range(ti+1, min(ti+30, len(tokens))):
                    cand = tokens[dj]
                    if re.match(r'^\d{2,4}[A-Z]{1,2}\d{3,}$', cand) and len(cand) >= 7:
                        return cand
                break
        # Fallback: scan all tokens for the pattern
        for ti, tok in enumerate(tokens):
            if re.match(r'^\d{2,4}[A-Z]{1,2}\d{3,}$', tok) and len(tok) >= 7:
                return tok

    for pat in [
        r'(?:炉号|熔炼号|HEAT\s*No|批号|产品批号|Batch\s*No)[：:]?\s*(\S{4,30})',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if 3 <= len(val) <= 25 and not val.startswith('执行') and not val.startswith('GB'):
                return val
    return None


def extract_certificate_no(text):
    """Extract certificate number from OCR text."""
    patterns = [
        r'(?:证明书号|证书号|Certificate\s*No|质量证明书编号)[：:]\s*(\S{4,30})',
        r'(?:编号|No\.)\s*[：:]\s*(\S{4,30})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if len(val) >= 4 and not re.match(r'^\d{1,3}$', val):
                return val
    return None


def extract_standard(text):
    """Extract technical standard from OCR text."""
    patterns = [
        r'(?:执行标准|技术标准|标准)[：:]\s*(\S{3,30})',
        r'(?:GB/T\s*\d{1,6}[\.-]\d{1,4})',
        r'(?:HG/T\s*\d{1,6}[\.-]\d{1,4})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            val = m.group(1) if m.lastindex else m.group(0)
            if len(val) >= 5:
                return val.strip()
    return None


# ====================================================================
# Main extraction
# ====================================================================

def ocr_extract_one(engine, img_path, category):
    """Run OCR on one image and extract all fields."""
    text = ''
    try:
        results, _ = engine(img_path)
        text = '\n'.join(t for _, t, _ in results)
    except Exception as e:
        print(f'    OCR error: {e}')
        return None

    if not text.strip():
        return None

    updates = {}

    mfr = extract_manufacturer(text, category)
    if mfr:
        updates['manufacturer'] = mfr

    # Also extract manufacturer from the "订货单位/收货单位" if it's a steel mill
    # (some certs are from trading companies, the actual mill may be elsewhere)
    # But for now, the top-level manufacturer is most useful.

    date = extract_cert_date(text)
    if date:
        updates['cert_date'] = date

    heat = extract_heat_batch_no(text, category)
    if heat:
        updates['heat_batch_no'] = heat

    cert_no = extract_certificate_no(text)
    if cert_no:
        updates['certificate_no'] = cert_no

    std = extract_standard(text)
    if std:
        updates['standard'] = std

    return updates, text


def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'processed': {}, 'failed': {}}


def save_checkpoint(data):
    os.makedirs(os.path.dirname(CHECKPOINT_PATH), exist_ok=True)
    with open(CHECKPOINT_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_pending_records(db_path, category=None, limit=0):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    if category:
        rows = cur.execute(
            'SELECT * FROM material_certificates WHERE category=? ORDER BY id',
            (category,)).fetchall()
    else:
        rows = cur.execute(
            'SELECT * FROM material_certificates ORDER BY '
            'CASE WHEN category IN (\"钢板\",\"螺栓\",\"钢筋\",\"铅\",\"油漆\",\"无缝管\") THEN 0 ELSE 1 END, '
            'category, id'
        ).fetchall()

    conn.close()
    result = [dict(r) for r in rows]
    if limit > 0:
        result = result[:limit]
    return result


def apply_updates(db_path, updates_by_id):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    updated = 0

    for rec_id, updates in updates_by_id.items():
        row = cur.execute(
            'SELECT params, cert_date FROM material_certificates WHERE id = ?',
            (int(rec_id),)).fetchone()
        if not row:
            continue

        try:
            params = json.loads(row['params'] or '{}')
        except (json.JSONDecodeError, TypeError):
            params = {}
        cert_date = (row['cert_date'] or '').strip()

        params_changed = False
        for key in ('manufacturer', 'heat_batch_no', 'standard', 'certificate_no'):
            if key in updates and updates[key]:
                if not params.get(key):
                    params[key] = updates[key]
                    params_changed = True

        date_changed = False
        new_date = updates.get('cert_date', '')
        if new_date and (not cert_date or len(cert_date) < 6):
            cert_date = new_date
            date_changed = True

        if params_changed or date_changed:
            cur.execute(
                'UPDATE material_certificates SET params = ?, cert_date = ?, updated_at = ? WHERE id = ?',
                (json.dumps(params, ensure_ascii=False), cert_date, now, int(rec_id)))
            updated += 1

    conn.commit()
    conn.close()
    return updated


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--limit', type=int, default=0)
    parser.add_argument('--category', type=str, default=None)
    parser.add_argument('--dry-run', action='store_true')
    parser.add_argument('--sample', type=int, default=0,
                        help='Show N OCR text samples for debugging')
    args = parser.parse_args()

    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR()

    checkpoint = load_checkpoint()
    already = len(checkpoint['processed'])
    failed = len(checkpoint['failed'])
    print(f'Checkpoint: {already} processed, {failed} failed')

    records = get_pending_records(DB_PATH, category=args.category, limit=args.limit)
    pending = [r for r in records if str(r['id']) not in checkpoint['processed']]
    print(f'Pending: {len(pending)}')

    if args.dry_run:
        print(f'DRY RUN — {len(pending)} records would be processed')
        return

    if not pending:
        print('Nothing to process.')
        return

    batch_updates = {}
    total_mfr = 0
    total_date = 0
    total_heat = 0

    for i, rec in enumerate(pending):
        rec_id = str(rec['id'])
        cat = rec['category']
        img_fn = rec['image_filename']
        img_path = os.path.join(IMAGE_LIB, cat, img_fn)

        if not os.path.exists(img_path):
            checkpoint['failed'][rec_id] = 'image_missing'
            continue

        print(f'[{i+1}/{len(pending)}] id={rec_id} [{cat}] {img_fn[:50]}')

        t0 = time.time()
        result = ocr_extract_one(engine, img_path, cat)
        elapsed = time.time() - t0

        if args.sample > 0 and i < args.sample:
            updates, text = result
            print(f'    --- OCR text ---')
            print(text[:600])
            print(f'    --- end ---')

        if result is None:
            checkpoint['failed'][rec_id] = 'ocr_error'
            print(f'  [{rec_id}] FAIL: OCR error')
            continue

        updates, _ = result

        # Compare with existing
        try:
            cur_params = json.loads((rec.get('params') or '{}'))
        except (json.JSONDecodeError, TypeError):
            cur_params = {}
        cur_date = (rec.get('cert_date') or '').strip()

        new_info = []
        for key in ('manufacturer', 'heat_batch_no', 'certificate_no', 'standard'):
            if updates.get(key) and not cur_params.get(key):
                new_info.append(f'{key}={updates[key][:20]}')
        if updates.get('cert_date') and (not cur_date or len(cur_date) < 6):
            new_info.append(f'date={updates["cert_date"]}')

        if new_info:
            print(f'  [{rec_id}] OK ({elapsed:.0f}s): {", ".join(new_info)}')
            if updates.get('manufacturer'): total_mfr += 1
            if updates.get('cert_date'): total_date += 1
            if updates.get('heat_batch_no'): total_heat += 1
            batch_updates[rec_id] = updates
        else:
            print(f'  [{rec_id}] OK ({elapsed:.0f}s): no new info')

        checkpoint['processed'][rec_id] = {
            'updates': updates,
            'elapsed': elapsed,
            'category': cat,
        }

        if len(batch_updates) >= 20:
            n = apply_updates(DB_PATH, batch_updates)
            print(f'  -> Committed {n} updates')
            batch_updates = {}

        save_checkpoint(checkpoint)

    if batch_updates:
        n = apply_updates(DB_PATH, batch_updates)
        print(f'  -> Final commit: {n} updates')

    processed = len(checkpoint['processed'])
    failed = len(checkpoint['failed'])
    print(f'\nDone. Processed: {processed}, Failed: {failed}')
    print(f'Manufacturers found: {total_mfr}, Dates: {total_date}, Heat batch: {total_heat}')


if __name__ == '__main__':
    main()
