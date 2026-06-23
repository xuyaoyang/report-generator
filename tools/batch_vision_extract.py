#!/usr/bin/env python
"""
Batch extract manufacturer, heat_batch_no, cert_date from material cert images
using local Qwen2-VL vision model. Supports checkpoint/resume.

Usage:
  python tools/batch_vision_extract.py                    # process all pending
  python tools/batch_vision_extract.py --limit 10         # process 10 only
  python tools/batch_vision_extract.py --retry-failed     # retry failed ones
  python tools/batch_vision_extract.py --dry-run          # show what would run
"""

import json
import os
import re
import sys
import io
import sqlite3
import time
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                   errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8',
                                   errors='replace')

# Import the vision model
sys.path.insert(0, 'G:/model')
from vision_mcp_server import analyze_image, load_model

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(PROJECT_ROOT, 'image_lib', 'material_certs.db')
IMAGE_LIB = os.path.join(PROJECT_ROOT, 'image_lib')
CHECKPOINT_PATH = os.path.join(PROJECT_ROOT, '临时文件', 'vision_extract_progress.json')

# Fields to extract per image
EXTRACTION_PROMPT = """仅输出以下JSON格式，字段值为null时保持null：

{"manufacturer": null, "certificate_no": null, "heat_batch_no": null, "factory_date": null, "standard": null}"""


def load_checkpoint():
    if os.path.exists(CHECKPOINT_PATH):
        with open(CHECKPOINT_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'processed': {}, 'failed': {}}


def save_checkpoint(data):
    os.makedirs(os.path.dirname(CHECKPOINT_PATH), exist_ok=True)
    with open(CHECKPOINT_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_pending_records(db_path, limit=None, retry_failed=False):
    """Get records that need vision extraction.
    Priority: records with empty params first, then records missing manufacturer."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    if retry_failed:
        checkpoint = load_checkpoint()
        failed_ids = list(checkpoint['failed'].keys())
        if not failed_ids:
            print('No failed records to retry.')
            return []
        placeholders = ','.join('?' for _ in failed_ids)
        rows = cur.execute(
            f'SELECT * FROM material_certificates WHERE id IN ({placeholders})',
            [int(x) for x in failed_ids]).fetchall()
    else:
        # All records that could benefit from vision extraction
        rows = cur.execute(
            'SELECT * FROM material_certificates ORDER BY '
            # Priority: empty params first, then by category
            "CASE WHEN params = '{}' THEN 0 ELSE 1 END, category, id"
        ).fetchall()

    conn.close()
    result = [dict(r) for r in rows]
    if limit and limit > 0:
        result = result[:limit]
    return result


def parse_json_response(text):
    """Try to extract JSON from model response."""
    text = text.strip()
    # Remove markdown code fences
    text = re.sub(r'^```(?:json)?\s*', '', text)
    text = re.sub(r'\s*```$', '', text)
    text = text.strip()

    # Try direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try to find JSON object
    m = re.search(r'\{[^{}]*\}', text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass

    return None


def extract_date_value(parsed, key='factory_date'):
    """Extract and normalize date string."""
    val = parsed.get(key, '')
    if not val or val in ('null', 'None', ''):
        return None
    val = str(val).strip()
    # Try common date formats
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d',
                '%Y-%m', '%Y/%m', '%Y.%m', '%Y%m']:
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime('%Y%m')
        except ValueError:
            continue
    # Try to extract year+month from various formats
    m = re.search(r'(20\d{2})[年/\-.](\d{1,2})', val)
    if m:
        return f'{m.group(1)}{int(m.group(2)):02d}'
    return val[:6] if len(val) >= 6 else val


def process_record(rec, checkpoint):
    """Run vision extraction on one record. Returns updates dict or None."""
    rec_id = str(rec['id'])
    category = rec['category']
    img_filename = rec['image_filename']

    # Skip already processed
    if rec_id in checkpoint['processed']:
        return None

    img_path = os.path.join(IMAGE_LIB, category, img_filename)
    if not os.path.exists(img_path):
        print(f'  [{rec_id}] SKIP: image not found: {img_path}')
        checkpoint['failed'][rec_id] = 'image_not_found'
        return None

    # Read current params
    try:
        current_params = json.loads(rec['params'] or '{}')
    except (json.JSONDecodeError, TypeError):
        current_params = {}

    # Check if we really need to extract — skip if all fields present
    has_mfr = bool(current_params.get('manufacturer', ''))
    has_heat = bool(current_params.get('heat_batch_no', ''))
    has_date = bool((rec.get('cert_date') or '').strip())
    # We always try to get manufacturer since it's 0/545

    try:
        t0 = time.time()
        raw = analyze_image(img_path, EXTRACTION_PROMPT)
        elapsed = time.time() - t0

        # Show a snippet of the raw response for debugging
        print(f'    Raw: {raw[:200].strip()}...' if len(raw) > 200 else f'    Raw: {raw.strip()}')

        parsed = parse_json_response(raw)
        if not parsed:
            print(f'  [{rec_id}] FAIL: could not parse JSON. Raw: {raw[:100]}...')
            checkpoint['failed'][rec_id] = 'json_parse_error'
            return None

        updates = {}
        # Manufacturer — handle both plain string and dict form
        mfr = parsed.get('manufacturer', '')
        if isinstance(mfr, dict):
            mfr = mfr.get('name', '') or mfr.get('厂家', '') or ''
        if mfr and str(mfr).strip() not in ('null', 'None', '', None):
            updates['manufacturer'] = str(mfr).strip()

        # Heat/batch no
        heat = parsed.get('heat_batch_no', '')
        if heat and heat not in ('null', 'None', '', None):
            updates['heat_batch_no'] = str(heat).strip()

        # Factory date
        date_val = extract_date_value(parsed)
        if date_val:
            updates['cert_date'] = date_val

        # Standard
        std = parsed.get('standard', '')
        if std and std not in ('null', 'None', '', None):
            updates['standard'] = str(std).strip()

        # Certificate no
        cert_no = parsed.get('certificate_no', '')
        if cert_no and cert_no not in ('null', 'None', '', None):
            updates['certificate_no'] = str(cert_no).strip()

        status = 'extracted'
        info = []
        if updates.get('manufacturer'):
            info.append(f'mfr={updates["manufacturer"][:20]}')
        if updates.get('heat_batch_no'):
            info.append(f'heat={updates["heat_batch_no"][:15]}')
        if updates.get('cert_date'):
            info.append(f'date={updates["cert_date"]}')

        checkpoint['processed'][rec_id] = {
            'status': status,
            'updates': updates,
            'elapsed': round(elapsed, 1),
            'category': category,
        }
        print(f'  [{rec_id}] OK ({elapsed:.0f}s): {", ".join(info) if info else "no new info"}')
        return updates

    except Exception as e:
        print(f'  [{rec_id}] ERROR: {e}')
        checkpoint['failed'][rec_id] = str(e)[:200]
        return None


def apply_updates(db_path, updates_by_id):
    """Apply extracted updates to database."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    updated = 0

    for rec_id, updates in updates_by_id.items():
        # Get current state
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

        # Merge new fields into params
        params_changed = False
        for key in ('manufacturer', 'heat_batch_no', 'standard', 'certificate_no'):
            if key in updates and updates[key]:
                if not params.get(key):
                    params[key] = updates[key]
                    params_changed = True

        # Update cert_date
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
    parser.add_argument('--limit', type=int, default=0, help='Process N records')
    parser.add_argument('--retry-failed', action='store_true')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    print(f'Database: {DB_PATH}')
    print(f'Checkpoint: {CHECKPOINT_PATH}')

    checkpoint = load_checkpoint()
    if args.retry_failed:
        checkpoint['failed'] = {}
        save_checkpoint(checkpoint)

    already = len(checkpoint['processed'])
    failed_count = len(checkpoint['failed'])
    print(f'Already processed: {already}, Failed: {failed_count}')

    records = get_pending_records(DB_PATH, limit=args.limit,
                                   retry_failed=args.retry_failed)

    # Filter out already processed
    pending = [r for r in records if str(r['id']) not in checkpoint['processed']]
    print(f'Pending: {len(pending)}')

    if args.dry_run:
        print('DRY RUN — showing first 20 pending:')
        for r in pending[:20]:
            cat = r['category']
            img = r['image_filename'][:60]
            has_params = r['params'] != '{}'
            has_date = bool((r.get('cert_date') or '').strip())
            print(f'  [{r["id"]}] {cat}: {img} (params={has_params}, date={has_date})')
        return

    if not pending:
        print('Nothing to process.')
        return

    # Load model once
    print('Loading vision model...')
    load_model()
    print('Model loaded. Starting extraction...\n')

    batch_updates = {}
    for i, rec in enumerate(pending):
        category = rec['category']
        img = rec['image_filename'][:60]
        has_params = bool(json.loads(rec['params'] or '{}'))
        print(f'[{i+1}/{len(pending)}] id={rec["id"]} [{category}] {img}')

        updates = process_record(rec, checkpoint)
        if updates:
            # Apply immediately with max frequency
            batch_updates[str(rec['id'])] = updates
            # Batch commit every 10 or at end
            if len(batch_updates) >= 10:
                n = apply_updates(DB_PATH, batch_updates)
                print(f'  -> Committed {n} updates to DB')
                batch_updates = {}

        # Save checkpoint after each record
        save_checkpoint(checkpoint)

    # Final commit
    if batch_updates:
        n = apply_updates(DB_PATH, batch_updates)
        print(f'  -> Final commit: {n} updates')

    # Report
    processed = len(checkpoint['processed'])
    failed = len(checkpoint['failed'])
    print(f'\nDone. Processed: {processed}, Failed: {failed}')
    print(f'Checkpoint saved to: {CHECKPOINT_PATH}')


if __name__ == '__main__':
    main()
