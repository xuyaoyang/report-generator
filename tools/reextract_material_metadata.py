#!/usr/bin/env python
"""Re-extract material certificate metadata from source images.

RapidOCR performs the full-library pass. Records whose important fields cannot
be located confidently are flagged for built-in vision review, which can be
supplied through material_metadata_builtin_review.json. Every change is logged;
back up the database before running without --dry-run.
"""

import argparse
import json
import os
import re
import shutil
import sqlite3
import time
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(ROOT, 'image_lib', 'material_certs.db')
IMAGE_ROOT = os.path.join(ROOT, 'image_lib')
TEMP_DIR = os.path.join(ROOT, '临时文件')
CACHE_PATH = os.path.join(TEMP_DIR, 'material_metadata_reextract.json')
REPORT_PATH = os.path.join(TEMP_DIR, 'material_metadata_changes_20260624.json')
REVIEW_PATH = os.path.join(TEMP_DIR, 'material_metadata_builtin_review.json')

GRADE_KEYS = {
    '钢板': ('thickness', 'material_grade'),
    '钢筋': ('diameter', 'material_grade'),
    '圆钢': ('diameter', 'material_grade'),
    '无缝管': ('outer_diameter', 'wall_thickness', 'material_grade'),
    '螺栓': ('specification', 'grade'),
    '铅': ('purity',),
    '油漆': ('brand',),
    '商混': ('strength_grade',),
}

BATCH_HEADERS = (
    (r'检验批号|CHECK\s*NO', 0),
    (r'^批号$|BATCH\s*(?:NO)?|LOT\s*(?:NO)?', 1),
    (r'钢卷号|COIL\s*(?:NO)?', 2),
    (r'炉号|熔炼号|HEAT\s*(?:NO)?', 3),
)

NOISE = {
    'NO', 'NO.', 'BATCH', 'LOT', 'HEAT', 'COIL', 'NUMBER', 'N/MM2',
    'MPA', 'MM', 'KG', 'QTY', 'GRADE', 'DIMENSION', 'SPECIFICATION',
}


def normalize_text(value):
    return (str(value or '').strip()
            .replace('：', ':').replace('×', 'x').replace('脳', 'x')
            .replace('＊', '*').replace('（', '(').replace('）', ')'))


def normalize_grade(value):
    value = normalize_text(value).upper().replace(' ', '')
    aliases = {'Q355': 'Q355B', '普板': 'Q235B'}
    return aliases.get(value, value)


def valid_batch(value):
    value = normalize_text(value).strip(',:;')
    upper = value.upper()
    if not 4 <= len(value) <= 30 or upper in NOISE:
        return False
    if re.search(r'[\u4e00-\u9fff\s]', value):
        return False
    if len(re.findall(r'\d', value)) < 3:
        return False
    if re.search(r'[|~]', value):
        return False
    if re.fullmatch(r'\d+(?:\.\d+)?-\d+(?:\.\d+)?', value):
        return False
    if re.match(r'^(?:GB|HG|JB|NB|EN|ISO)[/-]?T?\d', upper):
        return False
    if re.match(r'^(?:Q\d{3}|HRB\d+|HPB\d+|\d+#|\d+CR)', upper):
        return False
    if re.search(r'\d[*xX]\d', value) or re.search(r'[=%°]', value):
        return False
    if re.fullmatch(r'20\d{6}', value):
        return False
    return True


def make_items(results):
    items = []
    for box, text, score in results or []:
        xs = [point[0] for point in box]
        ys = [point[1] for point in box]
        items.append({
            'text': normalize_text(text), 'score': float(score),
            'x1': min(xs), 'x2': max(xs), 'y1': min(ys), 'y2': max(ys),
            'cx': (min(xs) + max(xs)) / 2,
            'cy': (min(ys) + max(ys)) / 2,
        })
    return items


def extract_manufacturer(items, image_height):
    candidates = []
    for item in items:
        text = item['text']
        if not re.search(r'(?:有限责任公司|股份有限公司|有限公司|钢铁集团|钢厂)$', text):
            continue
        if any(word in text for word in ('订货', '收货', '购货', '委托', '检测', '认证')):
            continue
        position_penalty = item['cy'] / max(image_height, 1)
        candidates.append((position_penalty, -len(text), -item['score'], text))
    if not candidates:
        return None, 0.0
    _, _, neg_score, text = min(candidates)
    return text, min(1.0, -neg_score)


def extract_grade(items, category):
    patterns_by_category = {
        '钢板': (r'Q(?:195|215|235|275|345|355|390|420|460|500|550|690)[A-Z]?',),
        '钢筋': (r'HRB\d{3,4}[A-Z]?', r'HPB\d{3,4}[A-Z]?'),
        '圆钢': (r'\d{2}CR[A-Z0-9]*', r'\d{2}SIMN', r'\d{2}#'),
        '无缝管': (r'Q\d{3}[A-Z]?', r'\d{2}CR[A-Z0-9]*', r'\d{2}SIMN', r'\d{2}#'),
    }
    patterns = patterns_by_category.get(category, ())
    if not patterns:
        return None, 0.0
    candidates = []
    for item in items:
        upper = item['text'].upper().replace(' ', '')
        for pattern in patterns:
            for match in re.finditer(pattern, upper):
                grade = normalize_grade(match.group(0))
                if grade not in ('QTY',):
                    candidates.append((item['score'], item['cy'], grade))
    if not candidates:
        return None, 0.0
    score, _, grade = min(candidates, key=lambda row: (row[1], -row[0]))
    return grade, score


def extract_batch(items, image_height):
    for pattern, priority in BATCH_HEADERS:
        headers = [item for item in items if re.search(pattern, item['text'], re.IGNORECASE)]
        if not headers:
            continue
        headers.sort(key=lambda item: item['cy'])
        for header in headers:
            aligned_headers = [item for item in headers
                               if abs(item['cx'] - header['cx']) <= 35
                               and item['cy'] <= header['cy'] + 90]
            start_y = max(item['y2'] for item in aligned_headers) + 3
            column_width = max(45, header['x2'] - header['x1'])
            candidates = []
            for item in items:
                if item['y1'] <= start_y or item['cy'] > image_height * 0.82:
                    continue
                if abs(item['cx'] - header['cx']) > column_width * 0.85:
                    continue
                value = item['text'].strip(',:;')
                if valid_batch(value):
                    candidates.append((item['cy'], -item['score'], value))
            if candidates:
                _, neg_score, value = min(candidates)
                return value, min(1.0, -neg_score), priority
    return None, 0.0, 99


def extract_date(items):
    joined = '\n'.join(item['text'] for item in items)
    patterns = (
        r'(?:签发日期|发货日期|出厂日期|制造日期|DATE\s*OF\s*ISSUE)[^\d]{0,20}(20\d{2})[-/.年]?(\d{1,2})',
        r'\b(20\d{2})(\d{2})\d{2}\b',
    )
    for pattern in patterns:
        match = re.search(pattern, joined, re.IGNORECASE)
        if match:
            month = int(match.group(2))
            if 1 <= month <= 12:
                return f'{int(match.group(1)):04d}{month:02d}'
    return None


def ocr_record(engine, row):
    from PIL import Image

    image_path = os.path.join(IMAGE_ROOT, row['category'], row['image_filename'])
    with Image.open(image_path) as image:
        width, height = image.size
    results, _ = engine(image_path)
    items = make_items(results)
    manufacturer, manufacturer_score = extract_manufacturer(items, height)
    grade, grade_score = extract_grade(items, row['category'])
    batch, batch_score, batch_priority = extract_batch(items, height)
    cert_date = extract_date(items)
    return {
        'manufacturer': manufacturer,
        'manufacturer_score': manufacturer_score,
        'material_grade': grade,
        'material_grade_score': grade_score,
        'batch_number': batch,
        'batch_score': batch_score,
        'batch_priority': batch_priority,
        'cert_date': cert_date,
        'ocr_text': '\n'.join(item['text'] for item in items),
        'image_size': [width, height],
    }


def sanitize_cached_ocr(row, ocr):
    """Apply current validation rules to OCR data saved by an older run."""
    params = json.loads(row['params'] or '{}')
    compact_text = re.sub(r'\s', '', ocr.get('ocr_text', '')).upper()

    if not valid_batch(ocr.get('batch_number')):
        ocr['batch_number'] = None
        ocr['batch_score'] = 0.0
        ocr['batch_priority'] = 99
    current_batch = normalize_text(row.get('batch_number'))
    if (not ocr.get('batch_number') and valid_batch(current_batch)
            and current_batch.upper() in compact_text):
        ocr['batch_number'] = current_batch
        ocr['batch_score'] = 0.99
        ocr['batch_priority'] = 1
    param_batch = normalize_text(params.get('heat_batch_no'))
    if (not ocr.get('batch_number') and valid_batch(param_batch)
            and param_batch.upper() in compact_text):
        ocr['batch_number'] = param_batch
        ocr['batch_score'] = 0.99
        ocr['batch_priority'] = 1
    if not ocr.get('batch_number') and row['category'] == '钢板':
        x_codes = []
        for line in ocr.get('ocr_text', '').splitlines():
            value = normalize_text(line).upper()
            if re.fullmatch(r'X[A-Z0-9]{8,}', value):
                x_codes.append(value)
        if x_codes:
            ocr['batch_number'] = x_codes[2] if len(x_codes) >= 3 else x_codes[-1]
            ocr['batch_score'] = 0.95
            ocr['batch_priority'] = 0

    manufacturer = normalize_text(ocr.get('manufacturer'))
    manufacturer = re.sub(r'^(?:生产企业|生产厂家|制造商)[:：]?\s*', '', manufacturer).strip()
    manufacturer_aliases = {
        '四川省达州钢铁集团有限责任公司': '四川省达州钢铁集团有限责任公司',
    }
    manufacturer = manufacturer_aliases.get(manufacturer, manufacturer)
    ocr['manufacturer'] = manufacturer or None
    if manufacturer and len(manufacturer) < 6:
        ocr['manufacturer'] = None
        ocr['manufacturer_score'] = 0.0
    if any(word in manufacturer for word in ('订货', '收货', '购货', '委托', '检测', '认证')):
        ocr['manufacturer'] = None
        ocr['manufacturer_score'] = 0.0
    current_manufacturer = normalize_text(params.get('manufacturer'))
    if (not ocr.get('manufacturer') and current_manufacturer
            and current_manufacturer.upper() in compact_text):
        ocr['manufacturer'] = current_manufacturer
        ocr['manufacturer_score'] = 0.99

    lines = [line for line in ocr.get('ocr_text', '').splitlines() if line.strip()]
    text_items = [{'text': line, 'score': 1.0, 'cy': index}
                  for index, line in enumerate(lines)]
    grade, score = extract_grade(text_items, row['category'])
    current_grade = normalize_grade(params.get('material_grade'))
    if current_grade and current_grade in compact_text:
        grade, score = current_grade, 0.99
    if row['category'] == '圆钢' and not grade:
        match = re.search(r'(?:牌号|STEELGRADE)(\d{2})(?!\d)', compact_text)
        if match:
            grade, score = f'{match.group(1)}#', 0.95
    ocr['material_grade'] = grade
    ocr['material_grade_score'] = score

    if row['category'] == '钢板':
        dimensions = re.findall(
            r'(?<!\d)(\d{1,3}(?:\.\d+)?)\s*[*Xx]\s*(\d{3,4})(?!\d)',
            ocr.get('ocr_text', ''), re.IGNORECASE)
        for thickness, width in dimensions:
            if 1 <= float(thickness) <= 200 and 500 <= float(width) <= 5000:
                ocr['thickness'] = thickness
                break
    if row['category'] == '钢筋' and not params.get('diameter'):
        match = re.search(r'HRB\d{3,4}[A-Z]?\s*\n\s*(\d{1,2})(?!\d)',
                          ocr.get('ocr_text', ''), re.IGNORECASE)
        if match and 5 <= int(match.group(1)) <= 50:
            ocr['diameter'] = match.group(1)

    text = ocr.get('ocr_text', '')
    if row['category'] == '油漆':
        production_batches = re.findall(r'(?<!\d)(20\d{8})(?!\d)', text)
        if production_batches:
            ocr['batch_number'] = '+'.join(dict.fromkeys(production_batches[:2]))
            ocr['batch_score'] = 0.95
        product_match = re.search(r'产品名称[:：]?\s*\n?\s*([^\n]{2,30})', text)
        if product_match:
            ocr['brand'] = product_match.group(1).strip()
    elif row['category'] == '螺栓':
        report_numbers = re.findall(r'QA\d{8,}', text, re.IGNORECASE)
        if report_numbers:
            ocr['batch_number'] = report_numbers[0].upper()
            ocr['batch_score'] = 0.95
        spec_match = re.search(r'M\d{1,3}\s*[*xX×]\s*\d{1,4}', text)
        if spec_match:
            ocr['specification'] = normalize_text(spec_match.group(0))
        grade_match = re.search(r'(?<!\d)(?:4\.8|5\.8|6\.8|8\.8|10\.9|12\.9)(?!\d)', text)
        if grade_match:
            ocr['grade'] = grade_match.group(0)
    elif row['category'] == '铅':
        batch_candidates = re.findall(r'(?<!\d)(25\d{6})(?!\d)', text)
        if batch_candidates:
            ocr['batch_number'] = batch_candidates[0]
            ocr['batch_score'] = 0.95
        purity_values = re.findall(r'99\.\d{3,5}', text)
        if purity_values:
            ocr['purity'] = max(purity_values, key=float)
    elif row['category'] == '商混':
        report_numbers = re.findall(r'\b(?:RH|YJD|E)\d{6,}\b', text, re.IGNORECASE)
        if report_numbers:
            ocr['batch_number'] = report_numbers[0].upper()
            ocr['batch_score'] = 0.95
        strength = re.search(r'(?<![A-Z0-9])C\d{2,3}(?!\d)', text, re.IGNORECASE)
        if strength:
            ocr['strength_grade'] = strength.group(0).upper()
    return ocr


def needs_vision(row, params, ocr):
    required = GRADE_KEYS.get(row['category'], ())
    if (not ocr.get('batch_number') or ocr.get('batch_score', 0) < 0.93
            or not valid_batch(ocr.get('batch_number'))):
        return True
    if not ocr.get('manufacturer') or ocr.get('manufacturer_score', 0) < 0.9:
        return True
    if ('material_grade' in required
            and (not ocr.get('material_grade') or ocr.get('material_grade_score', 0) < 0.8)):
        return True
    for key in required:
        if key != 'material_grade' and not params.get(key) and not ocr.get(key):
            return True
    return False


def safe_value(key, value):
    if value is None:
        return None
    value = normalize_text(value)
    if not value or value.lower() in ('null', 'none', 'unknown', '未识别'):
        return None
    if key == 'batch_number' and not valid_batch(value):
        return None
    if key == 'material_grade':
        return normalize_grade(value)
    if key == 'cert_date':
        digits = re.sub(r'\D', '', value)
        return digits[:6] if len(digits) >= 6 else None
    return value


def choose_updates(row, params, ocr, vision):
    source = vision or ocr
    updates = {'params': dict(params), 'category': source.get('category', row['category'])}
    batch = safe_value('batch_number', source.get('batch_number'))
    date = safe_value('cert_date', source.get('cert_date'))
    if batch:
        updates['batch_number'] = batch
        updates['params']['heat_batch_no'] = batch
    if date:
        updates['cert_date'] = date
    manufacturer = safe_value('manufacturer', source.get('manufacturer'))
    if manufacturer:
        updates['params']['manufacturer'] = manufacturer
    for key in GRADE_KEYS.get(updates['category'], ()):
        value = safe_value(key, source.get(key))
        if value:
            updates['params'][key] = value
    standard = safe_value('standard', source.get('standard'))
    if standard:
        updates['params']['standard'] = standard
    return updates


def row_snapshot(row):
    return {
        'category': row['category'],
        'batch_number': row['batch_number'],
        'cert_date': row['cert_date'],
        'params': json.loads(row['params'] or '{}'),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--limit', type=int, default=0)
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()

    from rapidocr_onnxruntime import RapidOCR

    os.makedirs(TEMP_DIR, exist_ok=True)
    cache = {'records': {}}
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, 'r', encoding='utf-8') as stream:
            cache = json.load(stream)
    reviews = {}
    if os.path.exists(REVIEW_PATH):
        with open(REVIEW_PATH, 'r', encoding='utf-8') as stream:
            reviews = json.load(stream)

    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    rows = [dict(row) for row in connection.execute(
        'SELECT * FROM material_certificates ORDER BY category, id')]
    if args.limit:
        rows = rows[:args.limit]

    engine = RapidOCR()
    vision_queue = []
    changes = []
    for index, row in enumerate(rows, 1):
        record_id = str(row['id'])
        params = json.loads(row['params'] or '{}')
        entry = cache['records'].setdefault(record_id, {})
        print(f'[{index}/{len(rows)}] id={record_id} {row["category"]}')

        if 'ocr' not in entry:
            try:
                entry['ocr'] = ocr_record(engine, row)
            except Exception as exc:
                entry['ocr_error'] = str(exc)
                print(f'  OCR error: {exc}')
                continue
            with open(CACHE_PATH, 'w', encoding='utf-8') as stream:
                json.dump(cache, stream, ensure_ascii=False, indent=2)
        ocr = sanitize_cached_ocr(row, entry['ocr'])

        vision = reviews.get(record_id)
        if needs_vision(row, params, ocr) and not vision:
            vision_queue.append({
                'id': row['id'], 'category': row['category'],
                'image_path': os.path.join(IMAGE_ROOT, row['category'], row['image_filename']),
                'current': row_snapshot(row),
                'ocr': {key: value for key, value in ocr.items() if key != 'ocr_text'},
                'required_params': list(GRADE_KEYS.get(row['category'], ())),
            })

        selected = choose_updates(row, params, ocr, vision)
        before = row_snapshot(row)
        after = {
            'category': selected.get('category', row['category']),
            'batch_number': selected.get('batch_number', row['batch_number']),
            'cert_date': selected.get('cert_date', row['cert_date']),
            'params': selected['params'],
        }
        if before == after:
            continue
        changes.append({
            'id': row['id'], 'category': row['category'],
            'image_filename': row['image_filename'],
            'source': 'builtin_vision' if vision else 'ocr',
            'before': before, 'after': after,
        })
        if not args.dry_run:
            if after['category'] != row['category']:
                source_image = os.path.join(IMAGE_ROOT, row['category'], row['image_filename'])
                target_dir = os.path.join(IMAGE_ROOT, after['category'])
                target_image = os.path.join(target_dir, row['image_filename'])
                os.makedirs(target_dir, exist_ok=True)
                if not os.path.exists(target_image):
                    shutil.copy2(source_image, target_image)
            connection.execute(
                'UPDATE material_certificates SET category=?, batch_number=?, cert_date=?, params=?, updated_at=? WHERE id=?',
                (after['category'], after['batch_number'], after['cert_date'],
                 json.dumps(after['params'], ensure_ascii=False),
                 datetime.now().strftime('%Y-%m-%d %H:%M:%S'), row['id']))
            connection.commit()

    connection.close()
    with open(REPORT_PATH, 'w', encoding='utf-8') as stream:
        json.dump({'generated_at': datetime.now().isoformat(timespec='seconds'),
                   'dry_run': args.dry_run, 'changes': changes,
                   'builtin_vision_queue': vision_queue}, stream,
                  ensure_ascii=False, indent=2)
    print(f'Done: {len(changes)} changes; built-in vision queue: {len(vision_queue)}')
    print(f'Report: {REPORT_PATH}')


if __name__ == '__main__':
    main()
