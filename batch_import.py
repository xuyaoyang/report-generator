"""
Batch import converted material certificate images into the MaterialManager DB.
Intelligently parses category-specific parameters from filenames.
Multi-page docs: first page imported, additional pages tracked in notes.
"""
import os
import sys
import re
import json
import shutil
import uuid
from datetime import datetime
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MATERIAL_DIR = os.path.join(BASE_DIR, '材质单')
IMAGE_LIB = os.path.join(BASE_DIR, 'image_lib')
sys.path.insert(0, BASE_DIR)

from core.material_manager import MaterialManager

# ── Category folder name → category key ────────────────────────────
FOLDER_TO_CATEGORY = {
    '钢板材质单': '钢板',
    '螺纹钢筋材质单': '钢筋',
    '圆钢材质单': '圆钢',
    '铅芯': '铅',
    '油漆': '油漆',
    '商混材质单': '商混',
    '缸筒材质单': '缸筒',
    '螺栓材质单': '螺栓',
}

# ── Parameter extractors per category ──────────────────────────────

def extract_steel_plate(name):
    """钢板: extract thickness + material_grade from filename."""
    params = {}
    # "10-Q355" "20-2200-235B" "16-Q235"
    m = re.search(r'(\d+\.?\d*)\s*[-x×]\s*Q(\d+)', name)
    if m:
        params['thickness'] = m.group(1)
        params['material_grade'] = 'Q' + m.group(2)
        return params
    m = re.search(r'(\d+\.?\d*)\s*[-x×]\s*(\d{3,4})\s*[-x×]\s*(\d+)\s*B', name)
    if m:
        params['thickness'] = m.group(1)
        params['material_grade'] = 'Q' + m.group(3) + 'B'
        return params
    # "Q355B 10mm" "Q235B 10MM"
    m = re.search(r'Q(\d+)\s*B?\s*[-]?\s*(\d+\.?\d*)\s*m?m', name, re.I)
    if m:
        params['material_grade'] = 'Q' + m.group(1) + ('B' if 'B' in name[m.start():m.end()] else '')
        params['thickness'] = m.group(2)
        return params
    # "15.75x1500(1)-Q355B"
    m = re.search(r'(\d+\.?\d*)\s*x\s*\d+.*?[-_]\s*Q(\d+)\s*B?', name)
    if m:
        params['thickness'] = m.group(1)
        params['material_grade'] = 'Q' + m.group(2) + ('B' if ('B' in name and '235' not in name.split('Q')[1][:3]) else '')
        return params
    # "Q235B-16-20-25-30" → first thickness listed
    m = re.search(r'Q(\d+)\s*B?\s*[-]\s*(\d+\.?\d*)', name)
    if m:
        params['material_grade'] = 'Q' + m.group(1) + ('B' if 'B' in name else '')
        params['thickness'] = m.group(2)
        return params
    # "Q235B 10mm" (space separated)
    m = re.search(r'Q(\d+)\s*B?\s+(\d+\.?\d*)', name)
    if m:
        params['material_grade'] = 'Q' + m.group(1) + ('B' if 'B' in name else '')
        params['thickness'] = m.group(2)
        return params
    # "18mm钢板" "10厚钢板"
    m = re.search(r'(\d+\.?\d*)\s*(?:mm|毫|厚)', name)
    if m:
        params['thickness'] = m.group(1)
    # "LY225 8-2200" "LY160 8-2000"
    m = re.search(r'(LY\d+)\s+(\d+\.?\d*)\s*[-]', name)
    if m:
        params['material_grade'] = m.group(1)
        params['thickness'] = m.group(2)
        return params
    # "低合金钢25 2200" "低合金钢11.5 1500" "普板20 2200" "碳板20 2200"
    m = re.search(r'(低合金钢|普板|碳板|C钢|Z\d+)\s*(\d+\.?\d*)', name)
    if m:
        params['material_grade'] = m.group(1)
        params['thickness'] = m.group(2)
        return params
    return params


def extract_rebar(name):
    """钢筋/螺纹钢筋: extract diameter + material_grade."""
    params = {}
    # "22.pdf" "32钢筋.png"
    m = re.search(r'^(\d{1,2})(?:\.pdf|钢筋)', name)
    if m:
        params['diameter'] = m.group(1)
        return params
    # "钢筋12mm" "螺纹钢20"
    m = re.search(r'(?:钢筋|螺纹)\s*(\d{1,2})', name)
    if m:
        params['diameter'] = m.group(1)
    # "昆钢三级25E" "略钢22E三级" "三级12E"
    m = re.search(r'([一二三]级)\s*(\d{1,2})\s*(E?)', name)
    if m:
        grade_map = {'一': '1', '二': '2', '三': '3'}
        level = grade_map.get(m.group(1), m.group(1))
        params['diameter'] = m.group(2)
        params['material_grade'] = f'HRB{level}00{"E" if m.group(3) else ""}'
        return params
    # "25E" "22E"
    m = re.search(r'(\d{1,2})\s*E\s*(?:三级|钢筋)?', name)
    if m:
        params['diameter'] = m.group(1)
        params['material_grade'] = f'HRB400E'
        return params
    # "6+8钢筋材质单" → first diameter
    m = re.search(r'(\d{1,2})\s*\+\s*\d{1,2}\s*(?:钢筋|毫米)', name)
    if m:
        params['diameter'] = m.group(1)
        return params
    # "20mm-" date pattern
    m = re.search(r'(\d{1,2})\s*mm\s*[-]', name)
    if m:
        params['diameter'] = m.group(1)
        return params
    return params


def extract_round_steel(name):
    """圆钢: extract diameter + material_grade."""
    params = {}
    # "40Cr---直径90mm" "40Cr-55mm"
    m = re.search(r'(40Cr|45#?|Q\d+)\s*[-]{1,3}.*?(\d{2,3})\s*(?:mm|直径)', name)
    if m:
        params['material_grade'] = m.group(1)
        params['diameter'] = m.group(2)
        return params
    # "45材质单25.22.30..." → first diameter
    m = re.search(r'(\d+)\s*[.#]?\s*(\d{2})', name)
    if m:
        params['material_grade'] = m.group(1) + '#'
        params['diameter'] = m.group(2)
        return params
    return params


def extract_lead(name):
    """铅: extract purity if discernible."""
    # Usually not in filename — leave empty
    return {}


def extract_paint(name):
    """油漆: extract brand/type from filename."""
    params = {}
    types = ['水性', '油性', '环氧', '聚氨酯', '丙烯酸', '云铁', '醇酸']
    found = [t for t in types if t in name]
    if found:
        params['brand'] = '/'.join(found)
    return params


def extract_concrete(name):
    """商混: extract strength grade."""
    params = {}
    # "C30" "C35" etc
    m = re.search(r'[Cc](\d{2})', name)
    if m:
        params['strength_grade'] = 'C' + m.group(1)
    return params


def extract_cylinder(name):
    """缸筒: extract outer_diameter x wall_thickness."""
    params = {}
    # "133x20" "180x22-松亚" "168×8连接管"
    m = re.search(r'(\d{2,3})\s*[x×]\s*(\d{1,3})', name)
    if m:
        params['outer_diameter'] = m.group(1)
        params['wall_thickness'] = m.group(2)
    # material_grade sometimes in name
    m2 = re.search(r'(Q\d+|20#|45#|40Cr|LY\d+)', name)
    if m2:
        params['material_grade'] = m2.group(1)
    return params


def extract_bolt(name):
    """螺栓: extract specification (Mxx×xx) and grade."""
    params = {}
    # "M12×80-淬黑" "M20×80-淬黑" "M27×80"
    m = re.search(r'[Mｍm]\s*(\d{2,3})\s*[×x\*]\s*(\d{2,3})', name)
    if m:
        params['specification'] = f'M{m.group(1)}×{m.group(2)}'
    # Grade: 淬黑, 10.9, 8.8, 发黑
    for g in ['淬黑', '发黑', '镀锌', '10.9', '8.8', '12.9']:
        if g in name:
            params['grade'] = g
            break
    return params


EXTRACTORS = {
    '钢板': extract_steel_plate,
    '钢筋': extract_rebar,
    '圆钢': extract_round_steel,
    '铅': extract_lead,
    '油漆': extract_paint,
    '商混': extract_concrete,
    '缸筒': extract_cylinder,
    '螺栓': extract_bolt,
}


def find_original_doc(converted_dir):
    """Given a converted dir like '10-Q355/', find the original source file."""
    parent = os.path.dirname(converted_dir)
    base = os.path.basename(converted_dir)
    # Try to match: the original file should be in the same parent dir
    for ext in ['.pdf', '.docx', '.doc', '.jpg', '.jpeg', '.png']:
        candidate = os.path.join(parent, base + ext)
        if os.path.isfile(candidate):
            return candidate
    return None


def find_best_image(converted_dir):
    """For multi-page conversions, return first page. For single images, return the file itself."""
    if os.path.isfile(converted_dir):
        return converted_dir
    if not os.path.isdir(converted_dir):
        return None
    pngs = sorted([f for f in os.listdir(converted_dir) if f.endswith('.png')])
    if pngs:
        return os.path.join(converted_dir, pngs[0])
    return None


def count_pages(converted_dir):
    if os.path.isfile(converted_dir):
        return 1
    if not os.path.isdir(converted_dir):
        return 0
    return len([f for f in os.listdir(converted_dir) if f.endswith('.png')])


def make_batch_number(filename, params):
    """Generate a batch number from filename and/or params."""
    base = os.path.splitext(filename)[0]
    # Clean up common noise
    base = re.sub(r'[\s\-_]{2,}', ' ', base).strip()
    if len(base) > 60:
        base = base[:60]
    return base


def _is_converted_output_dir(dirpath):
    """Check if a directory is a conversion output (contains page_*.png files)."""
    try:
        contents = os.listdir(dirpath)
        return any(f.startswith('page_') and f.endswith('.png') for f in contents)
    except Exception:
        return False


def _has_page_siblings(filepath):
    """Check if this PNG is one of several page_*.png files in a converted dir."""
    parent = os.path.dirname(filepath)
    if not _is_converted_output_dir(parent):
        return False
    # It's inside a converted output dir — skip it (handled via original doc)
    return True


def import_all(mgr, dry_run=False):
    """Walk MATERIAL_DIR, find all original images, import into MaterialManager."""
    stats = defaultdict(lambda: {'imported': 0, 'pages': 0, 'skipped': 0})

    # Pass 1: find all conversion output directories and their source docs
    converted_entries = []  # [(dpath, original_path, category), ...]
    standalone_images = []  # [(fpath, category), ...]

    for root, dirs, files in os.walk(MATERIAL_DIR):
        rel = os.path.relpath(root, MATERIAL_DIR)
        top_folder = rel.split(os.sep)[0]
        category = FOLDER_TO_CATEGORY.get(top_folder)
        if not category:
            continue

        for dname in list(dirs):
            dpath = os.path.join(root, dname)
            if _is_converted_output_dir(dpath):
                original = find_original_doc(dpath)
                if original:
                    converted_entries.append((dpath, original, category))
                dirs.remove(dname)  # don't walk into it

        for fname in files:
            ext = os.path.splitext(fname)[1].lower()
            if ext not in ('.jpg', '.jpeg', '.png'):
                continue
            # Skip page_*.png (conversion output, handled by converted_entries)
            if fname.startswith('page_') and fname.endswith('.png'):
                continue
            fpath = os.path.join(root, fname)
            standalone_images.append((fpath, category))

    processed = set()

    # Import converted docs (one cert per original document)
    for dpath, original, category in converted_entries:
        orig_filename = os.path.basename(original)
        pngs = sorted([f for f in os.listdir(dpath)
                       if f.startswith('page_') and f.endswith('.png')])
        if not pngs:
            continue

        img_path = os.path.join(dpath, pngs[0])
        num_pages = len(pngs)

        extractor = EXTRACTORS.get(category, lambda n: {})
        params = extractor(orig_filename)
        batch_number = make_batch_number(orig_filename, params)

        notes = f'共{num_pages}页' if num_pages > 1 else ''
        if original.endswith(('.docx', '.doc')):
            notes += ' (从Word转换)' if notes else '(从Word转换)'
        elif original.endswith('.pdf'):
            notes += ' (从PDF转换)' if notes else '(从PDF转换)'

        if dry_run:
            print(f'  [{category}] {orig_filename} → 批次={batch_number}, '
                  f'params={params}, pages={num_pages}')
            stats[category]['imported'] += 1
            stats[category]['pages'] += num_pages
        else:
            try:
                mgr.add_certificate(
                    category=category, batch_number=batch_number,
                    params=params, source_path=img_path, notes=notes,
                )
                stats[category]['imported'] += 1
                stats[category]['pages'] += num_pages
            except Exception as e:
                print(f'  FAIL [{category}] {orig_filename}: {e}')
                stats[category]['skipped'] += 1

    # Import standalone images (not from conversion)
    for fpath, category in standalone_images:
        if fpath in processed:
            continue
        processed.add(fpath)

        fname = os.path.basename(fpath)
        extractor = EXTRACTORS.get(category, lambda n: {})
        params = extractor(fname)
        batch_number = make_batch_number(fname, params)

        if dry_run:
            print(f'  [{category}] {fname} → 批次={batch_number}, '
                  f'params={params}')
            stats[category]['imported'] += 1
            stats[category]['pages'] += 1
        else:
            try:
                mgr.add_certificate(
                    category=category, batch_number=batch_number,
                    params=params, source_path=fpath,
                )
                stats[category]['imported'] += 1
                stats[category]['pages'] += 1
            except Exception as e:
                print(f'  FAIL [{category}] {fname}: {e}')
                stats[category]['skipped'] += 1

    return stats


def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--dry-run':
        dry_run = True
        print('=== DRY RUN (预览模式，不实际导入) ===\n')
    else:
        dry_run = False
        print('=== 材质单批量导入 ===\n')

    dummy_product_dir = os.path.join(BASE_DIR, 'products', 'isolation_bearing')
    mgr = MaterialManager(dummy_product_dir)

    stats = import_all(mgr, dry_run=dry_run)

    print()
    total_imported = 0
    total_pages = 0
    total_skipped = 0
    for cat in sorted(stats.keys()):
        s = stats[cat]
        print(f'  {cat}: {s["imported"]} 个文件, {s["pages"]} 页, {s["skipped"]} 失败')
        total_imported += s['imported']
        total_pages += s['pages']
        total_skipped += s['skipped']

    print(f'\n合计: {total_imported} 个材质单, {total_pages} 页图片')
    if total_skipped:
        print(f'失败: {total_skipped}')

    if dry_run:
        print('\n以上为预览。执行: python batch_import.py')


if __name__ == '__main__':
    main()
