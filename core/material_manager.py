"""
Material certificate image manager: storage, CRUD, matching, and
Word report image insertion.
"""
import os
import json
import sqlite3
import shutil
import uuid
import re
from datetime import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class MaterialManager:
    """Manage material certificate images for a product type."""

    def __init__(self, product_dir, product_type=None):
        self.project_root = os.path.dirname(os.path.dirname(
            os.path.abspath(__file__)))
        self.product_dir = product_dir
        self.product_type = product_type
        self.image_lib_dir = os.path.join(self.project_root, 'image_lib')

        config_path = os.path.join(self.project_root, 'config',
                                    'material_categories.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            self.category_config = json.load(f)
        self.categories = self.category_config['categories']
        self.expiry_months = self.category_config.get('expiry_months', 24)
        self._config_path = config_path

        self.db_path = os.path.join(self.image_lib_dir, 'material_certs.db')
        self._init_db()

    def add_category(self, key, label):
        """Add a new category to config and persist to JSON file."""
        max_order = max((c.get('order', 0) for c in self.categories.values()),
                        default=0)
        self.categories[key] = {
            'label': label,
            'order': max_order + 1,
            'group_by': [],
            'params': [],
        }
        self.category_config['categories'] = self.categories
        with open(self._config_path, 'w', encoding='utf-8') as f:
            json.dump(self.category_config, f, ensure_ascii=False, indent=2)
        category_dir = os.path.join(self.image_lib_dir, key)
        os.makedirs(category_dir, exist_ok=True)

    # =================================================================
    # Database
    # =================================================================

    def _init_db(self):
        os.makedirs(self.image_lib_dir, exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        conn.execute('''CREATE TABLE IF NOT EXISTS material_certificates (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            category        TEXT NOT NULL,
            batch_number    TEXT NOT NULL,
            params          TEXT NOT NULL DEFAULT '{}',
            image_filename  TEXT NOT NULL,
            original_filename TEXT NOT NULL,
            file_size       INTEGER DEFAULT 0,
            notes           TEXT DEFAULT '',
            is_default      INTEGER DEFAULT 0,
            cert_date       TEXT DEFAULT '',
            is_expired      INTEGER DEFAULT 0,
            created_at      TEXT NOT NULL,
            updated_at      TEXT NOT NULL
        )''')
        for col, defn in [('is_default', 'INTEGER DEFAULT 0'),
                           ('cert_date', "TEXT DEFAULT ''"),
                           ('is_expired', 'INTEGER DEFAULT 0'),
                           ('product_type', "TEXT DEFAULT ''"),
                           ('rotation', 'INTEGER DEFAULT 0')]:
            try:
                conn.execute(
                    f'ALTER TABLE material_certificates ADD COLUMN {col} {defn}'
                )
            except sqlite3.OperationalError:
                pass
        conn.commit()
        conn.close()

    def _connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _row_to_dict(self, row):
        d = dict(row)
        d['params'] = json.loads(d['params']) if d['params'] else {}
        d['rotation'] = int(d.get('rotation') or 0) % 360
        return d

    # =================================================================
    # CRUD
    # =================================================================

    def batch_auto_crop_existing(self):
        """Run auto-crop on all existing cert images in the library.
        Returns (total, cropped) count tuple.
        """
        conn = self._connect()
        rows = conn.execute(
            'SELECT id, category, image_filename FROM material_certificates'
        ).fetchall()
        conn.close()

        total = 0
        cropped = 0
        for row in rows:
            if not row['image_filename']:
                continue
            path = os.path.join(self.image_lib_dir, row['category'],
                               row['image_filename'])
            if not os.path.exists(path):
                continue
            total += 1
            if self.auto_crop_image(path):
                cropped += 1
                print(f'  Cropped: {row["category"]}/{row["image_filename"]}')
        print(f'Auto-crop done: {cropped}/{total} images cropped.')
        return total, cropped

    @staticmethod
    def auto_crop_image(image_path, margin_threshold=240, min_crop_ratio=0.05):
        """Crop white margins from a scanned image. Uses content-density
        projection to ignore sparse noise (e.g. staple marks, scanner edges).
        Only applies if the cropped area is significantly smaller than the
        original (at least min_crop_ratio reduction in width or height).

        Returns True if the image was cropped, False otherwise.
        """
        try:
            from PIL import Image
            img = Image.open(image_path)
            if img.mode not in ('RGB', 'L'):
                img = img.convert('RGB')
            gray = img.convert('L')
            w, h = img.size

            # Row/column projections: fraction of content pixels per line
            content = gray.point(lambda p: 0 if p >= margin_threshold else 1)
            pix = list(content.getdata())
            rows = [pix[i * w:(i + 1) * w] for i in range(h)]
            row_density = [sum(r) / w for r in rows]
            col_density = [
                sum(pix[x] for x in range(x, w * h, w)) / h
                for x in range(w)
            ]

            DENSITY = 0.02  # 2% of pixels in a line must be content

            # Find content bounds from density projections
            y1 = next((i for i, d in enumerate(row_density) if d > DENSITY), 0)
            y2 = next((i for i, d in enumerate(reversed(row_density)) if d > DENSITY), 0)
            y2 = h - y2
            x1 = next((i for i, d in enumerate(col_density) if d > DENSITY), 0)
            x2 = next((i for i, d in enumerate(reversed(col_density)) if d > DENSITY), 0)
            x2 = w - x2

            if x1 >= x2 or y1 >= y2:
                return False

            # Only crop if significant reduction
            w_ratio = 1.0 - (x2 - x1) / w
            h_ratio = 1.0 - (y2 - y1) / h
            if w_ratio < min_crop_ratio and h_ratio < min_crop_ratio:
                return False

            # Add padding
            pad = 15
            x1 = max(0, x1 - pad)
            y1 = max(0, y1 - pad)
            x2 = min(w, x2 + pad)
            y2 = min(h, y2 + pad)
            cropped = img.crop((x1, y1, x2, y2))
            cropped.save(image_path, quality=95)
            return True
        except Exception:
            return False

    def add_certificate(self, category, batch_number, params=None,
                        source_path=None, notes='', cert_date=''):
        if params is None:
            params = {}
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        image_filename = ''
        file_size = 0
        original_filename = ''
        if source_path and os.path.exists(source_path):
            original_filename = os.path.basename(source_path)
            ext = os.path.splitext(source_path)[1].lower()
            uid = uuid.uuid4().hex[:8]
            safe_batch = batch_number.replace('/', '_').replace('\\', '_')
            image_filename = f'{uid}_{safe_batch}{ext}'
            category_dir = os.path.join(self.image_lib_dir, category)
            os.makedirs(category_dir, exist_ok=True)
            dest = os.path.join(category_dir, image_filename)
            shutil.copy2(source_path, dest)
            if ext in ('.jpg', '.jpeg', '.png'):
                self.auto_crop_image(dest)
            file_size = os.path.getsize(dest)

        is_expired = self._calc_expired(cert_date, False)

        conn = self._connect()
        conn.execute('''INSERT INTO material_certificates
            (category, batch_number, params, image_filename,
             original_filename, file_size, notes, cert_date, is_expired,
             created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
            (category, batch_number, json.dumps(params, ensure_ascii=False),
             image_filename, original_filename, file_size, notes,
             cert_date, is_expired, now, now))
        conn.commit()
        cert_id = conn.execute('SELECT last_insert_rowid()').fetchone()[0]
        conn.close()
        return cert_id

    @staticmethod
    def convert_pdf_to_images(pdf_path, output_dir):
        """Convert each page of a PDF to a JPG image.
        Returns a list of (page_number, image_path) tuples.
        """
        import fitz
        doc = fitz.open(pdf_path)
        result = []
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        os.makedirs(output_dir, exist_ok=True)
        for i in range(len(doc)):
            page = doc[i]
            pix = page.get_pixmap(dpi=200)
            img_path = os.path.join(output_dir, f'{base_name}_p{i + 1}.jpg')
            pix.save(img_path)
            result.append((i + 1, img_path))
        doc.close()
        return result

    def update_certificate(self, cert_id, **kwargs):
        allowed = {
            'category', 'batch_number', 'params', 'notes', 'cert_date', 'is_expired',
            'rotation',
        }
        updates = {k: v for k, v in kwargs.items() if k in allowed}
        if not updates:
            return False

        old = self.get_certificate(cert_id)
        if not old:
            return False

        new_batch = updates.get('batch_number', old['batch_number'])
        new_params = updates.get('params', old['params'])
        new_category = updates.get('category', old['category'])
        if new_category not in self.categories:
            updates.pop('category', None)
            new_category = old['category']

        # Auto-calculate is_expired from cert_date if date is provided
        new_date = updates.get('cert_date', old.get('cert_date', ''))
        if new_date:
            updates['is_expired'] = self._calc_expired(new_date, False)
            updates['cert_date'] = new_date
        elif 'cert_date' not in updates and 'is_expired' in updates:
            pass  # explicit is_expired without date change — keep as-is
        elif 'cert_date' in updates and not new_date:
            # Date cleared — keep explicit is_expired or default to 0
            if 'is_expired' not in updates:
                updates['is_expired'] = old.get('is_expired', 0)

        # Rename/move file when filename-driving fields or category changed.
        name_changed = any(k in updates for k in (
            'category', 'batch_number', 'params', 'cert_date'))
        if name_changed:
            use_date = updates.get('cert_date', old.get('cert_date', ''))
            new_fn = self._build_image_filename(
                new_category, new_batch, new_params,
                old.get('image_filename', ''), use_date)
            if (new_fn != old.get('image_filename', '')
                    or new_category != old['category']):
                if old.get('image_filename'):
                    old_path = os.path.join(self.image_lib_dir,
                                           old['category'],
                                           old['image_filename'])
                    new_path = os.path.join(self.image_lib_dir,
                                           new_category, new_fn)
                    os.makedirs(os.path.dirname(new_path), exist_ok=True)
                    if os.path.exists(old_path):
                        try:
                            os.rename(old_path, new_path)
                        except OSError:
                            pass
                updates['image_filename'] = new_fn

        if 'params' in updates and isinstance(updates['params'], dict):
            updates['params'] = json.dumps(updates['params'], ensure_ascii=False)
        if 'rotation' in updates:
            try:
                updates['rotation'] = int(updates['rotation']) % 360
            except (TypeError, ValueError):
                updates['rotation'] = 0
        updates['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        set_clause = ', '.join(f'{k} = ?' for k in updates)
        values = list(updates.values()) + [cert_id]
        conn = self._connect()
        conn.execute(
            f'UPDATE material_certificates SET {set_clause} WHERE id = ?',
            values)
        conn.commit()
        rows = conn.total_changes
        conn.close()
        return rows > 0

    def _build_image_filename(self, category, batch_number, params,
                              old_filename, cert_date=''):
        """Build descriptive filename: {batch}_{param1}_{param2}_{date}.ext
        Resolves conflicts by appending （1）, （2）, etc."""
        cat_config = self.categories.get(category, {})
        param_parts = []
        for p_def in cat_config.get('params', []):
            val = params.get(p_def['key'], '')
            if val:
                param_parts.append(str(val))

        base = batch_number
        if param_parts:
            base += '_' + '_'.join(param_parts)
        if cert_date:
            base += '_' + cert_date
        base = base.replace('/', '_').replace('\\', '_').replace(':', '_')

        old_ext = os.path.splitext(old_filename)[1] if old_filename else '.jpg'
        old_base = os.path.splitext(old_filename)[0] if old_filename else ''

        if old_base == base:
            return f'{base}{old_ext}'

        category_dir = os.path.join(self.image_lib_dir, category)
        candidate = f'{base}{old_ext}'
        full = os.path.join(category_dir, candidate)
        if not os.path.exists(full) or candidate == old_filename:
            return candidate

        counter = 1
        while True:
            candidate = f'{base}（{counter}）{old_ext}'
            full = os.path.join(category_dir, candidate)
            if not os.path.exists(full) or candidate == old_filename:
                return candidate
            counter += 1

    def delete_certificate(self, cert_id):
        conn = self._connect()
        row = conn.execute(
            'SELECT category, image_filename FROM material_certificates '
            'WHERE id = ?', (cert_id,)).fetchone()
        if not row:
            conn.close()
            return False
        if row['image_filename']:
            filepath = os.path.join(self.image_lib_dir, row['category'],
                                     row['image_filename'])
            if os.path.exists(filepath):
                os.remove(filepath)
        conn.execute('DELETE FROM material_certificates WHERE id = ?',
                     (cert_id,))
        conn.commit()
        conn.close()
        return True

    def get_certificate(self, cert_id):
        conn = self._connect()
        row = conn.execute(
            'SELECT * FROM material_certificates WHERE id = ?',
            (cert_id,)).fetchone()
        conn.close()
        return self._row_to_dict(row) if row else None

    def toggle_default(self, cert_id, product_type=None):
        """Toggle default status for a certificate in the given product.

        Stores the product type in the product_type field as a comma-separated
        list (e.g. 'isolation_bearing,viscous_damper'). is_default is kept
        in sync: 1 when the list is non-empty, 0 when empty."""
        conn = self._connect()
        row = conn.execute(
            'SELECT is_default, product_type FROM material_certificates WHERE id = ?',
            (cert_id,)).fetchone()
        if not row:
            conn.close()
            return None

        pts = (row['product_type'] or '').strip()
        existing = [p.strip() for p in pts.split(',') if p.strip()] if pts else []
        pt_key = product_type or self.product_type or ''

        if pt_key in existing:
            existing.remove(pt_key)
        else:
            existing.append(pt_key)

        new_pts = ','.join(existing)
        new_is_default = 1 if existing else 0
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn.execute(
            'UPDATE material_certificates SET is_default = ?, product_type = ?, updated_at = ? WHERE id = ?',
            (new_is_default, new_pts, now, cert_id))
        conn.commit()
        conn.close()
        return new_is_default

    def get_certificates_by_category(self, category, exclude_expired=False):
        conn = self._connect()
        rows = conn.execute(
            '''SELECT * FROM material_certificates
               WHERE category = ? ORDER BY batch_number, created_at DESC''',
            (category,)).fetchall()
        conn.close()
        certs = [self._row_to_dict(r) for r in rows]
        if exclude_expired:
            certs = [c for c in certs
                     if not self._is_cert_expired(c.get('cert_date', ''),
                                                   c.get('is_expired', 0))]
        return certs

    def get_all_categories_with_counts(self, exclude_expired=False):
        conn = self._connect()
        result = []
        for cat_key in sorted(self.categories.keys(),
                              key=lambda k: self.categories[k].get('order', 99)):
            rows = conn.execute(
                'SELECT cert_date, is_expired FROM material_certificates WHERE category = ?',
                (cat_key,)).fetchall()
            if exclude_expired:
                count = sum(1 for r in rows
                            if not self._is_cert_expired(r['cert_date'] or '',
                                                          r['is_expired']))
            else:
                count = len(rows)
            result.append({
                'key': cat_key,
                'label': self.categories[cat_key]['label'],
                'order': self.categories[cat_key].get('order', 99),
                'count': count,
            })
        conn.close()
        return result

    def get_category_count(self, category, exclude_expired=False):
        """Return the count of certs in a single category."""
        conn = self._connect()
        rows = conn.execute(
            'SELECT cert_date, is_expired FROM material_certificates WHERE category = ?',
            (category,)).fetchall()
        if exclude_expired:
            count = sum(1 for r in rows
                        if not self._is_cert_expired(r['cert_date'] or '',
                                                      r['is_expired']))
        else:
            count = len(rows)
        conn.close()
        return count

    def _calc_expired(self, cert_date, fallback):
        """Return 1 if cert_date (YYYYMM) is more than expiry_months ago."""
        if not cert_date or len(cert_date) < 6:
            return fallback
        try:
            y = int(cert_date[:4])
            m = int(cert_date[4:6])
        except (ValueError, IndexError):
            return fallback
        now = datetime.now()
        cert_months = y * 12 + m
        now_months = now.year * 12 + now.month
        return 1 if now_months - cert_months > self.expiry_months else 0

    def get_expired_certificates(self):
        """Return all expired certs (real-time calc from cert_date + override)."""
        conn = self._connect()
        rows = conn.execute(
            'SELECT * FROM material_certificates ORDER BY category, batch_number'
        ).fetchall()
        conn.close()
        certs = [self._row_to_dict(r) for r in rows]
        return [c for c in certs
                if self._is_cert_expired(c.get('cert_date', ''),
                                          c.get('is_expired', 0))]

    def _is_cert_expired(self, cert_date, is_expired):
        """Real-time expired check: use date if present, else manual flag."""
        if cert_date:
            return bool(self._calc_expired(cert_date, 0))
        return bool(is_expired)

    def get_expired_count(self):
        """Return real-time count of expired certificates."""
        conn = self._connect()
        rows = conn.execute(
            'SELECT cert_date, is_expired FROM material_certificates'
        ).fetchall()
        conn.close()
        return sum(1 for r in rows
                   if self._is_cert_expired(r['cert_date'] or '',
                                             r['is_expired']))

    def get_certs_grouped_by_subcategory(self, category, exclude_expired=False):
        """Group certs within a category by the configured group_by params.

        Returns a list of dicts:
          [{'sub_label': 'Q345B·3mm', 'param_values': {...}, 'count': N}, ...]
        Returns None if the category has no group_by config.
        """
        cat_def = self.categories.get(category, {})
        group_by = cat_def.get('group_by', [])
        if not group_by:
            return None

        certs = self.get_certificates_by_category(category,
                                                   exclude_expired=exclude_expired)
        cat_params = cat_def.get('params', [])

        # Build short-label lookup for group_by keys
        param_short = {}
        for p in cat_params:
            param_short[p['key']] = p.get('short', p['label'])

        # Group by param value tuples
        groups = {}  # (val1, val2, ...) -> [cert, ...]
        unclassified = []

        for cert in certs:
            params = cert.get('params', {})
            key_vals = tuple(
                str(params.get(k, '')).strip() for k in group_by
            )
            if all(kv for kv in key_vals):
                if key_vals not in groups:
                    groups[key_vals] = []
                groups[key_vals].append(cert)
            else:
                unclassified.append(cert)

        result = []
        for key_vals in sorted(groups.keys()):
            label_parts = []
            for k, v in zip(group_by, key_vals):
                label_parts.append(f'{param_short.get(k, k)}{v}')
            sub_label = ' · '.join(label_parts)
            param_values = {k: v for k, v in zip(group_by, key_vals)}
            result.append({
                'sub_label': sub_label,
                'param_values': param_values,
                'count': len(groups[key_vals]),
            })

        if unclassified:
            result.append({
                'sub_label': '未分类',
                'param_values': {},
                'count': len(unclassified),
            })

        return result

    def get_default_certificates(self):
        """Return certificates marked as default for current product type."""
        conn = self._connect()
        if self.product_type:
            rows = conn.execute(
                '''SELECT * FROM material_certificates
                   WHERE is_default = 1 AND product_type LIKE ?
                   ORDER BY category, batch_number''',
                (f'%{self.product_type}%',)).fetchall()
        else:
            rows = conn.execute(
                '''SELECT * FROM material_certificates
                   WHERE is_default = 1
                   ORDER BY category, batch_number''').fetchall()
        conn.close()
        return [self._row_to_dict(r) for r in rows]

    def is_default_for(self, cert):
        """Check if a certificate is default for the current product type."""
        if not cert.get('is_default'):
            return False
        if not self.product_type:
            return bool(cert.get('is_default'))
        pts = (cert.get('product_type', '') or '').strip()
        return self.product_type in [p.strip() for p in pts.split(',') if p.strip()]

    def get_image_path(self, cert):
        if not cert.get('image_filename'):
            return None
        return os.path.join(self.image_lib_dir, cert['category'],
                            cert['image_filename'])

    def get_report_image_path(self, cert):
        """Return the image path to use in reports, applying saved rotation.
        The original image is left untouched; rotated copies are cached under
        image_lib/_rotated."""
        image_path = self.get_image_path(cert)
        if not image_path or not os.path.exists(image_path):
            return image_path

        rotation = int(cert.get('rotation') or 0) % 360
        if rotation == 0:
            return image_path

        cache_dir = os.path.join(self.image_lib_dir, '_rotated')
        os.makedirs(cache_dir, exist_ok=True)
        stem = os.path.splitext(os.path.basename(image_path))[0]
        stem = re.sub(r'[^A-Za-z0-9_.-]+', '_', stem)
        mtime = int(os.path.getmtime(image_path))
        cache_name = f'{cert.get("id", "cert")}_{rotation}_{mtime}_{stem}.png'
        cache_path = os.path.join(cache_dir, cache_name)
        if os.path.exists(cache_path):
            return cache_path

        try:
            from PIL import Image
            with Image.open(image_path) as img:
                rotated = img.rotate(-rotation, expand=True)
                if rotated.mode not in ('RGB', 'RGBA'):
                    rotated = rotated.convert('RGB')
                rotated.save(cache_path)
            return cache_path
        except Exception:
            return image_path

    # =================================================================
    # Matching
    # =================================================================

    def get_certs_by_ids(self, cert_ids):
        """Get certificates by their IDs, grouped by category in display order.

        Returns: [{'category': '钢板', 'certs': [cert, ...]}, ...]
        """
        if not cert_ids:
            return []

        conn = self._connect()
        placeholders = ','.join('?' for _ in cert_ids)
        rows = conn.execute(
            f'SELECT * FROM material_certificates WHERE id IN ({placeholders})',
            cert_ids).fetchall()
        conn.close()
        certs = [self._row_to_dict(r) for r in rows]

        # Group by category
        grouped = {}
        for cert in certs:
            cat = cert['category']
            if cat not in grouped:
                grouped[cat] = []
            grouped[cat].append(cert)

        # Sort within each category by batch_number
        for cat in grouped:
            grouped[cat].sort(key=lambda c: c['batch_number'])

        # Sort categories by config order
        result = []
        for cat_key in sorted(self.categories.keys(),
                              key=lambda k: self.categories[k].get('order', 99)):
            if cat_key in grouped:
                result.append({'category': cat_key,
                               'certs': grouped[cat_key]})

        return result

    def match_certificates_for_generation(self, excel_data):
        """Match material certificates based on product material requirements
        defined in the Excel '产品材料需求' sheet.

        Certificates marked as default (is_default=1) are always included
        regardless of Excel requirements.

        Returns: [{'category': '钢板', 'certs': [cert, ...]}, ...]
        """
        material_reqs = excel_data.get('material_requirements', [])
        product_models = {p.get('产品型号', '') for p
                          in excel_data.get('product_list', [])}

        # Parse requirements from Excel
        requirements = []
        for req in material_reqs:
            model = req.get('产品型号', '').strip()
            if model not in product_models:
                continue
            category = req.get('材料类别', '').strip()
            if not category or category not in self.categories:
                continue
            spec_str = req.get('规格参数', '').strip()
            spec = self._parse_spec(spec_str, category)
            requirements.append({'model': model, 'category': category,
                                 'spec': spec})

        # Load all certs from DB
        conn = self._connect()
        rows = conn.execute(
            'SELECT * FROM material_certificates ORDER BY category, batch_number'
        ).fetchall()
        conn.close()
        all_certs = [self._row_to_dict(r) for r in rows]

        matched = {}  # (category, batch_number) -> cert

        # Always include default certs for this product type
        for cert in all_certs:
            if cert.get('is_default'):
                pt = cert.get('product_type', '') or ''
                if pt and self.product_type and self.product_type not in pt.split(','):
                    continue
                key = (cert['category'], cert['batch_number'])
                matched[key] = cert

        # Match Excel requirements
        for req in requirements:
            cat = req['category']
            spec = req['spec']
            for cert in all_certs:
                if cert['category'] != cat:
                    continue
                if not self._params_match(cert.get('params', {}), spec, cat):
                    continue
                key = (cat, cert['batch_number'])
                if key not in matched:
                    matched[key] = cert

        if not matched:
            print('No material certificates matched.')
            return []

        # Group by category, sorted by category order then batch_number
        result = []
        for cat_key in sorted(self.categories.keys(),
                              key=lambda k: self.categories[k].get('order', 99)):
            cat_certs = [c for (ck, _), c in matched.items() if ck == cat_key]
            cat_certs.sort(key=lambda c: c['batch_number'])
            if cat_certs:
                result.append({'category': cat_key, 'certs': cat_certs})

        return result

    def _parse_spec(self, spec_str, category):
        """Parse '厚度=3, 牌号=Q345B' into {'thickness': '3', 'material_grade': 'Q345B'}
        Only matches against params defined for the given category."""
        result = {}
        if not spec_str:
            return result
        cat_def = self.categories.get(category, {})
        cat_params = cat_def.get('params', [])
        for part in spec_str.split(','):
            part = part.strip()
            if '=' in part:
                k, v = part.split('=', 1)
                k, v = k.strip(), v.strip()
                for p in cat_params:
                    short = p.get('short', '')
                    label_short = p['label'].split('(')[0] if '(' in p['label'] else p['label']
                    if k in (p['key'], p['label'], short, label_short):
                        result[p['key']] = v
                        break
        return result

    def _params_match(self, cert_params, req_spec, category):
        """Check if cert_params match the required spec.
        All keys in req_spec must exist and match in cert_params."""
        if not req_spec:
            return True  # no spec requirement = match any cert in category
        for key, val in req_spec.items():
            cert_val = cert_params.get(key, '')
            if str(cert_val).strip() != str(val).strip():
                return False
        return True

    # =================================================================
    # Report Insertion
    # =================================================================

    def _get_image_orientation(self, image_path):
        """Return 'landscape' if image width > height, else 'portrait'."""
        try:
            from PIL import Image
            with Image.open(image_path) as img:
                w, h = img.size
                return 'landscape' if w > h else 'portrait'
        except Exception:
            return 'portrait'

    def _make_section_break(self, orientation='portrait'):
        """Create a paragraph with w:sectPr that ends the current section
        and sets its page orientation. The section break also acts as a
        page break (nextPage is the default type)."""
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        sectPr = OxmlElement('w:sectPr')

        pgSz = OxmlElement('w:pgSz')
        if orientation == 'landscape':
            pgSz.set(qn('w:w'), '16839')
            pgSz.set(qn('w:h'), '11906')
            pgSz.set(qn('w:orient'), 'landscape')
        else:
            pgSz.set(qn('w:w'), '11906')
            pgSz.set(qn('w:h'), '16839')
        sectPr.append(pgSz)

        pPr.append(sectPr)
        p.append(pPr)
        return p

    def _make_page_break_paragraph(self):
        """Create a paragraph containing a hard page break."""
        p = OxmlElement('w:p')
        r = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        r.append(br)
        p.append(r)
        return p

    def _set_section_orientation(self, sectPr, orientation='portrait'):
        """Set page orientation on an existing w:sectPr element."""
        pgSz = sectPr.find(qn('w:pgSz'))
        if pgSz is None:
            pgSz = OxmlElement('w:pgSz')
            sectPr.append(pgSz)

        if orientation == 'landscape':
            pgSz.set(qn('w:w'), '16839')
            pgSz.set(qn('w:h'), '11906')
            pgSz.set(qn('w:orient'), 'landscape')
        else:
            pgSz.set(qn('w:w'), '11906')
            pgSz.set(qn('w:h'), '16839')
            orient_key = qn('w:orient')
            if orient_key in pgSz.attrib:
                del pgSz.attrib[orient_key]

    def _remove_page_break_before_anchor(self, body_children, anchor_idx):
        """Remove a hard page break immediately before an insertion anchor.
        The material certificate insertion starts with a next-page section
        break, so keeping an existing hard page break before it creates a
        blank page."""
        if anchor_idx <= 0:
            return 0

        removed = 0
        for i in range(anchor_idx - 1, -1, -1):
            child = body_children[i]
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'p':
                pPr = child.find(qn('w:pPr'))
                if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                    break

                page_breaks = [
                    br for br in child.iter(qn('w:br'))
                    if br.get(qn('w:type')) == 'page'
                ]
                if page_breaks:
                    for br in page_breaks:
                        parent = br.getparent()
                        if parent is not None:
                            parent.remove(br)
                            removed += 1

                    text = ''.join(t.text or '' for t in child.iter(qn('w:t'))).strip()
                    has_image = any(child.findall('.//' + qn(t_))
                                    for t_ in ('w:drawing', 'w:pict', 'w:object'))
                    has_break = any(br.get(qn('w:type')) == 'page'
                                    for br in child.iter(qn('w:br')))
                    if not text and not has_image and not has_break:
                        child.getparent().remove(child)
                    break

                text = ''.join(t.text or '' for t in child.iter(qn('w:t'))).strip()
                has_image = any(child.findall('.//' + qn(t_))
                                for t_ in ('w:drawing', 'w:pict', 'w:object'))
                if text or has_image:
                    break
                continue

            break

        return removed

    def _find_material_anchor_index(self, body_children):
        """Resolve the insertion anchor for material certificate pages."""
        if self.product_type == 'viscous_damper':
            for i in range(len(body_children) - 1, -1, -1):
                child = body_children[i]
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'sectPr':
                    return i
            return len(body_children) - 1 if body_children else 0

        for i in range(len(body_children) - 1, -1, -1):
            child = body_children[i]
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'p':
                text = ''.join(t.text or '' for t in child.iter(qn('w:t')))
                if '产品质量承诺书' in text:
                    return i
        return None

    def insert_material_certs(self, doc, matched_certs):
        """Append material certificate pages before '产品质量承诺书'.
        Each certificate gets its own section with page orientation
        matching the image aspect ratio."""
        if not matched_certs:
            return

        body = doc.element.body
        body_children = list(body)

        anchor_idx = self._find_material_anchor_index(body_children)

        if anchor_idx is None:
            print('Material cert anchor not found, appending at end.')
            anchor_idx = len(body_children) - 1

        anchor_elem = body_children[anchor_idx]
        removed_breaks = self._remove_page_break_before_anchor(body_children, anchor_idx)
        if removed_breaks:
            body_children = list(body)
            anchor_idx = body_children.index(anchor_elem)

        elements_to_insert = []
        anchor_tag = anchor_elem.tag.split('}')[-1] if '}' in anchor_elem.tag else anchor_elem.tag

        # When materials are appended at the very end of a report, avoid
        # stacking a break on top of an existing next-page section break.
        if anchor_tag == 'sectPr':
            prev_has_section_break = False
            if anchor_idx > 0:
                prev = body_children[anchor_idx - 1]
                prev_pPr = prev.find(qn('w:pPr'))
                prev_has_section_break = (
                    prev_pPr is not None and
                    prev_pPr.find(qn('w:sectPr')) is not None
                )
            if not prev_has_section_break:
                elements_to_insert.append(self._make_section_break('portrait'))
        else:
            # Section break to close the previous section as portrait before
            # the first material certificate. This prevents the previous
            # section from inheriting the first cert's orientation.
            elements_to_insert.append(self._make_section_break('portrait'))

        for group_idx, group in enumerate(matched_certs):
            cat_key = group['category']
            cat_config = self.categories.get(cat_key, {})
            cat_label = cat_config.get('label', cat_key)
            certs = group['certs']

            # Category title
            title_para = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            keep_next = OxmlElement('w:keepNext')
            pPr.append(keep_next)
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr.append(jc)
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), '0')
            spacing.set(qn('w:after'), '120')
            pPr.append(spacing)
            title_para.append(pPr)
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            b = OxmlElement('w:b')
            rPr.append(b)
            r.append(rPr)
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = f'{cat_label}材质证明书'
            r.append(t)
            title_para.append(r)
            elements_to_insert.append(title_para)

            for cert_idx, cert in enumerate(certs):
                img_path = self.get_image_path(cert)
                if img_path and os.path.exists(img_path):
                    report_img_path = self.get_report_image_path(cert)
                    img_para = self._make_image_paragraph(doc, report_img_path)
                    elements_to_insert.append(img_para)
                    orientation = self._get_image_orientation(report_img_path)
                else:
                    orientation = 'portrait'

                is_last_cert = (
                    group_idx == len(matched_certs) - 1 and
                    cert_idx == len(certs) - 1
                )

                if is_last_cert:
                    if anchor_tag == 'sectPr':
                        self._set_section_orientation(anchor_elem, orientation)
                    else:
                        # Last cert transitions back to portrait before the
                        # following quality commitment chapter.
                        elements_to_insert.append(
                            self._make_section_break('portrait'))
                else:
                    elements_to_insert.append(
                        self._make_section_break(orientation))

        # Insert all elements before anchor (forward order:
        # each addprevious inserts immediately before ref_elem,
        # so iterating [A,B,C] yields [A, B, C, anchor])
        ref_elem = body_children[anchor_idx]
        for elem in elements_to_insert:
            ref_elem.addprevious(elem)

        total = sum(len(g['certs']) for g in matched_certs)
        print(f'Inserted {total} material certificate page(s) '
              f'({len(matched_certs)} categories).')

    def _make_image_paragraph(self, doc, image_path):
        """Create a centered paragraph containing a DrawingML inline image."""
        para = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        jc = OxmlElement('w:jc')
        jc.set(qn('w:val'), 'center')
        pPr.append(jc)
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')
        pPr.append(spacing)
        para.append(pPr)

        run = OxmlElement('w:r')
        para.append(run)
        orientation = self._get_image_orientation(image_path)
        self._add_image_to_run(doc, run, image_path, orientation)
        return para

    def _add_image_to_run(self, doc, run, image_path, orientation='portrait'):
        """Add a DrawingML inline image to a w:r element."""
        try:
            from PIL import Image
            with Image.open(image_path) as img:
                img_w, img_h = img.size
        except Exception:
            img_w, img_h = 800, 600

        aspect = img_h / img_w if img_w > 0 else 1.0
        # A4 content area with conservative Word default margins, reserving
        # a little vertical space for the material certificate title.
        if orientation == 'landscape':
            max_width_inches = 9.7
            max_height_inches = 5.8
        else:
            max_width_inches = 6.5
            max_height_inches = 8.9

        width_inches = max_width_inches
        height_inches = width_inches * aspect
        if height_inches > max_height_inches:
            height_inches = max_height_inches
            width_inches = height_inches / aspect if aspect > 0 else max_width_inches

        width_emu = int(width_inches * 914400)
        height_emu = int(height_inches * 914400)

        r_id, _ = doc.part.get_or_add_image(image_path)

        drawing = OxmlElement('w:drawing')
        inline = OxmlElement('wp:inline')

        extent = OxmlElement('wp:extent')
        extent.set('cx', str(width_emu))
        extent.set('cy', str(height_emu))
        inline.append(extent)

        docPr = OxmlElement('wp:docPr')
        docPr.set('id', str(doc.part.next_id))
        docPr.set('name', os.path.basename(image_path))
        inline.append(docPr)

        cNvGraphicFramePr = OxmlElement('wp:cNvGraphicFramePr')
        inline.append(cNvGraphicFramePr)

        graphic = OxmlElement('a:graphic')
        graphicData = OxmlElement('a:graphicData')
        graphicData.set('uri',
                         'http://schemas.openxmlformats.org/drawingml/2006/picture')

        pic = OxmlElement('pic:pic')
        nvPicPr = OxmlElement('pic:nvPicPr')
        cNvPr = OxmlElement('pic:cNvPr')
        cNvPr.set('id', '0')
        cNvPr.set('name', os.path.basename(image_path))
        nvPicPr.append(cNvPr)
        cNvPicPr = OxmlElement('pic:cNvPicPr')
        nvPicPr.append(cNvPicPr)
        pic.append(nvPicPr)

        blipFill = OxmlElement('pic:blipFill')
        blip = OxmlElement('a:blip')
        blip.set(qn('r:embed'), r_id)
        blipFill.append(blip)
        stretch = OxmlElement('a:stretch')
        blipFill.append(stretch)
        pic.append(blipFill)

        spPr = OxmlElement('pic:spPr')
        xfrm = OxmlElement('a:xfrm')
        off = OxmlElement('a:off')
        off.set('x', '0')
        off.set('y', '0')
        xfrm.append(off)
        ext = OxmlElement('a:ext')
        ext.set('cx', str(width_emu))
        ext.set('cy', str(height_emu))
        xfrm.append(ext)
        spPr.append(xfrm)
        prstGeom = OxmlElement('a:prstGeom')
        prstGeom.set('prst', 'rect')
        spPr.append(prstGeom)
        pic.append(spPr)

        graphicData.append(pic)
        graphic.append(graphicData)
        inline.append(graphic)
        drawing.append(inline)
        run.append(drawing)
