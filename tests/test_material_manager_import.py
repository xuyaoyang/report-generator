import json
import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from core.app_paths import ensure_workspace, set_work_dir
from core.material_manager import MaterialManager


def _create_source_library(path):
    steel_dir = path / '钢板'
    steel_dir.mkdir(parents=True)
    image = steel_dir / '批次A.jpg'
    image.write_bytes(b'fake-image')

    conn = sqlite3.connect(path / 'material_certs.db')
    conn.execute('''CREATE TABLE material_certificates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT NOT NULL,
        batch_number TEXT NOT NULL,
        params TEXT NOT NULL DEFAULT '{}',
        image_filename TEXT NOT NULL,
        original_filename TEXT NOT NULL,
        file_size INTEGER DEFAULT 0,
        notes TEXT DEFAULT '',
        is_default INTEGER DEFAULT 0,
        cert_date TEXT DEFAULT '',
        is_expired INTEGER DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        product_type TEXT DEFAULT '',
        rotation INTEGER DEFAULT 0
    )''')
    conn.execute('''INSERT INTO material_certificates
        (category, batch_number, params, image_filename, original_filename,
         file_size, notes, cert_date, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        ('钢板', '批次A', json.dumps({'thickness': '5'}, ensure_ascii=False),
         '批次A.jpg', '批次A.jpg', image.stat().st_size, '备注',
         '202606', '2026-06-26 10:00:00', '2026-06-26 10:00:00'))
    conn.commit()
    conn.close()


class MaterialManagerImportTest(unittest.TestCase):
    def test_import_library_merges_db_without_overwriting(self):
        with tempfile.TemporaryDirectory() as tmp:
            appdata = os.path.join(tmp, 'appdata')
            work = os.path.join(tmp, 'work')
            source = Path(tmp) / 'old_lib'

            with mock.patch.dict(os.environ, {'APPDATA': appdata}):
                ensure_workspace(work)
                _create_source_library(source)

                manager = MaterialManager(os.path.abspath('products/isolation_bearing'))
                stats = manager.import_library(str(source))

                self.assertEqual(stats['imported'], 1)
                self.assertEqual(stats['skipped'], 0)

                certs = manager.get_certificates_by_category('钢板')
                self.assertEqual(len(certs), 1)
                self.assertEqual(certs[0]['batch_number'], '批次A')
                self.assertTrue(os.path.exists(manager.get_image_path(certs[0])))

                set_work_dir(work)


if __name__ == '__main__':
    unittest.main()
