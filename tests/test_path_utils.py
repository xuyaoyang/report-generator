import os
import tempfile
import unittest

from core.path_utils import (
    ensure_inside_base, sanitize_filename_stem, sanitize_path_component,
)


class PathUtilsTest(unittest.TestCase):
    def test_sanitize_windows_path_component(self):
        self.assertEqual(sanitize_path_component('项目/A:B*C?"<>|'), '项目_A_B_C_')
        self.assertEqual(sanitize_path_component('  ..  '), '未命名')
        self.assertEqual(
            sanitize_filename_stem('批次\\2026/06:01'),
            '批次_2026_06_01')

    def test_ensure_inside_base_blocks_traversal(self):
        with tempfile.TemporaryDirectory() as tmp:
            base = os.path.join(tmp, 'base')
            os.makedirs(base)
            inside = ensure_inside_base(base, os.path.join(base, '项目', '报告.docx'))
            self.assertIn(os.path.abspath(base), inside)

            with self.assertRaises(ValueError):
                ensure_inside_base(base, os.path.join(tmp, 'outside.docx'))


if __name__ == '__main__':
    unittest.main()
