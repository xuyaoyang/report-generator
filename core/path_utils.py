"""Helpers for safe Windows path and filename handling."""
import os
import re


WINDOWS_INVALID_CHARS = '<>:"/\\|?*'
_INVALID_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
_RESERVED_NAMES = {
    'CON', 'PRN', 'AUX', 'NUL',
    *(f'COM{i}' for i in range(1, 10)),
    *(f'LPT{i}' for i in range(1, 10)),
}


def sanitize_path_component(value, default='未命名', max_length=80):
    """Return a single safe Windows path component."""
    text = str(value or '').strip()
    text = _INVALID_RE.sub('_', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.strip(' .')
    text = text.replace('..', '_')
    if not text or text in ('.', '..'):
        text = default
    if text.upper() in _RESERVED_NAMES:
        text = f'{text}_'
    if len(text) > max_length:
        text = text[:max_length].rstrip(' ._') or default
    return text


def sanitize_filename_stem(value, default='未命名', max_length=120):
    """Return a safe filename stem without extension."""
    return sanitize_path_component(value, default=default, max_length=max_length)


def ensure_inside_base(base_dir, target_path):
    """Resolve target_path and ensure it stays inside base_dir."""
    base = os.path.abspath(base_dir)
    target = os.path.abspath(target_path)
    if os.path.commonpath([base, target]) != base:
        raise ValueError(f'路径越界，已阻止写入: {target}')
    return target
