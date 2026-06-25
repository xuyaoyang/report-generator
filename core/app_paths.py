"""Runtime paths and first-run workspace initialization."""
import json
import os
import shutil
import sys


APP_DIR_NAME = '融海报告生成'
WORK_DIR_NAME = '报告生成工作目录'
EXCEL_TEMPLATE_DIR_NAME = 'Excel模板'
MATERIAL_LIB_DIR_NAME = '材质单库'
OUTPUT_DIR_NAME = '报告输出'


def app_root():
    """Return bundled app root in frozen mode, otherwise project root."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def user_config_dir():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, APP_DIR_NAME)


def settings_path():
    return os.path.join(user_config_dir(), 'settings.json')


def default_work_dir():
    documents = os.path.join(os.path.expanduser('~'), 'Documents')
    if not os.path.isdir(documents):
        documents = os.path.expanduser('~')
    return os.path.join(documents, WORK_DIR_NAME)


def load_settings():
    path = settings_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings):
    os.makedirs(user_config_dir(), exist_ok=True)
    with open(settings_path(), 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def work_dir():
    return load_settings().get('work_dir') or default_work_dir()


def excel_templates_dir():
    return os.path.join(work_dir(), EXCEL_TEMPLATE_DIR_NAME)


def material_lib_dir():
    return os.path.join(work_dir(), MATERIAL_LIB_DIR_NAME)


def output_dir():
    return load_settings().get('output_dir') or os.path.join(work_dir(), OUTPUT_DIR_NAME)


def user_material_categories_path():
    return os.path.join(work_dir(), 'material_categories.json')


def bundled_path(*parts):
    return os.path.join(app_root(), *parts)


def _copytree_missing(src, dst):
    if not os.path.exists(src):
        return
    if not os.path.exists(dst):
        shutil.copytree(src, dst)
        return
    for root, dirs, files in os.walk(src):
        rel = os.path.relpath(root, src)
        target_root = dst if rel == '.' else os.path.join(dst, rel)
        os.makedirs(target_root, exist_ok=True)
        for dirname in dirs:
            os.makedirs(os.path.join(target_root, dirname), exist_ok=True)
        for filename in files:
            target = os.path.join(target_root, filename)
            if not os.path.exists(target):
                shutil.copy2(os.path.join(root, filename), target)


def _product_template_name(product_dir, fallback_name):
    config_path = os.path.join(product_dir, 'config.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            name = json.load(f).get('product_name') or fallback_name
    except (OSError, json.JSONDecodeError):
        name = fallback_name
    return f'{name}参数模板.xlsx'


def initialize_workspace():
    """Create first-run user workspace without overwriting existing data."""
    settings = load_settings()
    if not settings.get('work_dir'):
        settings['work_dir'] = default_work_dir()
    if not settings.get('output_dir'):
        settings['output_dir'] = os.path.join(settings['work_dir'], OUTPUT_DIR_NAME)
    save_settings(settings)

    os.makedirs(settings['work_dir'], exist_ok=True)
    os.makedirs(settings['output_dir'], exist_ok=True)
    os.makedirs(excel_templates_dir(), exist_ok=True)

    _copytree_missing(bundled_path('image_lib'), material_lib_dir())

    src_categories = bundled_path('config', 'material_categories.json')
    dst_categories = user_material_categories_path()
    if os.path.exists(src_categories) and not os.path.exists(dst_categories):
        shutil.copy2(src_categories, dst_categories)

    products_root = bundled_path('products')
    if os.path.isdir(products_root):
        for product_name in os.listdir(products_root):
            product_dir = os.path.join(products_root, product_name)
            if not os.path.isdir(product_dir) or product_name.startswith('_'):
                continue
            src_template = os.path.join(product_dir, 'excel_template.xlsx')
            if not os.path.exists(src_template):
                continue
            dst_name = _product_template_name(product_dir, product_name)
            dst_template = os.path.join(excel_templates_dir(), dst_name)
            if not os.path.exists(dst_template):
                shutil.copy2(src_template, dst_template)
