"""
Abstract base class for product types.
Each product type (隔震支座, 预埋件, 黏滞阻尼器, etc.) provides its own config.
"""
from abc import ABC, abstractmethod
import os
import json


class BaseProduct(ABC):
    """Base class for all product report types."""

    def __init__(self, product_dir, product_type=''):
        self.product_dir = product_dir
        self.product_type = product_type
        self.config = self._load_config()

    def _load_config(self):
        config_path = os.path.join(self.product_dir, 'config.json')
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    @property
    def product_name(self):
        return self.config.get('product_name', '未命名产品')

    @property
    def template_path(self):
        rel = self.config.get('template', 'template_prepared.docx')
        return os.path.join(self.product_dir, rel)

    @property
    def excel_template_path(self):
        rel = self.config.get('excel_template', 'excel_template.xlsx')
        return os.path.join(self.product_dir, rel)

    @property
    def mapping_path(self):
        rel = self.config.get('mapping', 'param_mapping.json')
        return os.path.join(self.product_dir, rel)

    def get_sheet_names(self):
        """Return the expected Excel sheet names."""
        return self.config.get('sheets', [])

    @property
    def has_mechanical_detail(self):
        """Whether this product has mechanical performance detail tables."""
        return self.config.get('has_mechanical_detail', False)

    @property
    def has_per_model_visual(self):
        """Whether visual inspection reports are per-model (vs shared)."""
        return self.config.get('has_per_model_visual', True)

    @property
    def max_models_in_template(self):
        """Number of model slots in the base template."""
        return self.config.get('max_models_in_template', 6)

    @property
    def max_mech_models_per_page(self):
        """Models per mechanical detail page (for cloning)."""
        return self.config.get('max_mech_models_per_page', 4)

    @property
    def anchor_text(self):
        """Key text anchors for finding sections in the document."""
        return self.config.get('anchor_text', {})

    def validate_excel_data(self, data):
        """Validate Excel data. Returns list of error messages (empty = valid)."""
        errors = []
        project = data.get('project_info', {})
        if not project.get('项目名称'):
            errors.append('缺少项目名称')
        products = data.get('product_list', [])
        if not products:
            errors.append('产品型号清单为空')
        return errors


def load_product(product_type='isolation_bearing'):
    """
    Factory function to load a product by type.
    Returns a BaseProduct instance.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    product_dir = os.path.join(base_dir, product_type)

    if not os.path.exists(product_dir):
        raise FileNotFoundError(f'Product type not found: {product_type}')

    # For now, all products use the same base class
    # Future: specialized subclasses per product type
    return BaseProduct(product_dir, product_type)


def list_products():
    """List all available product types."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    products = []
    for name in os.listdir(base_dir):
        if name.startswith('_') or name.startswith('.'):
            continue
        product_dir = os.path.join(base_dir, name)
        if os.path.isdir(product_dir):
            config_path = os.path.join(product_dir, 'config.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                products.append({
                    'type': name,
                    'name': cfg.get('product_name', name),
                    'description': cfg.get('description', ''),
                })
    return products
