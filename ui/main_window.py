"""
Main window for the Report Generation System.
"""
import os
import sys
import json
from datetime import datetime
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTableWidget, QTableWidgetItem,
    QGroupBox, QTabWidget, QFileDialog, QMessageBox, QStatusBar,
    QHeaderView, QSplitter, QFrame, QComboBox, QToolBar, QSizePolicy,
    QAbstractItemView,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QAction, QIcon, QFont

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.excel_reader import read_excel_data
from core.word_engine import generate_report
from core.pdf_converter import docx_to_pdf
from core.archiver import open_archive_dir
from products.base_product import load_product, list_products
from ui.image_manager import ImageManagerWidget

HYSTERESIS_FIELD = '\u662f\u5426\u9700\u8981\u6ede\u56de\u66f2\u7ebf'
YES_TEXT = '\u662f'
NO_TEXT = '\u5426'
ED_TYPE_FIELD = '\u9884\u57cb\u4ef6\u7c7b\u578b'
ED_TYPE_DEFAULT = '\u7c98\u6ede\u963b\u5c3c\u5668'


class ReportGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.product = None
        self.excel_data = None
        self.current_docx_path = None
        self.material_manager_widget = None
        self._material_widgets = {}  # Cache: product_type → ImageManagerWidget

        self.setWindowTitle('融海运通出厂检验报告生成系统')
        self.setMinimumSize(1000, 650)
        self.resize(1200, 800)

        self._setup_ui()
        self._load_products()
        self._apply_theme()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(12, 8, 12, 12)
        main_layout.setSpacing(8)

        # Toolbar
        self._create_toolbar()

        # Content area
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(8)

        # Left panel
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 6, 0)
        left_layout.setSpacing(8)

        # Project info group
        self._create_project_info_group(left_layout)

        # Product list group
        self._create_product_list_group(left_layout)

        left_panel.setLayout(left_layout)
        splitter.addWidget(left_panel)

        # Right panel
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(6, 0, 0, 0)
        right_layout.setSpacing(8)

        self.tabs = QTabWidget()
        self._create_mechanical_tab()
        self._create_visual_tab()
        self._create_image_tab()
        right_layout.addWidget(self.tabs)

        right_panel.setLayout(right_layout)
        splitter.addWidget(right_panel)

        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)

        main_layout.addWidget(splitter)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage('就绪 — 请导入 Excel 参数表开始')

        designed_by = QLabel('Designed by YaoyangXu')
        designed_by.setStyleSheet('color: #888; font-size: 11px; margin-right: 8px;')
        self.status_bar.addPermanentWidget(designed_by)

    def _create_toolbar(self):
        tb = QToolBar()
        tb.setIconSize(QSize(20, 20))
        tb.setMovable(False)
        self.addToolBar(tb)

        # Product selector
        tb.addWidget(QLabel('产品类型: '))
        self.product_combo = QComboBox()
        self.product_combo.setMinimumWidth(120)
        tb.addWidget(self.product_combo)
        tb.addSeparator()

        # Actions
        self.btn_import = QPushButton('导入 Excel')
        self.btn_import.clicked.connect(self.on_import_excel)
        tb.addWidget(self.btn_import)

        tb.addSeparator()

        self.btn_word = QPushButton('生成 Word')
        self.btn_word.setEnabled(False)
        self.btn_word.clicked.connect(self.on_generate_word)
        tb.addWidget(self.btn_word)

        self.btn_pdf = QPushButton('生成 PDF')
        self.btn_pdf.setEnabled(False)
        self.btn_pdf.clicked.connect(self.on_generate_pdf)
        tb.addWidget(self.btn_pdf)

        tb.addSeparator()

        self.btn_archive = QPushButton('打开归档')
        self.btn_archive.setObjectName('btnSecondary')
        self.btn_archive.clicked.connect(self.on_open_archive)
        tb.addWidget(self.btn_archive)

    def _create_project_info_group(self, parent_layout):
        group = QGroupBox('项目基本信息')
        layout = QVBoxLayout(group)

        self.project_fields = {}
        labels = ['项目名称', '产品供应商', '制造地址', '联系电话', '制造日期', '检验员', '审核']
        for label in labels:
            row = QHBoxLayout()
            lbl = QLabel(f'{label}:')
            lbl.setFixedWidth(80)
            row.addWidget(lbl)
            edit = QLineEdit()
            edit.setPlaceholderText(f'请输入{label}')
            row.addWidget(edit)
            layout.addLayout(row)
            self.project_fields[label] = edit

        self.embedded_type_row_widget = QWidget()
        row = QHBoxLayout(self.embedded_type_row_widget)
        row.setContentsMargins(0, 0, 0, 0)
        lbl = QLabel(f'{ED_TYPE_FIELD}:')
        lbl.setFixedWidth(120)
        row.addWidget(lbl)
        edit = QLineEdit()
        edit.setPlaceholderText(ED_TYPE_DEFAULT)
        row.addWidget(edit)
        layout.addWidget(self.embedded_type_row_widget)
        self.project_fields[ED_TYPE_FIELD] = edit

        self.hysteresis_row_widget = QWidget()
        row = QHBoxLayout(self.hysteresis_row_widget)
        row.setContentsMargins(0, 0, 0, 0)
        lbl = QLabel(f'{HYSTERESIS_FIELD}:')
        lbl.setFixedWidth(120)
        row.addWidget(lbl)
        combo = QComboBox()
        combo.addItems([YES_TEXT, NO_TEXT])
        combo.setCurrentText(YES_TEXT)
        row.addWidget(combo)
        layout.addWidget(self.hysteresis_row_widget)
        self.project_fields[HYSTERESIS_FIELD] = combo

        parent_layout.addWidget(group)

    def _default_project_field_value(self, label):
        if label == HYSTERESIS_FIELD:
            return YES_TEXT
        if label == ED_TYPE_FIELD:
            return ED_TYPE_DEFAULT
        return ''

    def _set_project_field_value(self, widget, value, label):
        text = str(value or self._default_project_field_value(label))
        if isinstance(widget, QComboBox):
            index = widget.findText(text)
            if index < 0:
                index = widget.findText(self._default_project_field_value(label))
            if index >= 0:
                widget.setCurrentIndex(index)
            return
        widget.setText(text)

    def _get_project_field_value(self, widget):
        if isinstance(widget, QComboBox):
            return widget.currentText()
        return widget.text()

    def _create_product_list_group(self, parent_layout):
        group = QGroupBox('产品型号清单')
        layout = QVBoxLayout(group)

        self.product_table = QTableWidget()
        self.product_table.setColumnCount(5)
        self.product_table.setHorizontalHeaderLabels(
            ['产品型号', '支座编号范围', '数量', '生产日期', '检验标准']
        )
        self.product_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.product_table.horizontalHeader().setStretchLastSection(True)
        self.product_table.setAlternatingRowColors(True)
        self.product_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.product_table.setMinimumHeight(200)
        layout.addWidget(self.product_table)

        parent_layout.addWidget(group)

    def _create_mechanical_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        self.mech_table = QTableWidget()
        self.mech_table.setColumnCount(10)
        self.mech_table.setHorizontalHeaderLabels([
            '产品型号', '有效直径\n(mm)', '高度\n(mm)', '橡胶层厚\n(mm)',
            '压缩应力\n(MPa)', '压缩刚度\n(KN/mm)', '水平等效刚度\n(KN/mm)',
            '屈服后刚度\n(KN/mm)', '屈服力\n(KN)', '阻尼比\n(%)'
        ])
        self.mech_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.mech_table.horizontalHeader().setStretchLastSection(True)
        self.mech_table.setAlternatingRowColors(True)
        layout.addWidget(self.mech_table)

        self.tabs.addTab(widget, '力学性能检测')

    def _create_visual_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        self.visual_table = QTableWidget()
        self.visual_table.setColumnCount(10)
        self.visual_table.setHorizontalHeaderLabels([
            '产品型号', '气泡', '杂质', '缺陷', '凹凸不平',
            '胶料粘结', '裂纹', '钢板外露', '产品标识', '尺寸偏差'
        ])
        self.visual_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.visual_table.horizontalHeader().setStretchLastSection(True)
        self.visual_table.setAlternatingRowColors(True)
        layout.addWidget(self.visual_table)

        self.tabs.addTab(widget, '外观尺寸检测')

    def _create_image_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        self.image_tab_container = widget
        self.image_tab_layout = layout
        self.tabs.addTab(widget, '材质单图片')

    def _apply_theme(self):
        qss_path = os.path.join(os.path.dirname(__file__), 'resources', 'style.qss')
        if os.path.exists(qss_path):
            with open(qss_path, 'r', encoding='utf-8') as f:
                self.setStyleSheet(f.read())

    def _settings_path(self):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(project_root, 'config', 'settings.json')

    def _load_settings(self):
        path = self._settings_path()
        if not os.path.exists(path):
            return {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (OSError, json.JSONDecodeError):
            return {}

    def _save_settings(self, settings):
        path = self._settings_path()
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)

    def _load_products(self):
        try:
            products = list_products()
            settings = self._load_settings()
            last_product_type = settings.get('last_product_type', '')
            self.product_combo.clear()
            for p in products:
                self.product_combo.addItem(p['name'], p['type'])
            if products:
                product_types = [p['type'] for p in products]
                initial_index = (
                    product_types.index(last_product_type)
                    if last_product_type in product_types else 0
                )
                self.product_combo.setCurrentIndex(initial_index)
                self._on_product_changed(products[initial_index]['type'])
        except Exception as e:
            self.status_bar.showMessage(f'加载产品列表失败: {e}')

        self.product_combo.currentIndexChanged.connect(
            lambda idx: self._on_product_changed(self.product_combo.currentData())
        )

    def _on_product_changed(self, product_type):
        if not product_type:
            return
        try:
            self.product = load_product(product_type)
            self.setWindowTitle(f'融海运通 - {self.product.product_name}出厂检验报告生成系统')
            self.status_bar.showMessage(f'已选择: {self.product.product_name}')

            # Reconfigure UI tabs for this product
            self._reconfigure_tabs()
            self._update_project_field_visibility()

            # Swap material manager widget (cached per product type)
            if self.material_manager_widget:
                self.material_manager_widget.setVisible(False)

            if product_type in self._material_widgets:
                self.material_manager_widget = self._material_widgets[product_type]
                self.material_manager_widget.setVisible(True)
            else:
                self.material_manager_widget = ImageManagerWidget(
                    self.product.product_dir, self.product.product_type)
                self._material_widgets[product_type] = self.material_manager_widget
                self.image_tab_layout.addWidget(self.material_manager_widget)

            settings = self._load_settings()
            settings['last_product_type'] = product_type
            self._save_settings(settings)
        except Exception as e:
            self.status_bar.showMessage(f'切换产品失败: {e}')

    def _reconfigure_tabs(self):
        """Update table headers based on product ui_config."""
        ui = self.product.config.get('ui_config', {})
        if not ui:
            return

        # Product table
        pt = ui.get('product_table', {})
        if pt.get('headers'):
            self.product_table.setColumnCount(len(pt['headers']))
            self.product_table.setHorizontalHeaderLabels(pt['headers'])
            self.product_table.horizontalHeader().setSectionResizeMode(
                QHeaderView.Interactive)
            self.product_table.horizontalHeader().setStretchLastSection(True)

        # Mechanical tab
        mt = ui.get('mechanical_tab', {})
        if mt.get('headers') and self.product.has_mechanical_detail:
            self.mech_table.setColumnCount(len(mt['headers']))
            self.mech_table.setHorizontalHeaderLabels(mt['headers'])
            self.mech_table.horizontalHeader().setSectionResizeMode(
                QHeaderView.Interactive)
            self.mech_table.horizontalHeader().setStretchLastSection(True)
            # Show mechanical tab
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == '力学性能检测':
                    self.tabs.setTabVisible(i, True)
                    break
        else:
            # Hide mechanical tab for products without it
            for i in range(self.tabs.count()):
                if self.tabs.tabText(i) == '力学性能检测':
                    self.tabs.setTabVisible(i, False)
                    break

        # Visual tab
        vt = ui.get('visual_tab', {})
        if vt.get('headers'):
            self.visual_table.setColumnCount(len(vt['headers']))
            self.visual_table.setHorizontalHeaderLabels(vt['headers'])
            self.visual_table.horizontalHeader().setSectionResizeMode(
                QHeaderView.Interactive)
            self.visual_table.horizontalHeader().setStretchLastSection(True)

    def _update_project_field_visibility(self):
        if hasattr(self, 'hysteresis_row_widget'):
            self.hysteresis_row_widget.setVisible(
                self.product.product_type == 'viscous_damper')
        if hasattr(self, 'embedded_type_row_widget'):
            self.embedded_type_row_widget.setVisible(
                self.product.product_type == 'embedded_damper_parts')

    # ===== Event Handlers =====

    def on_import_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择 Excel 参数文件',
            os.path.dirname(self.product.excel_template_path) if self.product else '',
            'Excel 文件 (*.xlsx)'
        )
        if not file_path:
            return

        try:
            self.excel_data = read_excel_data(file_path)
            self._populate_ui_from_data()
            self.btn_word.setEnabled(True)
            self.btn_pdf.setEnabled(True)
            self.status_bar.showMessage(f'导入成功 — {file_path}')
        except Exception as e:
            QMessageBox.warning(self, '导入失败', str(e))

    def _populate_ui_from_data(self):
        data = self.excel_data
        ui = self.product.config.get('ui_config', {})

        # Project info
        project = data.get('project_info', {})
        for label, edit in self.project_fields.items():
            self._set_project_field_value(edit, project.get(label, ''), label)

        # Product list
        products = data.get('product_list', [])
        self.product_table.setRowCount(len(products))
        pt_keys = ui.get('product_table', {}).get('data_keys',
            ['产品型号', '支座编号范围', '数量', '生产日期', '检验标准'])
        for r, p in enumerate(products):
            for c, k in enumerate(pt_keys):
                val = p.get(k, '')
                self.product_table.setItem(r, c, QTableWidgetItem(str(val) if val else ''))

        # Mechanical data
        mech_data = data.get('mechanical_data', [])
        self.mech_table.setRowCount(len(mech_data))
        mech_keys = ui.get('mechanical_tab', {}).get('data_keys',
            ['产品型号', '支座有效直径(mm)', '支座高度(mm)', '橡胶层总厚(mm)',
             '压缩应力(MPa)', '设计压缩刚度(KN/mm)', '水平等效刚度(KN/mm)',
             '屈服后刚度(KN/mm)', '屈服力(KN)', '等效阻尼比(%)'])
        for r, m in enumerate(mech_data):
            for c, k in enumerate(mech_keys):
                self.mech_table.setItem(r, c, QTableWidgetItem(str(m.get(k, ''))))

        # Visual data
        vis_data = data.get('visual_data', [])
        self.visual_table.setRowCount(len(vis_data))
        vis_keys = ui.get('visual_tab', {}).get('data_keys',
            ['产品型号',
             '检测1-气泡(合格/不合格)', '检测2-杂质(合格/不合格)',
             '检测3-缺陷(合格/不合格)', '检测4-凹凸不平(合格/不合格)',
             '检测5-胶料粘结不良(合格/不合格)', '检测6-裂纹(合格/不合格)',
             '检测7-钢板外露(合格/不合格)', '检测8-产品标识(合格/不合格)',
             '检测9-平面尺寸偏差(合格/不合格)'])
        for r, v in enumerate(vis_data):
            for c, k in enumerate(vis_keys):
                self.visual_table.setItem(r, c,
                                         QTableWidgetItem(str(v.get(k, '合格'))))

        # Auto-size columns to content, then stretch last column
        for table in [self.product_table, self.mech_table, self.visual_table]:
            table.resizeColumnsToContents()
            header = table.horizontalHeader()
            # Ensure minimum reasonable width for first column (model name)
            if header.sectionSize(0) < 100:
                header.resizeSection(0, 100)

    def _collect_ui_data(self):
        """Read current UI values back into excel_data dict."""
        if not self.excel_data:
            return None

        ui = self.product.config.get('ui_config', {})

        # Update project info
        for label, edit in self.project_fields.items():
            if label == HYSTERESIS_FIELD and self.product.product_type != 'viscous_damper':
                continue
            if label == ED_TYPE_FIELD and self.product.product_type != 'embedded_damper_parts':
                continue
            self.excel_data['project_info'][label] = self._get_project_field_value(edit)

        # Update product list
        pt_keys = ui.get('product_table', {}).get('data_keys',
            ['产品型号', '支座编号范围', '数量', '生产日期', '检验标准'])
        products = []
        for r in range(self.product_table.rowCount()):
            p = {}
            for c, h in enumerate(pt_keys):
                item = self.product_table.item(r, c)
                p[h] = item.text() if item else ''
            if p.get(pt_keys[0]):
                products.append(p)
        self.excel_data['product_list'] = products

        # Update mechanical data
        mech_keys = ui.get('mechanical_tab', {}).get('data_keys',
            ['产品型号', '支座有效直径(mm)', '支座高度(mm)', '橡胶层总厚(mm)',
             '压缩应力(MPa)', '设计压缩刚度(KN/mm)', '水平等效刚度(KN/mm)',
             '屈服后刚度(KN/mm)', '屈服力(KN)', '等效阻尼比(%)'])
        mech_data = []
        for r in range(self.mech_table.rowCount()):
            m = {}
            model = self.mech_table.item(r, 0)
            if not model or not model.text().strip():
                continue
            for c, k in enumerate(mech_keys):
                item = self.mech_table.item(r, c)
                m[k] = item.text() if item else ''
            mech_data.append(m)
        self.excel_data['mechanical_data'] = mech_data

        # Update visual data
        vis_keys = ui.get('visual_tab', {}).get('data_keys',
            ['产品型号',
             '检测1-气泡(合格/不合格)', '检测2-杂质(合格/不合格)',
             '检测3-缺陷(合格/不合格)', '检测4-凹凸不平(合格/不合格)',
             '检测5-胶料粘结不良(合格/不合格)', '检测6-裂纹(合格/不合格)',
             '检测7-钢板外露(合格/不合格)', '检测8-产品标识(合格/不合格)',
             '检测9-平面尺寸偏差(合格/不合格)'])
        vis_data = []
        for r in range(self.visual_table.rowCount()):
            v = {}
            model = self.visual_table.item(r, 0)
            if not model or not model.text().strip():
                continue
            for c, k in enumerate(vis_keys):
                item = self.visual_table.item(r, c)
                v[k] = item.text() if item else '合格'
            vis_data.append(v)
        self.excel_data['visual_data'] = vis_data

        return self.excel_data

    def on_generate_word(self):
        data = self._collect_ui_data()
        if not data:
            return

        try:
            output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'output')
            product_dir = os.path.dirname(self.product.template_path)
            material_mgr = (self.material_manager_widget.manager
                            if self.material_manager_widget else None)
            selected_ids = (self.material_manager_widget.get_checked_cert_ids()
                            if self.material_manager_widget else None)
            path = generate_report(
                self.product.template_path,
                data,
                output_dir,
                f'{self.product.product_name}出厂检验报告',
                product_dir,
                material_manager=material_mgr,
                selected_cert_ids=selected_ids,
                product_type=self.product.product_type,
            )
            self.current_docx_path = path

            self.status_bar.showMessage(f'Word 报告已生成 — {path}')
            QMessageBox.information(self, '生成成功',
                                    f'报告已生成:\n{path}\n\n已自动归档到 output 目录。')
        except Exception as e:
            QMessageBox.critical(self, '生成失败', str(e))

    def on_generate_pdf(self):
        if not self.current_docx_path:
            # Generate Word first
            self.on_generate_word()
            if not self.current_docx_path:
                return

        try:
            pdf_path = self.current_docx_path.replace('.docx', '.pdf')
            docx_to_pdf(self.current_docx_path, pdf_path)
            self.status_bar.showMessage(f'PDF 报告已生成 — {pdf_path}')
            QMessageBox.information(self, 'PDF 生成成功',
                                    f'PDF 已生成:\n{pdf_path}')
        except Exception as e:
            QMessageBox.critical(self, 'PDF 生成失败',
                                f'{str(e)}\n\n请确保本机已安装 Microsoft Word。')

    def on_open_archive(self):
        try:
            output_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'output')
            open_archive_dir(output_dir)
        except Exception as e:
            QMessageBox.warning(self, '打开失败', str(e))


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont('Microsoft YaHei', 10))
    window = ReportGenerator()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
