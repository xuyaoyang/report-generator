"""
Material certificate image management panel.
"""
import os
import tempfile
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTreeWidget, QTreeWidgetItem, QScrollArea, QFrame,
    QGroupBox, QLineEdit, QFileDialog, QMessageBox, QSplitter,
    QGridLayout, QSizePolicy, QCheckBox, QStackedWidget,
    QDialog, QComboBox, QListWidget, QDialogButtonBox,
)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QPixmap, QDragEnterEvent, QDropEvent, QPalette, QTransform

from core.material_manager import MaterialManager


class ImageCard(QFrame):
    """A clickable thumbnail card for one material certificate image."""
    clicked = Signal(int)
    double_clicked = Signal(int, str)  # cert_id, image_path
    checked_changed = Signal(int, bool)  # cert_id, checked

    STYLE_NORMAL = (
        'ImageCard { border: 1px solid #d0d0d0; border-radius: 4px; '
        'background: white; }'
        'ImageCard:hover { border-color: #4A90D9; }'
    )
    STYLE_CHECKED = (
        'ImageCard { border: 2px solid #4CAF50; border-radius: 4px; '
        'background: #f0faf0; }'
        'ImageCard:hover { border-color: #4A90D9; }'
    )
    STYLE_SELECTED = (
        'ImageCard { border: 2px solid #4A90D9; border-radius: 4px; '
        'background: #f0f6ff; }'
    )

    def __init__(self, cert_data, image_path, parent=None, manager=None):
        super().__init__(parent)
        self.cert_id = cert_data['id']
        self._checked = False
        self._image_path = image_path or ''
        self._rotation = int(cert_data.get('rotation') or 0) % 360
        self._is_selected = False
        self._is_default = manager and manager.is_default_for(cert_data)
        self.setFixedSize(150, 170)
        self.setCursor(Qt.PointingHandCursor)
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet(self.STYLE_NORMAL)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 2, 6, 6)
        layout.setSpacing(2)

        # Top row: checkbox + default badge + spacer
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        self._check = QCheckBox()
        self._check.setToolTip('勾选以插入报告')
        self._check.setStyleSheet(self._CHECK_STYLE)
        self._check.toggled.connect(self._on_check_toggled)
        top_row.addWidget(self._check)
        self._default_badge = QLabel('默认')
        self._default_badge.setStyleSheet(
            'font-size: 9px; color: #B8860B; background: #FFF8DC; '
            'border: 1px solid #DAA520; border-radius: 2px; padding: 0 3px;')
        self._default_badge.setFixedHeight(16)
        self._default_badge.setVisible(self._is_default)
        top_row.addWidget(self._default_badge)
        top_row.addStretch()
        layout.addLayout(top_row)

        # Thumbnail
        self.thumb = QLabel()
        self.thumb.setFixedSize(138, 100)
        self.thumb.setAlignment(Qt.AlignCenter)
        self.thumb.setScaledContents(False)
        layout.addWidget(self.thumb)
        self._load_thumbnail()

        # Batch number
        self._batch_label = QLabel(cert_data.get('batch_number', '-'))
        self._batch_label.setAlignment(Qt.AlignCenter)
        self._batch_label.setStyleSheet('font-size: 11px; font-weight: bold; '
                                   'color: #333;')
        self._batch_label.setWordWrap(True)
        layout.addWidget(self._batch_label)

        # Category + key param
        cat = cert_data.get('category', '')
        params = cert_data.get('params', {})
        cat_config = cert_data.get('_cat_config', {})
        detail_parts = [cat]
        for p_def in cat_config.get('params', []):
            val = params.get(p_def['key'], '')
            if val:
                detail_parts.append(f'{val}')
                break
        self._detail_label = QLabel(' · '.join(detail_parts))
        self._detail_label.setAlignment(Qt.AlignCenter)
        self._detail_label.setStyleSheet('font-size: 10px; color: #888;')
        layout.addWidget(self._detail_label)

    def update_info(self, cert_data, image_path):
        """Update labels after cert data changed without recreating widget."""
        self._image_path = image_path or ''
        self._rotation = int(cert_data.get('rotation') or 0) % 360
        self._load_thumbnail()
        self._batch_label.setText(cert_data.get('batch_number', '-'))
        cat = cert_data.get('category', '')
        params = cert_data.get('params', {})
        cat_config = cert_data.get('_cat_config', {})
        detail_parts = [cat]
        for p_def in cat_config.get('params', []):
            val = params.get(p_def['key'], '')
            if val:
                detail_parts.append(f'{val}')
                break
        self._detail_label.setText(' · '.join(detail_parts))

    def _on_check_toggled(self, checked):
        self._checked = checked
        self._update_style()
        self.checked_changed.emit(self.cert_id, checked)

    _CHECK_STYLE = (
        'QCheckBox { background: transparent; spacing: 0px; }'
        'QCheckBox::indicator {'
        '    width: 20px; height: 20px;'
        '    border: 2px solid #888;'
        '    border-radius: 3px;'
        '    background: #f5f5f5;'
        '}'
        'QCheckBox::indicator:checked {'
        '    background: #2d7d2d;'
        '    border-color: #2d7d2d;'
        '}'
        'QCheckBox::indicator:hover {'
        '    border-color: #4A90D9;'
        '}'
    )

    def _update_style(self):
        if self._checked:
            self.setStyleSheet(self.STYLE_CHECKED)
        elif self._is_selected:
            self.setStyleSheet(self.STYLE_SELECTED)
        else:
            self.setStyleSheet(self.STYLE_NORMAL)
        # Qt quirk: parent stylesheet cascades to children, so re-apply
        self._check.setStyleSheet(self._CHECK_STYLE)

    def isChecked(self):
        return self._checked

    def setChecked(self, checked):
        self._checked = checked
        self._check.blockSignals(True)
        self._check.setChecked(checked)
        self._check.blockSignals(False)
        self._update_style()

    def setSelected(self, selected):
        self._is_selected = selected
        self._update_style()

    def setImagePath(self, path):
        self._image_path = path
        self._load_thumbnail()

    def _load_thumbnail(self):
        self.thumb.clear()
        self.thumb.setStyleSheet('')
        if self._image_path and os.path.exists(self._image_path):
            pix = QPixmap(self._image_path)
            if not pix.isNull():
                if self._rotation:
                    pix = pix.transformed(QTransform().rotate(self._rotation),
                                          Qt.SmoothTransformation)
                pix = pix.scaled(138, 100, Qt.KeepAspectRatio,
                                 Qt.SmoothTransformation)
                self.thumb.setPixmap(pix)
                return
            self.thumb.setText('图片无法加载')
        else:
            self.thumb.setText('无图片')
        self.thumb.setStyleSheet('color: #999;')

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and event.pos().y() > 22:
            self.clicked.emit(self.cert_id)
            return
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton and event.pos().y() > 22:
            self.double_clicked.emit(self.cert_id, self._image_path)
            return
        super().mouseDoubleClickEvent(event)


class _TitleBar(QWidget):
    """Draggable title bar for the image viewer."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._dragging = False
        self._drag_start = None
        self.setFixedHeight(42)
        self.setStyleSheet('background: #222;')
        self.setCursor(Qt.ArrowCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._dragging = True
            self._drag_start = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self._dragging:
            delta = event.globalPosition().toPoint() - self._drag_start
            self._drag_start = event.globalPosition().toPoint()
            w = self.window()
            w.move(w.x() + delta.x(), w.y() + delta.y())

    def mouseReleaseEvent(self, event):
        self._dragging = False

    def mouseDoubleClickEvent(self, event):
        w = self.window()
        if w.isMaximized():
            w.showNormal()
        else:
            w.showMaximized()


class _PannableScrollArea(QScrollArea):
    """Scroll area that supports click-drag panning of the image."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._panning = False
        self._pan_start = None

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._panning = True
            self._pan_start = event.globalPosition().toPoint()
            self.setCursor(Qt.ClosedHandCursor)
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._panning:
            delta = event.globalPosition().toPoint() - self._pan_start
            self._pan_start = event.globalPosition().toPoint()
            h = self.horizontalScrollBar()
            v = self.verticalScrollBar()
            if h:
                h.setValue(h.value() - delta.x())
            if v:
                v.setValue(v.value() - delta.y())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._panning = False
            self.setCursor(Qt.ArrowCursor)
        super().mouseReleaseEvent(event)


class ImageViewerDialog(QFrame):
    """A frameless, draggable, resizable window for viewing and rotating
    material certificate images."""

    _EDGE_MARGIN = 6
    _MIN_W = 400
    _MIN_H = 300

    def __init__(self, image_path, cert_info='', parent=None,
                 initial_rotation=0):
        super().__init__(parent)
        self._image_path = image_path
        self._rotation = int(initial_rotation or 0) % 360
        self._scale = 0.0

        self._resize_edge = None
        self._resizing = False
        self._resize_start = None
        self._resize_geom = None

        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setWindowTitle('材质单查看')
        self.setMinimumSize(self._MIN_W, self._MIN_H)
        self.setMouseTracking(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(1, 1, 1, 1)
        layout.setSpacing(0)

        # ── Title bar ──────────────────────────────────────────────
        title_bar = _TitleBar(self)
        tb_layout = QHBoxLayout(title_bar)
        tb_layout.setContentsMargins(12, 0, 8, 0)

        info_lbl = QLabel(cert_info)
        info_lbl.setStyleSheet('color: #ccc; font-size: 13px; background: transparent;')
        tb_layout.addWidget(info_lbl)
        tb_layout.addStretch()

        btn_style = (
            'QPushButton { background: #3a3a3a; color: #ccc; border: 1px solid #555; '
            'border-radius: 4px; padding: 5px 12px; font-size: 12px; }'
            'QPushButton:hover { background: #555; color: white; }'
        )

        self.btn_rotate_ccw = QPushButton('↺ 左转90°')
        self.btn_rotate_ccw.setStyleSheet(btn_style)
        self.btn_rotate_ccw.clicked.connect(lambda: self._rotate(-90))
        tb_layout.addWidget(self.btn_rotate_ccw)

        self.btn_rotate_cw = QPushButton('↻ 右转90°')
        self.btn_rotate_cw.setStyleSheet(btn_style)
        self.btn_rotate_cw.clicked.connect(lambda: self._rotate(90))
        tb_layout.addWidget(self.btn_rotate_cw)

        self.btn_reset = QPushButton('还原')
        self.btn_reset.setStyleSheet(btn_style)
        self.btn_reset.clicked.connect(self._reset_view)
        tb_layout.addWidget(self.btn_reset)

        self.btn_max = QPushButton('□')
        self.btn_max.setStyleSheet(btn_style)
        self.btn_max.clicked.connect(self._toggle_maximize)
        tb_layout.addWidget(self.btn_max)

        self.btn_min = QPushButton('−')
        self.btn_min.setStyleSheet(btn_style)
        self.btn_min.clicked.connect(self.showMinimized)
        tb_layout.addWidget(self.btn_min)

        self.btn_close = QPushButton('✕')
        self.btn_close.setStyleSheet(
            btn_style.replace('#3a3a3a', '#822').replace('#555', '#a44'))
        self.btn_close.clicked.connect(self.close)
        tb_layout.addWidget(self.btn_close)

        layout.addWidget(title_bar)

        # ── Image area ─────────────────────────────────────────────
        self.scroll = _PannableScrollArea()
        self.scroll.setStyleSheet(
            'QScrollArea { border: none; background: #1a1a1a; }')
        self.scroll.setWidgetResizable(False)
        self.scroll.setAlignment(Qt.AlignCenter)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.scroll.setWidget(self.image_label)

        self._load_image()
        layout.addWidget(self.scroll, stretch=1)

        # ── Hint bar ───────────────────────────────────────────────
        hint = QLabel('滚轮缩放 | 图片上拖拽平移 | 边缘拖拽调整窗口')
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet('color: #555; font-size: 11px; padding: 3px; '
                           'background: #1a1a1a;')
        layout.addWidget(hint)

        self.resize(900, 700)
        self._center_on_parent()

    # =================================================================
    # Image loading
    # =================================================================

    def _load_image(self):
        if not os.path.exists(self._image_path):
            self.image_label.setText('图片不存在')
            self.image_label.setStyleSheet('color: #999; font-size: 16px;')
            return

        pix = QPixmap(self._image_path)
        if pix.isNull():
            self.image_label.setText('无法加载图片')
            self.image_label.setStyleSheet('color: #999; font-size: 16px;')
            return

        if self._scale <= 0:
            win_w = max(self.width(), 400)
            target_w = max(400, win_w - 200)
            self._scale = max(0.1, min(2.0, target_w / pix.width()))

        if self._rotation != 0:
            from PySide6.QtGui import QTransform
            pix = pix.transformed(QTransform().rotate(self._rotation),
                                  Qt.SmoothTransformation)
        if self._scale != 1.0:
            w = int(pix.width() * self._scale)
            h = int(pix.height() * self._scale)
            pix = pix.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)

        self.image_label.setPixmap(pix)
        self.image_label.setFixedSize(pix.size())

    def _rotate(self, degrees):
        self._rotation = (self._rotation + degrees) % 360
        self._load_image()

    def _reset_view(self):
        self._rotation = 0
        self._scale = 0.0
        self._load_image()

    def _toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
            self.btn_max.setText('□')
        else:
            self.showMaximized()
            self.btn_max.setText('❐')

    def _center_on_parent(self):
        p = self.parent()
        if p:
            pr = p.geometry()
            fr = self.frameGeometry()
            fr.moveCenter(pr.center())
            # Offset downward to avoid overlapping the main window toolbar
            self.move(fr.topLeft().x(), fr.topLeft().y() + 50)

    # =================================================================
    # Edge resize
    # =================================================================

    def _get_edge(self, pos):
        """Return which edge/corner the position is near, or None."""
        x, y = pos.x(), pos.y()
        w, h = self.width(), self.height()
        m = self._EDGE_MARGIN

        top = y <= m
        bottom = y >= h - m
        left = x <= m
        right = x >= w - m

        if top and left:
            return 'top-left'
        if top and right:
            return 'top-right'
        if bottom and left:
            return 'bottom-left'
        if bottom and right:
            return 'bottom-right'
        if top:
            return 'top'
        if bottom:
            return 'bottom'
        if left:
            return 'left'
        if right:
            return 'right'
        return None

    def _cursor_for_edge(self, edge):
        if edge in ('top-left', 'bottom-right'):
            return Qt.SizeFDiagCursor
        if edge in ('top-right', 'bottom-left'):
            return Qt.SizeBDiagCursor
        if edge in ('top', 'bottom'):
            return Qt.SizeVerCursor
        if edge in ('left', 'right'):
            return Qt.SizeHorCursor
        return Qt.ArrowCursor

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            edge = self._get_edge(event.pos())
            if edge:
                self._resizing = True
                self._resize_edge = edge
                self._resize_start = event.globalPosition().toPoint()
                self._resize_geom = self.geometry()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._resizing:
            delta = event.globalPosition().toPoint() - self._resize_start
            g = self._resize_geom
            e = self._resize_edge

            x, y, w, h = g.x(), g.y(), g.width(), g.height()

            if 'left' in e:
                nx = x + delta.x()
                nw = w - delta.x()
                if nw >= self._MIN_W:
                    x, w = nx, nw
            elif 'right' in e:
                w = max(self._MIN_W, w + delta.x())

            if 'top' in e:
                ny = y + delta.y()
                nh = h - delta.y()
                if nh >= self._MIN_H:
                    y, h = ny, nh
            elif 'bottom' in e:
                h = max(self._MIN_H, h + delta.y())

            self.setGeometry(x, y, w, h)
            return

        # Update cursor when hovering near edges
        edge = self._get_edge(event.pos())
        self.setCursor(self._cursor_for_edge(edge))
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._resizing = False
        self._resize_edge = None
        super().mouseReleaseEvent(event)

    def leaveEvent(self, event):
        self.setCursor(Qt.ArrowCursor)
        super().leaveEvent(event)

    # =================================================================
    # Keyboard / wheel
    # =================================================================

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

    def mouseDoubleClickEvent(self, event):
        # Only close on double-click if not on title bar
        if event.pos().y() > 42:
            self.close()

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        factor = 1.15 if delta > 0 else 1 / 1.15
        self._scale = max(0.1, min(5.0, self._scale * factor))
        self._load_image()


class UploadDialog(QDialog):
    """Dialog for entering cert details when uploading files."""

    def __init__(self, file_paths, manager, current_category,
                 current_sub_filter=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle('上传材质单')
        self.setMinimumWidth(480)
        self.manager = manager
        self.file_paths = file_paths
        if current_sub_filter and current_sub_filter != '__unclassified__':
            self.current_sub_filter = current_sub_filter
        else:
            self.current_sub_filter = {}
        self._new_category_edit = None

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # File summary
        count_label = QLabel(f'共 {len(file_paths)} 个文件')
        count_label.setStyleSheet('font-weight: bold; font-size: 13px;')
        layout.addWidget(count_label)

        file_list = QListWidget()
        file_list.setMaximumHeight(100)
        for fp in file_paths:
            file_list.addItem(os.path.basename(fp))
        layout.addWidget(file_list)

        # Category
        cat_row = QHBoxLayout()
        cat_row.addWidget(QLabel('材料类别:'))
        self.cat_combo = QComboBox()
        visible = self.manager.categories
        cat_keys = sorted(visible.keys(),
                          key=lambda k: visible[k].get('order', 99))
        for key in cat_keys:
            self.cat_combo.addItem(visible[key]['label'], key)
        if current_category and current_category in visible:
            idx = self.cat_combo.findData(current_category)
            if idx >= 0:
                self.cat_combo.setCurrentIndex(idx)
        self.cat_combo.currentIndexChanged.connect(self._on_category_changed)
        cat_row.addWidget(self.cat_combo)
        layout.addLayout(cat_row)

        new_cat_row = QHBoxLayout()
        new_cat_row.addWidget(QLabel('或新建分类:'))
        self.edit_new_category = QLineEdit()
        self.edit_new_category.setPlaceholderText('输入新分类名称')
        self.edit_new_category.textChanged.connect(self._on_new_cat_text_changed)
        new_cat_row.addWidget(self.edit_new_category)
        layout.addLayout(new_cat_row)

        # Params
        self.params_group = QGroupBox('材质单信息')
        self.params_form = QVBoxLayout(self.params_group)
        self.params_form.setSpacing(6)

        batch_row = QHBoxLayout()
        batch_row.addWidget(QLabel('批次号:'))
        self.edit_batch = QLineEdit()
        self.edit_batch.setPlaceholderText('输入批次号')
        if file_paths:
            self.edit_batch.setText(
                os.path.splitext(os.path.basename(file_paths[0]))[0])
        batch_row.addWidget(self.edit_batch)
        self.params_form.addLayout(batch_row)

        self._param_rows_start = self.params_form.count()
        self.param_fields = {}
        layout.addWidget(self.params_group)

        # Date
        date_row = QHBoxLayout()
        date_row.addWidget(QLabel('日期(年月):'))
        self.edit_date = QLineEdit()
        self.edit_date.setPlaceholderText('YYYYMM，如 202506')
        self.edit_date.setMaxLength(6)
        date_row.addWidget(self.edit_date)
        layout.addLayout(date_row)

        # Notes
        notes_row = QHBoxLayout()
        notes_row.addWidget(QLabel('备注:'))
        self.edit_notes = QLineEdit()
        notes_row.addWidget(self.edit_notes)
        layout.addLayout(notes_row)

        # Buttons
        btn_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._validate_and_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self._build_param_fields()
        self._prefill_params()

    def _on_category_changed(self):
        self._build_param_fields()
        self._prefill_params()

    def _on_new_cat_text_changed(self, text):
        self.cat_combo.setEnabled(not bool(text.strip()))

    def _build_param_fields(self):
        # Clear only dynamic param rows (keep batch row at index 0)
        while self.params_form.count() > self._param_rows_start:
            item = self.params_form.takeAt(self.params_form.count() - 1)
            if item.layout():
                while item.layout().count():
                    w = item.layout().takeAt(0)
                    if w.widget():
                        w.widget().deleteLater()
        self.param_fields.clear()

        cat_key = self.cat_combo.currentData()
        if not cat_key:
            return
        cat_config = self.manager.categories.get(cat_key, {})
        for p_def in cat_config.get('params', []):
            row = QHBoxLayout()
            lbl = QLabel(f'{p_def["label"]}:')
            lbl.setFixedWidth(80)
            row.addWidget(lbl)
            edit = QLineEdit()
            edit.setPlaceholderText(f'输入{p_def["label"]}')
            row.addWidget(edit)
            self.params_form.addLayout(row)
            self.param_fields[p_def['key']] = edit

    def _prefill_params(self):
        for key, val in self.current_sub_filter.items():
            if key in self.param_fields:
                self.param_fields[key].setText(str(val))

    def _validate_and_accept(self):
        new_cat = self.edit_new_category.text().strip()
        cat_key = self.cat_combo.currentData()
        if not new_cat and not cat_key:
            QMessageBox.warning(self, '提示', '请选择材料类别或输入新分类名称。')
            return
        self.accept()

    def get_result(self):
        """Return (category_key, is_new, new_label, batch_number, params,
                   cert_date, notes)."""
        batch = self.edit_batch.text().strip()
        new_cat = self.edit_new_category.text().strip()
        if new_cat:
            return (new_cat, True, new_cat, batch,
                    self._collect_params(),
                    self.edit_date.text().strip(),
                    self.edit_notes.text().strip())
        return (self.cat_combo.currentData(), False, None, batch,
                self._collect_params(),
                self.edit_date.text().strip(),
                self.edit_notes.text().strip())

    def _collect_params(self):
        result = {}
        for key, edit in self.param_fields.items():
            text = edit.text().strip()
            if text:
                result[key] = text
        return result


class ImageManagerWidget(QWidget):
    """Material certificate image management panel."""

    material_changed = Signal()

    def __init__(self, product_dir, product_type=None, parent=None):
        super().__init__(parent)
        self.manager = MaterialManager(product_dir, product_type)
        self.current_category = None
        self._current_sub_filter = None
        self._current_is_default = False
        self._showing_defaults = False
        self._showing_expired = False
        self._saved_category = None
        self.selected_cert_id = None
        self._checked_cert_ids = set()
        self.setAcceptDrops(True)
        self._setup_ui()
        self._load_tree()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Toolbar
        toolbar = QHBoxLayout()
        self.btn_upload = QPushButton('上传材质单')
        self.btn_upload.clicked.connect(self._on_upload)
        toolbar.addWidget(self.btn_upload)

        self.btn_import_library = QPushButton('导入材质单库')
        self.btn_import_library.setObjectName('btnSecondary')
        self.btn_import_library.clicked.connect(self._on_import_library)
        toolbar.addWidget(self.btn_import_library)

        self.lbl_drop_hint = QLabel('支持拖拽上传图片/PDF')
        self.lbl_drop_hint.setStyleSheet('color: #888; font-size: 11px;')
        toolbar.addWidget(self.lbl_drop_hint)

        self.btn_delete = QPushButton('删除选中')
        self.btn_delete.setEnabled(False)
        self.btn_delete.clicked.connect(self._on_delete)
        toolbar.addWidget(self.btn_delete)

        toolbar.addSpacing(16)

        self.btn_select_all = QPushButton('全选')
        self.btn_select_all.clicked.connect(self._on_select_all)
        toolbar.addWidget(self.btn_select_all)

        self.btn_deselect_all = QPushButton('取消全选')
        self.btn_deselect_all.clicked.connect(self._on_deselect_all)
        toolbar.addWidget(self.btn_deselect_all)

        toolbar.addSpacing(12)

        self.btn_show_defaults = QPushButton('默认材质单')
        self.btn_show_defaults.setCheckable(True)
        self.btn_show_defaults.clicked.connect(self._on_show_defaults)
        toolbar.addWidget(self.btn_show_defaults)

        self.lbl_checked_count = QLabel('已选: 0')
        self.lbl_checked_count.setStyleSheet('color: #4CAF50; font-weight: bold;')
        toolbar.addWidget(self.lbl_checked_count)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Splitter: category list | image grid + detail editor
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(8)

        # Left: category tree
        self.category_tree = QTreeWidget()
        self.category_tree.setMinimumWidth(120)
        self.category_tree.setHeaderHidden(True)
        self.category_tree.setIndentation(14)
        self.category_tree.setStyleSheet(
            'QTreeWidget { border: 1px solid #d0d0d0; border-radius: 4px; }'
            'QTreeWidget::item { padding: 3px 6px; }'
            'QTreeWidget::item:selected { background: #4A90D9; color: white; }'
        )
        self.category_tree.currentItemChanged.connect(self._on_node_changed)
        splitter.addWidget(self.category_tree)

        # Right panel
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        # Stacked image grids: page 0 = normal, page 1 = defaults
        self.grid_stack = QStackedWidget()

        # Page 0: Normal category grid
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet(
            'QScrollArea { border: 1px solid #d0d0d0; border-radius: 4px; }')
        self.grid_widget = QWidget()
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setContentsMargins(8, 8, 8, 8)
        self.grid_layout.setSpacing(8)
        self.grid_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.scroll_area.setWidget(self.grid_widget)
        self.grid_stack.addWidget(self.scroll_area)

        # Page 1: Defaults grid
        self.defaults_scroll_area = QScrollArea()
        self.defaults_scroll_area.setWidgetResizable(True)
        self.defaults_scroll_area.setStyleSheet(
            'QScrollArea { border: 1px solid #d0d0d0; border-radius: 4px; }')
        self.defaults_grid_widget = QWidget()
        self.defaults_grid_layout = QGridLayout(self.defaults_grid_widget)
        self.defaults_grid_layout.setContentsMargins(8, 8, 8, 8)
        self.defaults_grid_layout.setSpacing(8)
        self.defaults_grid_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.defaults_scroll_area.setWidget(self.defaults_grid_widget)
        self.grid_stack.addWidget(self.defaults_scroll_area)

        right_layout.addWidget(self.grid_stack, stretch=2)

        # Detail editor
        self._create_detail_editor(right_layout)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        layout.addWidget(splitter, stretch=1)

    def _create_detail_editor(self, parent_layout):
        group = QGroupBox('详情编辑')
        group.setStyleSheet(
            'QGroupBox { font-weight: bold; border: 1px solid #d0d0d0; '
            'border-radius: 4px; margin-top: 8px; padding-top: 16px; }')
        form = QVBoxLayout(group)
        form.setSpacing(6)

        # Category
        row = QHBoxLayout()
        row.addWidget(QLabel('分类:'))
        self.combo_detail_category = QComboBox()
        for key in sorted(self.manager.categories.keys(),
                          key=lambda k: self.manager.categories[k].get('order', 99)):
            cfg = self.manager.categories[key]
            self.combo_detail_category.addItem(cfg.get('label', key), key)
        self.combo_detail_category.currentIndexChanged.connect(
            self._on_detail_category_changed)
        row.addWidget(self.combo_detail_category)
        form.addLayout(row)

        # Batch number
        row = QHBoxLayout()
        row.addWidget(QLabel('批次号:'))
        self.edit_batch = QLineEdit()
        self.edit_batch.setPlaceholderText('输入批次号')
        row.addWidget(self.edit_batch)
        form.addLayout(row)

        # Dynamic param fields
        self.params_container = QWidget()
        self.params_layout = QVBoxLayout(self.params_container)
        self.params_layout.setContentsMargins(0, 0, 0, 0)
        self.params_layout.setSpacing(6)
        self.param_fields = {}
        form.addWidget(self.params_container)

        # Date + expired
        row = QHBoxLayout()
        row.addWidget(QLabel('日期(年月):'))
        self.edit_date = QLineEdit()
        self.edit_date.setPlaceholderText('YYYYMM，如 202506')
        self.edit_date.setMaxLength(6)
        row.addWidget(self.edit_date)
        self.chk_expired = QCheckBox('已过期')
        self.chk_expired.setToolTip('无日期时可手动勾选')
        row.addWidget(self.chk_expired)
        form.addLayout(row)

        # Rotation
        row = QHBoxLayout()
        row.addWidget(QLabel('图片方向:'))
        self.combo_rotation = QComboBox()
        self.combo_rotation.addItem('不旋转', 0)
        self.combo_rotation.addItem('右转90°', 90)
        self.combo_rotation.addItem('转180°', 180)
        self.combo_rotation.addItem('左转90°', 270)
        row.addWidget(self.combo_rotation)
        form.addLayout(row)

        # Notes
        row = QHBoxLayout()
        row.addWidget(QLabel('备注:'))
        self.edit_notes = QLineEdit()
        row.addWidget(self.edit_notes)
        form.addLayout(row)

        # Buttons
        btn_row = QHBoxLayout()
        self.btn_toggle_default = QPushButton('设为默认')
        self.btn_toggle_default.setObjectName('btnSecondary')
        self.btn_toggle_default.setEnabled(False)
        self.btn_toggle_default.clicked.connect(self._on_toggle_default)
        btn_row.addWidget(self.btn_toggle_default)
        btn_row.addStretch()
        self.btn_save = QPushButton('保存修改')
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(self.btn_save)
        form.addLayout(btn_row)

        parent_layout.addWidget(group)

    # =================================================================
    # Data loading
    # =================================================================

    def _load_tree(self):
        """Build the category→subcategory tree."""
        self.category_tree.blockSignals(True)
        self.category_tree.clear()

        cats = self.manager.get_all_categories_with_counts(exclude_expired=True)

        for cat in cats:
            cat_item = QTreeWidgetItem()
            cat_item.setText(0, f'{cat["label"]}  ({cat["count"]})')
            cat_item.setData(0, Qt.UserRole, cat['key'])
            cat_item.setData(0, Qt.UserRole + 1, 'category')
            self.category_tree.addTopLevelItem(cat_item)

            # Subcategories
            subs = self.manager.get_certs_grouped_by_subcategory(
                cat['key'], exclude_expired=True)
            if subs:
                for sub in subs:
                    sub_item = QTreeWidgetItem()
                    sub_item.setText(0, f'{sub["sub_label"]}  ({sub["count"]})')
                    sub_item.setData(0, Qt.UserRole, cat['key'])
                    sub_item.setData(0, Qt.UserRole + 1, 'subcategory')
                    sub_item.setData(0, Qt.UserRole + 2, sub.get('param_values', {}))
                    if sub['sub_label'] == '未分类':
                        sub_item.setData(0, Qt.UserRole + 3, True)
                    cat_item.addChild(sub_item)

            # Restore previous selection
            if (not self._showing_expired and not self._showing_defaults
                    and cat['key'] == self.current_category):
                self.category_tree.setCurrentItem(cat_item)
                cat_item.setExpanded(True)

        # Expired node
        expired_count = self.manager.get_expired_count()
        expired_item = QTreeWidgetItem()
        expired_item.setText(0, f'过期材质单  ({expired_count})')
        expired_item.setData(0, Qt.UserRole, '__expired__')
        expired_item.setData(0, Qt.UserRole + 1, 'expired')
        expired_item.setForeground(0, self.category_tree.palette().color(
            QPalette.Disabled, QPalette.Text))
        self.category_tree.addTopLevelItem(expired_item)
        if self._showing_expired:
            self.category_tree.setCurrentItem(expired_item)

        self.category_tree.blockSignals(False)

        if (not self.category_tree.currentItem()
                and self.category_tree.topLevelItemCount() > 0):
            first = self.category_tree.topLevelItem(0)
            first.setExpanded(True)
            self.category_tree.setCurrentItem(first)

    def _refresh_category_node(self, category_key):
        """Update a single category node's count and subcategories."""
        cat_config = self.manager.categories.get(category_key, {})
        label = cat_config.get('label', category_key)
        count = self.manager.get_category_count(category_key,
                                                 exclude_expired=True)

        # Find the top-level item
        target_item = None
        for i in range(self.category_tree.topLevelItemCount()):
            item = self.category_tree.topLevelItem(i)
            if (item.data(0, Qt.UserRole) == category_key
                    and item.data(0, Qt.UserRole + 1) == 'category'):
                target_item = item
                break
        if target_item is None:
            return

        self.category_tree.blockSignals(True)

        target_item.setText(0, f'{label}  ({count})')

        # Remove old children
        while target_item.childCount():
            target_item.removeChild(target_item.child(0))

        # Re-add subcategories
        subs = self.manager.get_certs_grouped_by_subcategory(
            category_key, exclude_expired=True)
        if subs:
            for sub in subs:
                sub_item = QTreeWidgetItem()
                sub_item.setText(0, f'{sub["sub_label"]}  ({sub["count"]})')
                sub_item.setData(0, Qt.UserRole, category_key)
                sub_item.setData(0, Qt.UserRole + 1, 'subcategory')
                sub_item.setData(0, Qt.UserRole + 2,
                                 sub.get('param_values', {}))
                if sub['sub_label'] == '未分类':
                    sub_item.setData(0, Qt.UserRole + 3, True)
                target_item.addChild(sub_item)

        # Restore selection
        target_item.setExpanded(True)
        if self.current_category == category_key:
            if self._current_sub_filter:
                found = False
                if self._current_sub_filter == '__unclassified__':
                    for i in range(target_item.childCount()):
                        child = target_item.child(i)
                        if child.data(0, Qt.UserRole + 3):
                            self.category_tree.setCurrentItem(child)
                            found = True
                            break
                else:
                    for i in range(target_item.childCount()):
                        child = target_item.child(i)
                        cf = child.data(0, Qt.UserRole + 2) or {}
                        if cf == self._current_sub_filter:
                            self.category_tree.setCurrentItem(child)
                            found = True
                            break
                if not found:
                    self._current_sub_filter = None
                    self.category_tree.setCurrentItem(target_item)
            else:
                self.category_tree.setCurrentItem(target_item)

        self._refresh_expired_node()
        self.category_tree.blockSignals(False)

    def _refresh_expired_node(self):
        """Update the expired node count at the bottom of the tree."""
        expired_count = self.manager.get_expired_count()
        # Find existing expired node
        for i in range(self.category_tree.topLevelItemCount()):
            item = self.category_tree.topLevelItem(i)
            if item.data(0, Qt.UserRole + 1) == 'expired':
                item.setText(0, f'过期材质单  ({expired_count})')
                return

    def _on_node_changed(self, current, previous):
        """Handle category, subcategory, or expired node selection."""
        if not current:
            return

        node_type = current.data(0, Qt.UserRole + 1)
        cat_key = current.data(0, Qt.UserRole)

        if node_type == 'expired':
            self._showing_expired = True
            self._showing_defaults = False
            self.btn_show_defaults.setChecked(False)
            self.btn_show_defaults.setText('默认材质单')
            self.btn_show_defaults.setStyleSheet('')
            self.current_category = None
            self._current_sub_filter = None
            self.selected_cert_id = None
            self._clear_detail()
            self._create_param_fields()
            self._load_images()
        elif node_type == 'category':
            self._showing_expired = False
            # Expand to show subcategories
            if current.childCount() > 0:
                current.setExpanded(True)
            old_cat = self.current_category
            self.current_category = cat_key
            self._current_sub_filter = None
            if old_cat != self.current_category:
                self.selected_cert_id = None
                self._clear_detail()
            self._create_param_fields(self.current_category)
            self._load_images()
        elif node_type == 'subcategory':
            self._showing_expired = False
            self.current_category = cat_key
            is_unclassified = bool(current.data(0, Qt.UserRole + 3))
            if is_unclassified:
                self._current_sub_filter = '__unclassified__'
            else:
                self._current_sub_filter = current.data(0, Qt.UserRole + 2) or {}
            self.selected_cert_id = None
            self._clear_detail()
            self._create_param_fields(self.current_category)
            # Pre-fill param fields from sub_filter
            if not is_unclassified:
                for key, val in self._current_sub_filter.items():
                    if key in self.param_fields:
                        self.param_fields[key].setText(str(val))
            self._load_images(sub_filter=self._current_sub_filter)

        self.btn_delete.setEnabled(False)
        self.btn_save.setEnabled(False)

    def _detail_category(self):
        if hasattr(self, 'combo_detail_category'):
            cat = self.combo_detail_category.currentData()
            if cat:
                return cat
        return self.current_category

    def _set_detail_category(self, category_key):
        if not hasattr(self, 'combo_detail_category') or not category_key:
            return
        idx = self.combo_detail_category.findData(category_key)
        if idx >= 0:
            was_blocked = self.combo_detail_category.blockSignals(True)
            self.combo_detail_category.setCurrentIndex(idx)
            self.combo_detail_category.blockSignals(was_blocked)

    def _on_detail_category_changed(self):
        old_params = {}
        for key, edit in self.param_fields.items():
            text = edit.text().strip()
            if text:
                old_params[key] = text
        self._create_param_fields(self._detail_category())
        for key, value in old_params.items():
            if key in self.param_fields:
                self.param_fields[key].setText(value)

    def _create_param_fields(self, category_key=None):
        # Clear existing
        for i in reversed(range(self.params_layout.count())):
            item = self.params_layout.itemAt(i)
            if item.layout():
                while item.layout().count():
                    w = item.layout().takeAt(0)
                    if w.widget():
                        w.widget().deleteLater()
                self.params_layout.removeItem(item)
        self.param_fields.clear()

        category_key = category_key if category_key is not None else self.current_category
        if not category_key:
            return

        self._set_detail_category(category_key)
        cat_config = self.manager.categories.get(category_key, {})
        for p_def in cat_config.get('params', []):
            row = QHBoxLayout()
            lbl = QLabel(f'{p_def["label"]}:')
            lbl.setFixedWidth(80)
            row.addWidget(lbl)
            edit = QLineEdit()
            edit.setPlaceholderText(f'输入{p_def["label"]}')
            row.addWidget(edit)
            self.params_layout.addLayout(row)
            self.param_fields[p_def['key']] = edit

    def _load_images(self, sub_filter=None):
        if self._showing_expired:
            # Clear defaults grid first (not using it)
            while self.defaults_grid_layout.count():
                item = self.defaults_grid_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            # Clear normal grid
            while self.grid_layout.count():
                item = self.grid_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            certs = self.manager.get_expired_certificates()
            if not certs:
                placeholder = QLabel('暂无过期材质单')
                placeholder.setAlignment(Qt.AlignCenter)
                placeholder.setStyleSheet('color: #999; font-size: 14px;')
                self.grid_layout.addWidget(placeholder, 0, 0)
            else:
                cols = 3
                for idx, cert in enumerate(certs):
                    cat_config = self.manager.categories.get(
                        cert['category'], {})
                    cert['_cat_config'] = cat_config
                    img_path = self.manager.get_image_path(cert)
                    card = ImageCard(cert, img_path, manager=self.manager)
                    card.setImagePath(img_path or '')
                    if cert['id'] in self._checked_cert_ids:
                        card.setChecked(True)
                    card.clicked.connect(self._on_card_clicked)
                    card.double_clicked.connect(self._on_card_double_clicked)
                    card.checked_changed.connect(self._on_card_checked)
                    r, c = divmod(idx, cols)
                    self.grid_layout.addWidget(card, r, c)
            self.grid_stack.setCurrentIndex(0)
            return

        if self._showing_defaults:
            # Clear defaults grid
            while self.defaults_grid_layout.count():
                item = self.defaults_grid_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            certs = self.manager.get_default_certificates()
            if not certs:
                placeholder = QLabel('暂无默认材质单')
                placeholder.setAlignment(Qt.AlignCenter)
                placeholder.setStyleSheet('color: #999; font-size: 14px;')
                self.defaults_grid_layout.addWidget(placeholder, 0, 0)
            else:
                cols = 3
                for idx, cert in enumerate(certs):
                    cat_config = self.manager.categories.get(
                        cert['category'], {})
                    cert['_cat_config'] = cat_config
                    img_path = self.manager.get_image_path(cert)
                    card = ImageCard(cert, img_path, manager=self.manager)
                    card.setImagePath(img_path or '')
                    if cert['id'] in self._checked_cert_ids:
                        card.setChecked(True)
                    card.clicked.connect(self._on_card_clicked)
                    card.double_clicked.connect(self._on_card_double_clicked)
                    card.checked_changed.connect(self._on_card_checked)
                    r, c = divmod(idx, cols)
                    self.defaults_grid_layout.addWidget(card, r, c)
            self.grid_stack.setCurrentIndex(1)
            return

        # Normal mode: clear normal grid
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.current_category:
            return

        certs = self.manager.get_certificates_by_category(
            self.current_category, exclude_expired=True)
        cat_config = self.manager.categories.get(self.current_category, {})

        # Apply subcategory filter
        if sub_filter and sub_filter != '__unclassified__':
            filtered = []
            for cert in certs:
                params = cert.get('params', {})
                match = True
                for key, val in sub_filter.items():
                    if str(params.get(key, '')).strip() != str(val).strip():
                        match = False
                        break
                if match:
                    filtered.append(cert)
            certs = filtered
        elif sub_filter == '__unclassified__':
            group_by = cat_config.get('group_by', [])
            certs = [c for c in certs
                     if not all(str(c.get('params', {}).get(k, '')).strip()
                                for k in group_by)]

        cols = 3
        for idx, cert in enumerate(certs):
            cert['_cat_config'] = cat_config
            img_path = self.manager.get_image_path(cert)
            card = ImageCard(cert, img_path, manager=self.manager)
            card.setImagePath(img_path or '')
            # Restore checked state
            if cert['id'] in self._checked_cert_ids:
                card.setChecked(True)
            card.clicked.connect(self._on_card_clicked)
            card.double_clicked.connect(self._on_card_double_clicked)
            card.checked_changed.connect(self._on_card_checked)
            r, c = divmod(idx, cols)
            self.grid_layout.addWidget(card, r, c)
        self.grid_stack.setCurrentIndex(0)

    # =================================================================
    # Detail editor
    # =================================================================

    def _on_card_clicked(self, cert_id):
        self.selected_cert_id = cert_id
        cert = self.manager.get_certificate(cert_id)
        if cert:
            self._populate_detail(cert)
            self.btn_save.setEnabled(True)
            self.btn_delete.setEnabled(True)
        # Highlight selected card in whichever grid is visible
        target_layout = (self.defaults_grid_layout if self._showing_defaults
                         else self.grid_layout)
        for i in range(target_layout.count()):
            w = target_layout.itemAt(i).widget()
            if isinstance(w, ImageCard):
                w.setSelected(w.cert_id == cert_id)

    def _on_card_double_clicked(self, cert_id, image_path):
        if not image_path or not os.path.exists(image_path):
            QMessageBox.warning(self, '提示', '该材质单没有可查看的图片。')
            return
        cert = self.manager.get_certificate(cert_id)
        info_parts = []
        if cert:
            info_parts.append(cert.get('category', ''))
            info_parts.append(f'批次号: {cert.get("batch_number", "-")}')
            params = cert.get('params', {})
            for k, v in params.items():
                if v:
                    info_parts.append(f'{k}={v}')
        rotation = int(cert.get('rotation') or 0) if cert else 0
        viewer = ImageViewerDialog(
            image_path, '  |  '.join(info_parts), self,
            initial_rotation=rotation)
        viewer.show()

    def _on_card_checked(self, cert_id, checked):
        if checked:
            self._checked_cert_ids.add(cert_id)
        else:
            self._checked_cert_ids.discard(cert_id)
        self._update_checked_count()

    def _on_select_all(self):
        target_layout = (self.defaults_grid_layout if self._showing_defaults
                         else self.grid_layout)
        for i in range(target_layout.count()):
            w = target_layout.itemAt(i).widget()
            if isinstance(w, ImageCard):
                w.setChecked(True)
                self._checked_cert_ids.add(w.cert_id)
        self._update_checked_count()

    def _on_deselect_all(self):
        target_layout = (self.defaults_grid_layout if self._showing_defaults
                         else self.grid_layout)
        for i in range(target_layout.count()):
            w = target_layout.itemAt(i).widget()
            if isinstance(w, ImageCard):
                w.setChecked(False)
        self._checked_cert_ids.clear()
        self._update_checked_count()

    def _update_checked_count(self):
        self.lbl_checked_count.setText(f'已选: {len(self._checked_cert_ids)}')

    def get_checked_cert_ids(self):
        """Return sorted list of checked certificate IDs."""
        return sorted(self._checked_cert_ids)

    def _populate_detail(self, cert):
        self._set_detail_category(cert.get('category', ''))
        self._create_param_fields(cert.get('category', ''))
        self.edit_batch.setText(cert.get('batch_number', ''))
        self.edit_notes.setText(cert.get('notes', ''))
        rotation = int(cert.get('rotation') or 0) % 360
        idx = self.combo_rotation.findData(rotation)
        self.combo_rotation.setCurrentIndex(idx if idx >= 0 else 0)

        params = cert.get('params', {})
        for key, edit in self.param_fields.items():
            edit.setText(str(params.get(key, '')))

        cert_date = cert.get('cert_date', '')
        self.edit_date.setText(cert_date)
        if cert_date:
            expired = bool(self.manager._calc_expired(cert_date, False))
            self.chk_expired.setChecked(expired)
            self.chk_expired.setEnabled(False)
        else:
            self.chk_expired.setChecked(bool(cert.get('is_expired', False)))
            self.chk_expired.setEnabled(True)

        self._current_is_default = self.manager.is_default_for(cert)
        self.btn_toggle_default.setEnabled(True)
        if self._current_is_default:
            self.btn_toggle_default.setText('取消默认')
            self.btn_toggle_default.setStyleSheet(
                'QPushButton { background-color: #FFF8DC; color: #B8860B; '
                'border: 1px solid #DAA520; }'
                'QPushButton:hover { background-color: #FFE4B5; }')
        else:
            self.btn_toggle_default.setText('设为默认')
            self.btn_toggle_default.setStyleSheet('')

    def _clear_detail(self):
        self._set_detail_category(self.current_category)
        self.edit_batch.clear()
        self.edit_notes.clear()
        for edit in self.param_fields.values():
            edit.clear()
        self.edit_date.clear()
        self.combo_rotation.setCurrentIndex(0)
        self.chk_expired.setChecked(False)
        self.chk_expired.setEnabled(True)
        self._current_is_default = False
        self.btn_toggle_default.setEnabled(False)
        self.btn_toggle_default.setText('设为默认')
        self.btn_toggle_default.setStyleSheet('')

    def _collect_detail(self):
        params = {}
        for key, edit in self.param_fields.items():
            text = edit.text().strip()
            if text:
                params[key] = text
        cert_date = self.edit_date.text().strip()
        is_expired = 1 if self.chk_expired.isChecked() else 0
        return {
            'category': self._detail_category(),
            'batch_number': self.edit_batch.text().strip(),
            'params': params,
            'notes': self.edit_notes.text().strip(),
            'cert_date': cert_date,
            'is_expired': is_expired,
            'rotation': self.combo_rotation.currentData() or 0,
        }

    # =================================================================
    # Event handlers
    # =================================================================

    def _import_files(self, file_paths, category, params, cert_date, notes,
                      batch_number=''):
        """Import image or PDF files. PDF pages are each added as a cert.
        Returns list of new certificate IDs."""
        new_ids = []
        used_batch = False  # only override first non-PDF file's batch

        for f in file_paths:
            ext = os.path.splitext(f)[1].lower()
            if ext == '.pdf':
                with tempfile.TemporaryDirectory() as tmpdir:
                    pages = self.manager.convert_pdf_to_images(f, tmpdir)
                    base = os.path.splitext(os.path.basename(f))[0]
                    for page_num, img_path in pages:
                        cert_id = self.manager.add_certificate(
                            category=category,
                            batch_number=f'{base}_p{page_num}',
                            params=params,
                            source_path=img_path,
                            notes=notes,
                            cert_date=cert_date,
                        )
                        new_ids.append(cert_id)
            else:
                if batch_number and not used_batch:
                    base = batch_number
                    used_batch = True
                else:
                    base = os.path.splitext(os.path.basename(f))[0]
                cert_id = self.manager.add_certificate(
                    category=category,
                    batch_number=base,
                    params=params,
                    source_path=f,
                    notes=notes,
                    cert_date=cert_date,
                )
                new_ids.append(cert_id)
        return new_ids

    def _on_upload(self):
        if self._showing_expired:
            QMessageBox.warning(self, '提示', '过期材质单视图为只读，请先切换到材料类别。')
            return

        files, _ = QFileDialog.getOpenFileNames(
            self, '选择材质单', '',
            '图片和PDF文件 (*.png *.jpg *.jpeg *.pdf);;'
            '图片文件 (*.png *.jpg *.jpeg);;'
            'PDF文件 (*.pdf)')
        if not files:
            return

        self._show_upload_dialog(files)

    def _on_import_library(self):
        folder = QFileDialog.getExistingDirectory(
            self, '选择旧材质单库文件夹', '')
        if not folder:
            return
        try:
            stats = self.manager.import_library(folder)
            self.selected_cert_id = None
            self._checked_cert_ids.clear()
            self._clear_detail()
            self._load_tree()
            self._load_images(sub_filter=self._current_sub_filter)
            self.material_changed.emit()
            message = (
                '导入完成：\n'
                f'成功导入 {stats["imported"]} 条\n'
                f'跳过 {stats["skipped"]} 条\n'
                f'新增分类 {stats["categories_added"]} 个'
            )
            if stats.get('errors'):
                preview = '\n'.join(stats['errors'][:5])
                more = '' if len(stats['errors']) <= 5 else '\n……'
                message += f'\n\n部分错误：\n{preview}{more}'
            QMessageBox.information(self, '导入材质单库', message)
        except Exception as e:
            QMessageBox.warning(self, '导入失败', str(e))

    def _show_upload_dialog(self, files):
        """Show upload dialog and process files if accepted."""
        dlg = UploadDialog(files, self.manager, self.current_category,
                           self._current_sub_filter, self)
        if dlg.exec() != QDialog.Accepted:
            return

        (cat_key, is_new, new_label, batch_number,
         params, cert_date, notes) = dlg.get_result()
        if is_new:
            self.manager.add_category(cat_key, new_label)
            self._add_category_to_tree(cat_key, new_label)

        new_ids = self._import_files(files, cat_key, params, cert_date, notes,
                                     batch_number)

        was_category_change = (cat_key != self.current_category)
        self.current_category = cat_key
        self._current_sub_filter = None
        self.selected_cert_id = None
        self._clear_detail()
        self._create_param_fields(cat_key)

        if was_category_change or is_new:
            # Switching to a different category — full reload needed
            self._refresh_category_node(cat_key)
            self._load_images()
        else:
            # Same category — incrementally add new cards and refresh count
            self._refresh_category_node(cat_key)
            cat_config = self.manager.categories.get(cat_key, {})
            new_certs = [self.manager.get_certificate(cid) for cid in new_ids]
            new_certs = [c for c in new_certs if c is not None]
            for cert in new_certs:
                cert['_cat_config'] = cat_config
                img_path = self.manager.get_image_path(cert)
                card = ImageCard(cert, img_path, manager=self.manager)
                card.setImagePath(img_path or '')
                card.clicked.connect(self._on_card_clicked)
                card.double_clicked.connect(self._on_card_double_clicked)
                card.checked_changed.connect(self._on_card_checked)
                # Place at the first available position
                idx = self.grid_layout.count()
                # Remove placeholder label if present
                if idx == 1:
                    w = self.grid_layout.itemAt(0).widget()
                    if isinstance(w, QLabel) and not isinstance(w, ImageCard):
                        w.deleteLater()
                        idx = 0
                r, c = divmod(idx, 3)
                self.grid_layout.addWidget(card, r, c)

        self.material_changed.emit()

    def _add_category_to_tree(self, cat_key, label):
        """Add a single new category node to the tree without full rebuild."""
        cat_config = self.manager.categories.get(cat_key, {})
        order = cat_config.get('order', 99)
        count = 0

        if hasattr(self, 'combo_detail_category') \
                and self.combo_detail_category.findData(cat_key) < 0:
            self.combo_detail_category.addItem(label, cat_key)

        cat_item = QTreeWidgetItem()
        cat_item.setText(0, f'{label}  ({count})')
        cat_item.setData(0, Qt.UserRole, cat_key)
        cat_item.setData(0, Qt.UserRole + 1, 'category')

        # Insert in order position before the expired node
        insert_idx = self.category_tree.topLevelItemCount() - 1  # before expired
        for i in range(self.category_tree.topLevelItemCount()):
            item = self.category_tree.topLevelItem(i)
            if item.data(0, Qt.UserRole + 1) == 'expired':
                insert_idx = i
                break
            item_order = self.manager.categories.get(
                item.data(0, Qt.UserRole), {}).get('order', 99)
            if order < item_order:
                insert_idx = i
                break

        self.category_tree.insertTopLevelItem(insert_idx, cat_item)
        self.category_tree.setCurrentItem(cat_item)

    def _on_show_defaults(self):
        if self.btn_show_defaults.isChecked():
            self._showing_defaults = True
            self._showing_expired = False
            self._saved_category = self.current_category
            self._saved_sub_filter = self._current_sub_filter
            self.current_category = None
            self._current_sub_filter = None
            self.selected_cert_id = None
            self._clear_detail()
            self._create_param_fields()
            self.btn_show_defaults.setText('返回分类视图')
            self.btn_show_defaults.setStyleSheet(
                'QPushButton { background-color: #FFF8DC; color: #B8860B; '
                'border: 1px solid #DAA520; }'
                'QPushButton:hover { background-color: #FFE4B5; }')
            self.category_tree.blockSignals(True)
            self._load_images()
        else:
            self._showing_defaults = False
            self.current_category = self._saved_category
            self._current_sub_filter = self._saved_sub_filter
            self._saved_category = None
            self._saved_sub_filter = None
            self.selected_cert_id = None
            self._clear_detail()
            self.btn_show_defaults.setText('默认材质单')
            self.btn_show_defaults.setStyleSheet('')
            self.category_tree.blockSignals(False)
            self.grid_stack.setCurrentIndex(0)

    def _on_toggle_default(self):
        if self.selected_cert_id is None:
            return
        new_val = self.manager.toggle_default(self.selected_cert_id)
        if new_val is None:
            return
        # Re-read cert to check default status for THIS product
        cert = self.manager.get_certificate(self.selected_cert_id)
        self._current_is_default = self.manager.is_default_for(cert) if cert else False
        if self._current_is_default:
            self.btn_toggle_default.setText('取消默认')
            self.btn_toggle_default.setStyleSheet(
                'QPushButton { background-color: #FFF8DC; color: #B8860B; '
                'border: 1px solid #DAA520; }'
                'QPushButton:hover { background-color: #FFE4B5; }')
        else:
            self.btn_toggle_default.setText('设为默认')
            self.btn_toggle_default.setStyleSheet('')
        # Update the card badge directly instead of reloading all images
        target_layout = (self.defaults_grid_layout if self._showing_defaults
                         else self.grid_layout)
        for i in range(target_layout.count()):
            w = target_layout.itemAt(i).widget()
            if isinstance(w, ImageCard) and w.cert_id == self.selected_cert_id:
                w._is_default = self._current_is_default
                w._default_badge.setVisible(self._current_is_default)
                break
        # If viewing defaults, reload to reflect additions/removals
        if self._showing_defaults:
            self._load_images()

    def _on_delete(self):
        if self.selected_cert_id is None:
            return
        reply = QMessageBox.question(
            self, '确认删除', '确定要删除选中的材质单图片吗？此操作不可恢复。',
            QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.manager.delete_certificate(self.selected_cert_id)
            self.selected_cert_id = None
            self.btn_save.setEnabled(False)
            self.btn_delete.setEnabled(False)
            self._clear_detail()
            self._refresh_category_node(self.current_category)
            self._load_images(sub_filter=self._current_sub_filter)
            self.material_changed.emit()

    def _on_save(self):
        if self.selected_cert_id is None:
            return

        old_cert = self.manager.get_certificate(self.selected_cert_id)
        old_category = old_cert.get('category', '') if old_cert else ''
        data = self._collect_detail()
        if not data['batch_number']:
            QMessageBox.warning(self, '提示', '批次号不能为空。')
            return

        self.manager.update_certificate(
            self.selected_cert_id,
            category=data['category'],
            batch_number=data['batch_number'],
            params=data['params'],
            notes=data['notes'],
            cert_date=data['cert_date'],
            is_expired=data['is_expired'],
            rotation=data['rotation'],
        )
        new_category = data.get('category') or old_category
        if old_category and new_category != old_category:
            self.current_category = new_category
            self._current_sub_filter = None
            self._load_tree()
            self._load_images()
        elif self.current_category:
            old_filter = self._current_sub_filter
            self._refresh_category_node(self.current_category)
            if old_filter and not self._current_sub_filter:
                self._load_images()
        else:
            self._refresh_expired_node()
        self._refresh_card(self.selected_cert_id)
        self.material_changed.emit()

    def _refresh_card(self, cert_id):
        """Update a single card's labels in-place instead of reloading grid."""
        layout = (self.defaults_grid_layout if self._showing_defaults
                  else self.grid_layout)
        cert = self.manager.get_certificate(cert_id)
        if not cert:
            return
        if self._showing_expired and not self.manager._is_cert_expired(
                cert.get('cert_date', ''), cert.get('is_expired', 0)):
            self._load_images()
            return
        if not self._showing_expired and not self._showing_defaults \
                and self.manager._is_cert_expired(
                    cert.get('cert_date', ''), cert.get('is_expired', 0)):
            self._load_images(sub_filter=self._current_sub_filter)
            return
        if self._showing_defaults and not cert.get('is_default'):
            self._load_images()
            return
        if self._current_sub_filter:
            if self._current_sub_filter == '__unclassified__':
                cat_config = self.manager.categories.get(
                    cert.get('category', ''), {})
                group_by = cat_config.get('group_by', [])
                still_unclassified = not all(
                    str(cert.get('params', {}).get(k, '')).strip()
                    for k in group_by)
                if not still_unclassified:
                    self._load_images(sub_filter=self._current_sub_filter)
                    return
            else:
                params = cert.get('params', {})
                for key, val in self._current_sub_filter.items():
                    if str(params.get(key, '')).strip() != str(val).strip():
                        self._load_images(sub_filter=self._current_sub_filter)
                        return
        cat_config = self.manager.categories.get(cert.get('category', ''), {})
        cert['_cat_config'] = cat_config
        img_path = self.manager.get_image_path(cert)
        for i in range(layout.count()):
            w = layout.itemAt(i).widget()
            if isinstance(w, ImageCard) and w.cert_id == cert_id:
                w.update_info(cert, img_path or '')
                break

    # =================================================================
    # Drag and drop
    # =================================================================

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                ext = url.toLocalFile().lower()
                if ext.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        if self._showing_expired:
            QMessageBox.warning(self, '提示', '过期材质单视图为只读，请先切换到材料类别。')
            return

        files = []
        for url in event.mimeData().urls():
            f = url.toLocalFile()
            ext = f.lower()
            if ext.endswith(('.png', '.jpg', '.jpeg', '.pdf')):
                files.append(f)

        if files:
            self._show_upload_dialog(files)
