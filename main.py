"""
检测报告生成系统 - 程序入口
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
from PySide6.QtGui import QFont
from core.app_paths import (
    default_work_dir, ensure_workspace, work_dir,
)
from ui.main_window import ReportGenerator


def _ensure_startup_workspace():
    if work_dir():
        ensure_workspace()
        return True

    default_dir = default_work_dir()
    message = (
        '首次启动需要选择一个工作目录，用来保存 Excel 模板、材质单库和报告输出。\n\n'
        f'推荐目录：{default_dir}\n\n'
        '是否使用推荐目录？'
    )
    reply = QMessageBox.question(
        None, '选择工作目录', message,
        QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
        QMessageBox.Yes,
    )
    if reply == QMessageBox.Cancel:
        QMessageBox.information(None, '已取消', '未选择工作目录，程序将退出。')
        return False
    if reply == QMessageBox.Yes:
        chosen = default_dir
    else:
        start_dir = os.path.dirname(default_dir)
        chosen = QFileDialog.getExistingDirectory(
            None, '选择融海报告生成工作目录', start_dir)
        if not chosen:
            QMessageBox.information(None, '已取消', '未选择工作目录，程序将退出。')
            return False

    ensure_workspace(chosen)
    return True


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont('Microsoft YaHei', 10))
    app.setApplicationName('融海运通出厂检验报告生成系统')
    if not _ensure_startup_workspace():
        return
    window = ReportGenerator()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
