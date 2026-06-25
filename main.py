"""
检测报告生成系统 - 程序入口
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from core.app_paths import initialize_workspace
from ui.main_window import ReportGenerator


def main():
    initialize_workspace()
    app = QApplication(sys.argv)
    app.setFont(QFont('Microsoft YaHei', 10))
    app.setApplicationName('融海运通出厂检验报告生成系统')
    window = ReportGenerator()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
