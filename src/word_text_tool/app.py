"""Application identity and desktop entry point."""

import sys

from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication

from word_text_tool.ui.main_window import MainWindow


def create_application(argv=None) -> QApplication:
    application = QApplication.instance() or QApplication(
        sys.argv if argv is None else argv
    )
    application.setApplicationName("Word 文本处理工具")
    application.setApplicationVersion("1.0")
    application.setOrganizationName("WordTextTool")
    application.setFont(QFont("Microsoft YaHei UI", 10))
    return application


def run() -> int:
    application = create_application()
    window = MainWindow()
    window.show()
    return application.exec()
