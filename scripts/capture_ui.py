"""Render representative V2.0 widgets without opening user documents."""
import os
import sys
os.environ["QT_QPA_PLATFORM"] = "windows" if "--native" in sys.argv else "offscreen"
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from word_text_tool.app import create_application
from word_text_tool.ui.main_window import MainWindow
from PyQt6.QtCore import Qt

app = create_application([])
window = MainWindow()
if "--native" in sys.argv:
    window.setAttribute(Qt.WidgetAttribute.WA_DontShowOnScreen, True)
output = ROOT / "build" / "ui-verification"
output.mkdir(exist_ok=True, parents=True)
window.number_panel.mode.setCurrentIndex(1)
window.number_panel.fixed_text.setText("PROJECT-2026")
for width, height in [(820, 700), (1000, 900), (1400, 1000)]:
    window.resize(width, height)
    window.show()
    for index in (0, 1, 2):
        window.tabs.setCurrentIndex(index)
        app.processEvents()
        window.grab().save(str(output / f"tab{index}-{width}x{height}.png"))
window.close()
print(output)
