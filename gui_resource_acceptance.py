from __future__ import annotations

import sys
import time
from pathlib import Path

import openpyxl
from PyQt6.QtCore import Qt
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox, QPushButton

from excel_tool import ExcelCompareTool, configure_application


ROOT = Path(__file__).resolve().parent
LEFT_FILE = ROOT / "resource" / "城市表.xlsx"
RIGHT_FILE = ROOT / "resource" / "明细表.xlsx"
OUTPUT_FILE = ROOT / "resource" / "资源验收GUI测试.xlsx"


def wait_for(predicate, app: QApplication, timeout_ms: int = 30000, step_ms: int = 50) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        app.processEvents()
        if predicate():
            return True
        QTest.qWait(step_ms)
    return predicate()


def find_button(window: ExcelCompareTool, role: str) -> QPushButton:
    for button in window.findChildren(QPushButton):
        if button.property("role") == role:
            return button
    raise RuntimeError(f"找不到按钮: {role}")


def main() -> int:
    app = QApplication([])
    configure_application(app)
    app.setQuitOnLastWindowClosed(False)

    if OUTPUT_FILE.exists():
        OUTPUT_FILE.unlink()

    selected_files = [str(LEFT_FILE), str(RIGHT_FILE)]
    message_calls: list[tuple[str, str, str]] = []

    original_dialog = QFileDialog.getOpenFileName
    original_info = QMessageBox.information
    original_warning = QMessageBox.warning
    original_critical = QMessageBox.critical

    def fake_dialog(*args, **kwargs):
        if not selected_files:
            return "", ""
        return selected_files.pop(0), ""

    def fake_messagebox(kind: str):
        def _handler(parent, title, text, *args, **kwargs):
            message_calls.append((kind, title, text))
            return QMessageBox.StandardButton.Ok

        return _handler

    QFileDialog.getOpenFileName = fake_dialog
    QMessageBox.information = fake_messagebox("info")
    QMessageBox.warning = fake_messagebox("warning")
    QMessageBox.critical = fake_messagebox("critical")

    try:
        window = ExcelCompareTool()
        window.show()
        window.raise_()
        window.activateWindow()
        app.processEvents()
        QTest.qWait(1200)

        QTest.mouseClick(find_button(window, "select-file-1"), Qt.MouseButton.LeftButton)
        QTest.mouseClick(find_button(window, "select-file-2"), Qt.MouseButton.LeftButton)

        if not wait_for(lambda: len(window.file_states[1].sheet_names) > 0, app):
            raise AssertionError("城市表工作表列表未加载完成")
        if not wait_for(lambda: len(window.file_states[2].sheet_names) > 0, app):
            raise AssertionError("明细表工作表列表未加载完成")

        window.file_widgets[1]["sheet_combo"].setCurrentText("广州")
        window.file_widgets[2]["sheet_combo"].setCurrentText("匹配明细")
        app.processEvents()
        if not wait_for(
            lambda: window.file_states[1].selected_sheet == "广州"
            and window.file_states[2].selected_sheet == "匹配明细"
            and bool(window.file_states[1].columns)
            and bool(window.file_states[2].columns),
            app,
        ):
            raise AssertionError("工作表切换未完成")

        if not wait_for(lambda: len(window.match_rows) == 1, app):
            raise AssertionError("默认匹配行未创建")

        row = window.match_rows[0]
        row.left_combo.setCurrentIndex(row.left_combo.findData("集团诊断名称"))
        row.right_combo.setCurrentIndex(row.right_combo.findData("分类名称"))
        app.processEvents()

        left_city = next(
            checkbox for checkbox in window.output_checkboxes["left"] if checkbox.property("raw_name") == "城市"
        )
        right_status = next(
            checkbox for checkbox in window.output_checkboxes["right"] if checkbox.property("raw_name") == "状态"
        )
        right_category_id = next(
            checkbox for checkbox in window.output_checkboxes["right"] if checkbox.property("raw_name") == "分类ID"
        )

        left_city.setChecked(True)
        right_status.setChecked(True)
        right_category_id.setChecked(True)
        window.output_name_edit.setText("资源验收GUI测试")
        app.processEvents()
        QTest.qWait(500)

        QTest.mouseClick(find_button(window, "export-result"), Qt.MouseButton.LeftButton)

        if not wait_for(lambda: window.export_worker is None, app, timeout_ms=120000, step_ms=100):
            raise AssertionError("真实资源导出未在超时前完成")

        if not OUTPUT_FILE.exists():
            raise AssertionError("GUI 验收导出文件不存在")

        workbook = openpyxl.load_workbook(OUTPUT_FILE, read_only=True)
        try:
            sheet = workbook.active
            rows = sheet.iter_rows(values_only=True)
            header = next(rows)
            data_count = sum(1 for _ in rows)
        finally:
            workbook.close()

        print("acceptance_output", OUTPUT_FILE)
        print("acceptance_header", header)
        print("acceptance_rows", data_count)
        print("message_calls", message_calls)

        QTest.qWait(1200)
        window.close()
        app.processEvents()
        return 0
    finally:
        QFileDialog.getOpenFileName = original_dialog
        QMessageBox.information = original_info
        QMessageBox.warning = original_warning
        QMessageBox.critical = original_critical


if __name__ == "__main__":
    raise SystemExit(main())
