from __future__ import annotations

import csv
import os
import sys
import tempfile
import time
from pathlib import Path

import openpyxl
from PyQt6.QtCore import Qt
from PyQt6.QtTest import QTest
from PyQt6.QtWidgets import QApplication, QFileDialog, QMessageBox, QPushButton

from excel_tool import ExcelCompareTool, configure_application


os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def wait_for(predicate, app: QApplication, timeout_ms: int = 10000, step_ms: int = 50) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        app.processEvents()
        if predicate():
            return True
        QTest.qWait(step_ms)
    return predicate()


def write_csv(path: Path, rows: list[list[object]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerows(rows)


def find_button_by_role(window: ExcelCompareTool, role: str) -> QPushButton:
    for button in window.findChildren(QPushButton):
        if button.property("role") == role:
            return button
    raise AssertionError(f"找不到按钮角色: {role}")


def main() -> int:
    app = QApplication([])
    configure_application(app)
    app.setQuitOnLastWindowClosed(False)

    with tempfile.TemporaryDirectory(prefix="excel_tool_gui_test_") as temp_dir:
        temp_path = Path(temp_dir)
        left_path = temp_path / "left.csv"
        right_path = temp_path / "right.csv"

        write_csv(
            left_path,
            [
                ["id", "name", "city"],
                [1, "Alice", "Shanghai"],
                [2, "Bob", "Beijing"],
                [3, "Cindy", "Shenzhen"],
            ],
        )
        write_csv(
            right_path,
            [
                ["user_id", "order_no", "amount"],
                [1, "A001", 32],
                [1, "A002", 64],
                [3, "C009", 12],
            ],
        )

        selected_files = [str(left_path), str(right_path)]
        dialog_calls: list[tuple[str, str]] = []
        message_calls: list[tuple[str, str, str]] = []

        original_dialog = QFileDialog.getOpenFileName
        original_info = QMessageBox.information
        original_warning = QMessageBox.warning
        original_critical = QMessageBox.critical

        def fake_dialog(*args, **kwargs):
            if not selected_files:
                return "", ""
            path = selected_files.pop(0)
            dialog_calls.append(("open", path))
            return path, ""

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
            app.processEvents()

            choose_left = find_button_by_role(window, "select-file-1")
            choose_right = find_button_by_role(window, "select-file-2")
            export_button = find_button_by_role(window, "export-result")

            QTest.mouseClick(choose_left, Qt.MouseButton.LeftButton)
            QTest.mouseClick(choose_right, Qt.MouseButton.LeftButton)

            if not wait_for(lambda: bool(window.file_states[1].columns), app):
                raise AssertionError("左侧文件预览未在超时前加载完成")
            if not wait_for(lambda: bool(window.file_states[2].columns), app):
                raise AssertionError("右侧文件预览未在超时前加载完成")
            if not wait_for(lambda: len(window.match_rows) == 1, app):
                raise AssertionError("默认匹配行未创建")

            row = window.match_rows[0]
            left_index = row.left_combo.findData("id")
            right_index = row.right_combo.findData("user_id")
            if left_index < 0 or right_index < 0:
                raise AssertionError("匹配列下拉框未加载预期字段")
            row.left_combo.setCurrentIndex(left_index)
            row.right_combo.setCurrentIndex(right_index)
            app.processEvents()

            left_checkbox = next(
                checkbox for checkbox in window.output_checkboxes["left"] if checkbox.property("raw_name") == "name"
            )
            right_checkbox = next(
                checkbox for checkbox in window.output_checkboxes["right"] if checkbox.property("raw_name") == "order_no"
            )
            left_checkbox.setChecked(True)
            right_checkbox.setChecked(True)
            app.processEvents()

            window.output_name_edit.setText("gui_smoke_result")
            QTest.mouseClick(export_button, Qt.MouseButton.LeftButton)

            if not wait_for(lambda: window.export_worker is None, app, timeout_ms=15000):
                raise AssertionError("导出任务未在超时前完成")

            output_path = temp_path / "gui_smoke_result.xlsx"
            if not output_path.exists():
                raise AssertionError("导出文件不存在")

            workbook = openpyxl.load_workbook(output_path, read_only=True)
            try:
                rows = list(workbook.active.iter_rows(values_only=True))
            finally:
                workbook.close()

            expected_header = ("A · id", "B · name", "B · order_no · 被匹配文件")
            if rows[0] != expected_header:
                raise AssertionError(f"导出表头不符合预期: {rows[0]!r}")
            if len(rows) != 4:
                raise AssertionError(f"导出行数不符合预期: {len(rows)}")
            expected_rows = [
                ("1", "Alice", "A001"),
                ("1", "Alice", "A002"),
                ("3", "Cindy", "C009"),
            ]
            if list(rows[1:]) != expected_rows:
                raise AssertionError(f"导出内容不符合预期: {rows[1:]!r}")
            if not any(kind == "info" and title == "导出完成" for kind, title, _ in message_calls):
                raise AssertionError(f"未捕获到成功提示框: {message_calls!r}")

            print("dialog_calls", dialog_calls)
            print("message_calls", message_calls)
            print("output_rows", rows)
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
