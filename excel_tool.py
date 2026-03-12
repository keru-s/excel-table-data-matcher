from __future__ import annotations

import sys
import threading
from dataclasses import dataclass, field

from PyQt6.QtCore import QEvent, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QPalette, QStandardItem, QStandardItemModel
from PyQt6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from excel_core import (
    CSV_SHEET_NAME,
    DEDUP_STRATEGY_BOTH_KEYS,
    DEDUP_STRATEGY_LEFT_KEY_RIGHT_ROW,
    DEDUP_STRATEGY_NONE,
    LEFT_SOURCE,
    RIGHT_SOURCE,
    ColumnOption,
    ExportCancelled,
    ExportConfig,
    MatchPair,
    OutputColumn,
    PreviewData,
    export_matches,
    load_preview,
)

COMMON_UI_FONT_FAMILY = "Arial"


@dataclass(frozen=True)
class ThemeTokens:
    background: str
    surface: str
    elevated: str
    border: str
    text: str
    muted: str
    accent: str
    accent_hover: str
    accent_soft: str
    danger: str
    table_alt: str
    shadow: str


THEMES = {
    "light": ThemeTokens(
        background="#f5f5f5",
        surface="#ffffff",
        elevated="#fbfbfb",
        border="#dddddd",
        text="#161616",
        muted="#6b7280",
        accent="#1677ff",
        accent_hover="#0f5fd1",
        accent_soft="#e8f1ff",
        danger="#dc5d5d",
        table_alt="#fafafa",
        shadow="rgba(15, 23, 42, 0.08)",
    ),
    "dark": ThemeTokens(
        background="#0b0b0c",
        surface="#151517",
        elevated="#1c1c1f",
        border="#2b2c30",
        text="#f5f5f5",
        muted="#a1a1aa",
        accent="#3b82f6",
        accent_hover="#2563eb",
        accent_soft="#1d2838",
        danger="#ff7b7b",
        table_alt="#17181b",
        shadow="rgba(0, 0, 0, 0.28)",
    ),
}


@dataclass
class FileUiState:
    path: str = ""
    sheet_names: list[str] = field(default_factory=list)
    selected_sheet: str = ""
    columns: list[ColumnOption] = field(default_factory=list)
    total_rows: int = 0
    used_encoding: str = "gbk"


class PreviewWorker(QThread):
    loaded = pyqtSignal(int, int, object)
    failed = pyqtSignal(int, int, str)

    def __init__(self, file_key: int, request_id: int, file_path: str, sheet_name: str | None, encoding: str):
        super().__init__()
        self.file_key = file_key
        self.request_id = request_id
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.encoding = encoding

    def run(self) -> None:
        try:
            preview = load_preview(self.file_path, self.sheet_name, self.encoding)
            self.loaded.emit(self.file_key, self.request_id, preview)
        except Exception as exc:
            self.failed.emit(self.file_key, self.request_id, str(exc))


class ExportWorker(QThread):
    succeeded = pyqtSignal(str)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()
    progress_changed = pyqtSignal(int, str)

    def __init__(self, config: ExportConfig):
        super().__init__()
        self.config = config
        self._cancel_event = threading.Event()

    def request_cancel(self) -> None:
        self._cancel_event.set()

    def run(self) -> None:
        try:
            output_path = export_matches(
                self.config,
                progress_callback=lambda percent, message: self.progress_changed.emit(percent, message),
                is_cancelled=self._cancel_event.is_set,
            )
        except ExportCancelled:
            self.cancelled.emit()
        except Exception as exc:
            self.failed.emit(str(exc))
        else:
            self.succeeded.emit(output_path)


class MatchRowWidget(QFrame):
    removed = pyqtSignal(object)
    changed = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("MatchRow")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.left_combo = QComboBox()
        self.right_combo = QComboBox()
        self.remove_button = QPushButton("移除")
        self.remove_button.setObjectName("GhostButton")

        layout.addWidget(self.left_combo, 1)
        layout.addWidget(self.right_combo, 1)
        layout.addWidget(self.remove_button)

        self.left_combo.currentIndexChanged.connect(self.changed)
        self.right_combo.currentIndexChanged.connect(self.changed)
        self.remove_button.clicked.connect(lambda: self.removed.emit(self))

    def current_left_raw(self) -> str:
        return self.left_combo.currentData() or ""

    def current_right_raw(self) -> str:
        return self.right_combo.currentData() or ""

    def current_left_label(self) -> str:
        return self.left_combo.currentText()

    def current_right_label(self) -> str:
        return self.right_combo.currentText()


class ExcelCompareTool(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        configure_application(QApplication.instance())
        self.theme_preference = "system"
        self.file_states = {1: FileUiState(), 2: FileUiState()}
        self.preview_workers: dict[int, PreviewWorker] = {}
        self.preview_request_ids = {1: 0, 2: 0}
        self.export_worker: ExportWorker | None = None
        self.progress_dialog: QProgressDialog | None = None
        self.match_rows: list[MatchRowWidget] = []
        self.output_checkboxes = {LEFT_SOURCE: [], RIGHT_SOURCE: []}
        self.current_theme = "light"

        self._build_ui()
        self._bind_system_theme_tracking()
        self._apply_theme()

    def _build_ui(self) -> None:
        self.setWindowTitle("Excel 数据匹配工具")
        self.resize(1280, 820)
        self.setMinimumSize(1040, 700)

        root = QWidget()
        root.setObjectName("Root")
        self.setCentralWidget(root)

        main_layout = QVBoxLayout(root)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)

        self.file_widgets = {
            1: self._create_file_panel(1, "匹配源文件", "左侧记录作为主输出来源"),
            2: self._create_file_panel(2, "被匹配文件", "右侧记录补充附加字段"),
        }

        file_splitter = QSplitter(Qt.Orientation.Horizontal)
        file_splitter.setChildrenCollapsible(False)
        file_splitter.setHandleWidth(8)
        file_splitter.addWidget(self.file_widgets[1]["card"])
        file_splitter.addWidget(self.file_widgets[2]["card"])
        file_splitter.setSizes([620, 620])
        file_splitter.setStretchFactor(0, 1)
        file_splitter.setStretchFactor(1, 1)

        match_card = self._create_match_card()
        export_card = self._create_export_card()

        config_splitter = QSplitter(Qt.Orientation.Vertical)
        config_splitter.setChildrenCollapsible(False)
        config_splitter.setHandleWidth(8)
        config_splitter.addWidget(match_card)
        config_splitter.addWidget(export_card)
        config_splitter.setSizes([190, 430])
        config_splitter.setStretchFactor(0, 0)
        config_splitter.setStretchFactor(1, 1)

        page_splitter = QSplitter(Qt.Orientation.Vertical)
        page_splitter.setChildrenCollapsible(False)
        page_splitter.setHandleWidth(8)
        page_splitter.addWidget(file_splitter)
        page_splitter.addWidget(config_splitter)
        page_splitter.setSizes([280, 540])
        page_splitter.setStretchFactor(0, 1)
        page_splitter.setStretchFactor(1, 1)
        main_layout.addWidget(page_splitter, 1)

        status_bar = QStatusBar()
        status_bar.showMessage("就绪")
        self.setStatusBar(status_bar)

    def _create_file_panel(self, file_key: int, title_text: str, subtitle_text: str) -> dict[str, object]:
        card = self._create_card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        title = QLabel(title_text)
        title.setObjectName("CardTitle")

        path_edit = QLineEdit()
        path_edit.setPlaceholderText("请选择 Excel / CSV 文件")
        path_edit.setReadOnly(True)

        choose_button = QPushButton("选择文件")
        choose_button.setObjectName("PrimaryButton")
        choose_button.setProperty("role", f"select-file-{file_key}")
        choose_button.clicked.connect(lambda: self._select_file(file_key))

        sheet_combo = QComboBox()
        sheet_combo.setEnabled(False)
        sheet_combo.currentIndexChanged.connect(lambda: self._reload_selected_sheet(file_key))

        info_label = QLabel("尚未加载文件")
        info_label.setObjectName("InfoLabel")

        top_grid = QGridLayout()
        top_grid.setHorizontalSpacing(8)
        top_grid.setVerticalSpacing(6)
        top_grid.setColumnStretch(1, 1)
        top_grid.addWidget(QLabel("文件路径"), 0, 0)
        top_grid.addWidget(path_edit, 0, 1)
        top_grid.addWidget(choose_button, 0, 2)
        top_grid.addWidget(QLabel("工作表"), 1, 0)
        top_grid.addWidget(sheet_combo, 1, 1, 1, 2)

        preview_table = QTableView()
        preview_table.setAlternatingRowColors(True)
        preview_table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        preview_table.setSelectionMode(QTableView.SelectionMode.NoSelection)
        preview_table.verticalHeader().setVisible(False)
        preview_table.verticalHeader().setDefaultSectionSize(26)
        preview_table.horizontalHeader().setFixedHeight(30)
        preview_table.horizontalHeader().setStretchLastSection(False)
        preview_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        preview_table.horizontalHeader().setMinimumSectionSize(110)
        preview_table.setWordWrap(False)
        preview_table.setMinimumHeight(140)
        preview_table.setHorizontalScrollMode(QTableView.ScrollMode.ScrollPerPixel)
        preview_table.setVerticalScrollMode(QTableView.ScrollMode.ScrollPerPixel)
        preview_table.setVisible(False)

        model = QStandardItemModel()
        preview_table.setModel(model)

        preview_empty = QLabel("选择 Excel 或 CSV 文件后，这里会显示预览。")
        preview_empty.setObjectName("EmptyState")
        preview_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_empty.setMinimumHeight(76)

        layout.addWidget(title)
        layout.addLayout(top_grid)
        layout.addWidget(info_label)
        layout.addWidget(preview_empty, 1)
        layout.addWidget(preview_table, 1)

        return {
            "card": card,
            "path_edit": path_edit,
            "choose_button": choose_button,
            "sheet_combo": sheet_combo,
            "info_label": info_label,
            "preview_empty": preview_empty,
            "preview_table": preview_table,
            "model": model,
        }

    def _create_match_card(self) -> QWidget:
        card = self._create_card()
        self.match_card = card
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        title = QLabel("匹配规则")
        title.setObjectName("CardTitle")
        subtitle = QLabel("支持多列联合匹配。预览和列配置都会在后台刷新。")
        subtitle.setObjectName("CardSubtitle")

        button_row = QHBoxLayout()
        add_button = QPushButton("新增匹配列")
        add_button.setObjectName("PrimaryButton")
        add_button.clicked.connect(self._add_match_row)
        button_row.addWidget(add_button)
        button_row.addStretch()

        self.match_rows_container = QWidget()
        self.match_rows_layout = QVBoxLayout(self.match_rows_container)
        self.match_rows_layout.setContentsMargins(0, 0, 0, 0)
        self.match_rows_layout.setSpacing(6)
        self.match_rows_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        match_scroll = QScrollArea()
        match_scroll.setWidgetResizable(True)
        match_scroll.setFrameShape(QFrame.Shape.NoFrame)
        match_scroll.setObjectName("SelectionScroll")
        match_scroll.setWidget(self.match_rows_container)
        match_scroll.setMinimumHeight(0)
        match_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        match_scroll.setVisible(False)
        self.match_scroll = match_scroll
        self.match_empty_state = QLabel("加载两侧文件后，点击“新增匹配列”开始配置。")
        self.match_empty_state.setObjectName("EmptyState")
        self.match_empty_state.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addLayout(button_row)
        layout.addWidget(self.match_empty_state)
        layout.addWidget(match_scroll, 1)
        self._update_match_empty_state()
        return card

    def _create_export_card(self) -> QWidget:
        card = self._create_card()
        card.setMinimumHeight(380)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        title = QLabel("输出设置")
        title.setObjectName("CardTitle")
        subtitle = QLabel("导出时按需选择输出列，未选中的列不会进入最终结果。")
        subtitle.setObjectName("CardSubtitle")

        name_row = QGridLayout()
        name_row.setHorizontalSpacing(8)
        name_row.setVerticalSpacing(6)
        name_row.setColumnStretch(1, 1)

        self.output_name_edit = QLineEdit("匹配结果")
        name_row.addWidget(QLabel("输出文件名"), 0, 0)
        name_row.addWidget(self.output_name_edit, 0, 1)
        name_row.addWidget(QLabel(".xlsx"), 0, 2)

        self.encoding_group = QButtonGroup(self)
        self.gbk_radio = QRadioButton("GBK")
        self.gbk_radio.setChecked(True)
        self.utf8_radio = QRadioButton("UTF-8")
        self.encoding_group.addButton(self.gbk_radio)
        self.encoding_group.addButton(self.utf8_radio)

        self.gbk_radio.toggled.connect(self._reload_previews_after_encoding_change)
        self.utf8_radio.toggled.connect(self._reload_previews_after_encoding_change)
        self.gbk_radio.setObjectName("CompactChoiceButton")
        self.utf8_radio.setObjectName("CompactChoiceButton")

        self.dedup_combo = QComboBox()
        self.dedup_combo.setObjectName("CompactCombo")
        self.dedup_combo.addItem("左侧按匹配列去重", DEDUP_STRATEGY_LEFT_KEY_RIGHT_ROW)
        self.dedup_combo.addItem("不去重", DEDUP_STRATEGY_NONE)
        self.dedup_combo.addItem("两侧都按匹配列去重", DEDUP_STRATEGY_BOTH_KEYS)

        output_panels = QWidget()
        output_panels.setMinimumHeight(280)
        output_panels.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        output_panels_layout = QHBoxLayout(output_panels)
        output_panels_layout.setContentsMargins(0, 0, 0, 0)
        output_panels_layout.setSpacing(8)
        output_panels_layout.addWidget(self._create_checkbox_group(LEFT_SOURCE, "匹配源文件输出列"), 1)
        output_panels_layout.addWidget(self._create_checkbox_group(RIGHT_SOURCE, "被匹配文件输出列"), 1)

        footer_row = QHBoxLayout()
        footer_row.setContentsMargins(0, 0, 0, 0)
        footer_row.setSpacing(12)

        encoding_row = QHBoxLayout()
        encoding_row.setContentsMargins(0, 0, 0, 0)
        encoding_row.setSpacing(6)
        encoding_row.addWidget(QLabel("CSV 编码"))
        encoding_row.addWidget(self.gbk_radio)
        encoding_row.addWidget(self.utf8_radio)

        dedup_row = QHBoxLayout()
        dedup_row.setContentsMargins(0, 0, 0, 0)
        dedup_row.setSpacing(6)
        dedup_row.addWidget(QLabel("去重策略"))
        self.dedup_combo.setMaximumWidth(260)
        dedup_row.addWidget(self.dedup_combo)

        footer_row.addLayout(encoding_row, 0)
        footer_row.addLayout(dedup_row, 1)
        footer_row.addStretch()
        self.export_button = QPushButton("导出结果")
        self.export_button.setObjectName("PrimaryButton")
        self.export_button.setProperty("role", "export-result")
        self.export_button.clicked.connect(self._start_export)
        footer_row.addWidget(self.export_button)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addLayout(name_row)
        layout.addWidget(output_panels, 1)
        layout.addLayout(footer_row)
        return card

    def _create_setting_panel(self, title_text: str, caption_text: str, content_layout: QVBoxLayout | QHBoxLayout) -> QWidget:
        panel = QFrame()
        panel.setObjectName("SettingPanel")
        panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        title = QLabel(title_text)
        title.setObjectName("SettingTitle")
        caption = QLabel(caption_text)
        caption.setObjectName("SettingCaption")

        layout.addWidget(title)
        layout.addWidget(caption)
        layout.addLayout(content_layout)
        return panel

    def _create_checkbox_group(self, source: str, title_text: str) -> QWidget:
        group = QFrame()
        group.setObjectName("SelectionPanel")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        title = QLabel(title_text)
        title.setObjectName("PanelTitle")
        caption = QLabel("匹配列会自动禁用，避免重复输出。")
        caption.setObjectName("PanelCaption")

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setObjectName("SelectionScroll")
        scroll.setMinimumHeight(0)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(8)
        container_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll.setWidget(container)

        layout.addWidget(title)
        layout.addWidget(caption)
        layout.addWidget(scroll)
        self.output_checkboxes[source] = []
        setattr(self, f"{source}_checkbox_container", container)
        setattr(self, f"{source}_checkbox_layout", container_layout)
        return group

    def _create_card(self, object_name: str = "Card") -> QFrame:
        card = QFrame()
        card.setObjectName(object_name)
        return card

    def _bind_system_theme_tracking(self) -> None:
        app = QApplication.instance()
        if app is None:
            return
        app.installEventFilter(self)
        style_hints = app.styleHints()
        if hasattr(style_hints, "colorSchemeChanged"):
            style_hints.colorSchemeChanged.connect(lambda *_: self._apply_theme_if_following_system())

    def eventFilter(self, watched: object, event: QEvent) -> bool:
        theme_types = {QEvent.Type.ApplicationPaletteChange, QEvent.Type.PaletteChange}
        theme_change = getattr(QEvent.Type, "ThemeChange", None)
        if theme_change is not None:
            theme_types.add(theme_change)

        if watched is QApplication.instance() and event.type() in theme_types:
            self._apply_theme_if_following_system()
        return super().eventFilter(watched, event)

    def _apply_theme_if_following_system(self) -> None:
        if self.theme_preference == "system":
            self._apply_theme()

    def _apply_theme(self) -> None:
        self.current_theme = self._resolve_theme_name(self.theme_preference)
        tokens = THEMES[self.current_theme]
        self._apply_palette(tokens)
        self.setStyleSheet(self._build_stylesheet(tokens))
        self.statusBar().setStyleSheet(f"QStatusBar {{ background: {tokens.surface}; color: {tokens.muted}; border-top: 1px solid {tokens.border}; }}")

    def _resolve_theme_name(self, preference: str) -> str:
        if preference == "light":
            return "light"
        if preference == "dark":
            return "dark"
        app = QApplication.instance()
        if app is not None:
            style_hints = app.styleHints()
            if hasattr(style_hints, "colorScheme"):
                try:
                    if style_hints.colorScheme() == Qt.ColorScheme.Dark:
                        return "dark"
                    return "light"
                except Exception:
                    pass
            lightness = app.palette().color(QPalette.ColorRole.Window).lightness()
            return "dark" if lightness < 128 else "light"
        return "light"

    def _apply_palette(self, tokens: ThemeTokens) -> None:
        app = QApplication.instance()
        if app is None:
            return

        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(tokens.background))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(tokens.text))
        palette.setColor(QPalette.ColorRole.Base, QColor(tokens.surface))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(tokens.table_alt))
        palette.setColor(QPalette.ColorRole.Text, QColor(tokens.text))
        palette.setColor(QPalette.ColorRole.Button, QColor(tokens.surface))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(tokens.text))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(tokens.surface))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(tokens.text))
        palette.setColor(QPalette.ColorRole.BrightText, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(tokens.accent))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(tokens.muted))
        app.setPalette(palette)

    def _build_stylesheet(self, tokens: ThemeTokens) -> str:
        return f"""
        QWidget {{
            font-family: "{COMMON_UI_FONT_FAMILY}";
        }}
        QWidget#Root {{
            background: {tokens.background};
        }}
        QFrame#Card, QFrame#SelectionPanel {{
            background: {tokens.surface};
            border: 1px solid {tokens.border};
            border-radius: 16px;
        }}
        QLabel {{
            color: {tokens.text};
        }}
        QLabel#CardSubtitle {{
            color: {tokens.muted};
            font-size: 11px;
        }}
        QLabel#CardTitle {{
            font-size: 16px;
            font-weight: 700;
        }}
        QLabel#PanelTitle {{
            font-size: 13px;
            font-weight: 700;
        }}
        QLabel#PanelCaption, QLabel#InfoLabel, QLabel#SettingCaption {{
            color: {tokens.muted};
            font-size: 11px;
        }}
        QLabel#EmptyState {{
            color: {tokens.muted};
            font-size: 11px;
            border: 1px dashed {tokens.border};
            border-radius: 12px;
            background: {tokens.elevated};
            padding: 12px 10px;
        }}
        QLineEdit, QComboBox, QTableView, QScrollArea#SelectionScroll {{
            background: {tokens.elevated};
            color: {tokens.text};
        }}
        QLineEdit, QComboBox {{
            border: 1px solid {tokens.border};
            border-radius: 9px;
            padding: 6px 9px;
            min-height: 18px;
            font-size: 12px;
        }}
        QComboBox::drop-down {{
            border: none;
            width: 24px;
        }}
        QComboBox QAbstractItemView {{
            background: {tokens.surface};
            color: {tokens.text};
            border: 1px solid {tokens.border};
            selection-background-color: {tokens.accent_soft};
        }}
        QPushButton {{
            border-radius: 9px;
            padding: 6px 12px;
            border: 1px solid {tokens.border};
            background: {tokens.surface};
            color: {tokens.text};
            font-weight: 600;
            min-height: 18px;
            font-size: 12px;
        }}
        QPushButton:hover {{
            border-color: {tokens.accent};
        }}
        QPushButton#PrimaryButton {{
            background: {tokens.accent};
            border-color: {tokens.accent};
            color: white;
        }}
        QPushButton#PrimaryButton:hover {{
            background: {tokens.accent_hover};
            border-color: {tokens.accent_hover};
        }}
        QPushButton#GhostButton {{
            background: transparent;
            color: {tokens.muted};
        }}
        QTableView {{
            border: 1px solid {tokens.border};
            border-radius: 12px;
            gridline-color: {tokens.border};
            alternate-background-color: {tokens.table_alt};
            selection-background-color: {tokens.accent_soft};
            background-color: {tokens.surface};
            font-size: 12px;
        }}
        QHeaderView::section {{
            background: {tokens.elevated};
            color: {tokens.muted};
            padding: 6px 8px;
            border: none;
            border-bottom: 1px solid {tokens.border};
            font-size: 11px;
            font-weight: 600;
        }}
        QTableCornerButton::section {{
            background: {tokens.elevated};
            border: none;
            border-bottom: 1px solid {tokens.border};
            border-right: 1px solid {tokens.border};
        }}
        QAbstractScrollArea {{
            background: {tokens.surface};
        }}
        QAbstractScrollArea > QWidget {{
            background: {tokens.surface};
            color: {tokens.text};
        }}
        QCheckBox {{
            spacing: 8px;
            padding: 6px 8px;
            border-radius: 10px;
            background: {tokens.elevated};
            border: 1px solid {tokens.border};
            font-size: 12px;
        }}
        QCheckBox::indicator {{
            width: 14px;
            height: 14px;
            border-radius: 5px;
            border: 1px solid {tokens.border};
            background: {tokens.surface};
        }}
        QCheckBox::indicator:checked {{
            background: {tokens.accent};
            border-color: {tokens.accent};
        }}
        QCheckBox:disabled {{
            color: {tokens.muted};
            background: {tokens.table_alt};
        }}
        QRadioButton#ChoiceButton {{
            spacing: 0px;
            font-size: 13px;
            font-weight: 600;
            color: {tokens.muted};
            background: {tokens.elevated};
            border: 1px solid {tokens.border};
            border-radius: 10px;
            padding: 7px 12px;
            min-height: 18px;
        }}
        QRadioButton#ChoiceButton::indicator {{
            width: 0px;
            height: 0px;
            margin: 0px;
        }}
        QRadioButton#ChoiceButton:hover {{
            border-color: {tokens.accent};
            color: {tokens.text};
        }}
        QRadioButton#ChoiceButton:checked {{
            background: {tokens.accent_soft};
            border-color: {tokens.accent};
            color: {tokens.text};
        }}
        QRadioButton#CompactChoiceButton {{
            spacing: 0px;
            font-size: 11px;
            font-weight: 600;
            color: {tokens.muted};
            background: {tokens.elevated};
            border: 1px solid {tokens.border};
            border-radius: 9px;
            padding: 4px 8px;
            min-height: 14px;
        }}
        QRadioButton#CompactChoiceButton::indicator {{
            width: 0px;
            height: 0px;
            margin: 0px;
        }}
        QRadioButton#CompactChoiceButton:checked {{
            background: {tokens.accent_soft};
            border-color: {tokens.accent};
            color: {tokens.text};
        }}
        QComboBox#CompactCombo {{
            min-width: 200px;
            min-height: 16px;
            font-size: 11px;
            padding: 4px 8px;
        }}
        QFrame#MatchRow {{
            background: {tokens.elevated};
            border: 1px solid {tokens.border};
            border-radius: 10px;
            padding: 6px;
        }}
        QFrame#SelectionPanel {{
            background: transparent;
            border: none;
        }}
        QScrollArea#SelectionScroll {{
            border: 1px solid {tokens.border};
            border-radius: 12px;
            background: {tokens.elevated};
        }}
        QSplitter::handle {{
            background: transparent;
        }}
        """

    def _select_file(self, file_key: int) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择文件",
            "",
            "Excel/CSV files (*.xlsx *.xls *.csv);;All files (*.*)",
        )
        if not file_path:
            return

        state = self.file_states[file_key]
        state.path = file_path
        state.selected_sheet = ""
        widgets = self.file_widgets[file_key]
        widgets["path_edit"].setText(file_path)
        widgets["info_label"].setText("正在后台加载预览...")
        widgets["preview_empty"].setText("正在后台加载预览...")
        widgets["preview_empty"].setVisible(True)
        widgets["preview_table"].setVisible(False)
        self.statusBar().showMessage(f"正在加载文件 {file_key} 预览")
        self._load_preview_async(file_key, file_path, None)

    def _reload_selected_sheet(self, file_key: int) -> None:
        state = self.file_states[file_key]
        if not state.path:
            return
        sheet_name = self.file_widgets[file_key]["sheet_combo"].currentText()
        if not sheet_name:
            return
        if sheet_name == state.selected_sheet:
            return
        self._load_preview_async(file_key, state.path, sheet_name)

    def _reload_previews_after_encoding_change(self) -> None:
        if not self.sender() or not isinstance(self.sender(), QRadioButton):
            return
        for file_key, state in self.file_states.items():
            if state.path and state.path.lower().endswith(".csv"):
                self._load_preview_async(file_key, state.path, state.selected_sheet or CSV_SHEET_NAME)

    def _load_preview_async(self, file_key: int, file_path: str, sheet_name: str | None) -> None:
        encoding = self.selected_encoding()
        self.preview_request_ids[file_key] += 1
        request_id = self.preview_request_ids[file_key]
        worker = PreviewWorker(file_key, request_id, file_path, sheet_name, encoding)
        worker.setParent(self)
        worker.loaded.connect(self._on_preview_loaded)
        worker.failed.connect(self._on_preview_failed)
        worker.finished.connect(
            lambda file_key=file_key, worker=worker: self._cleanup_preview_worker(file_key, worker)
        )
        self.preview_workers[file_key] = worker
        worker.start()

    def _on_preview_loaded(self, file_key: int, request_id: int, preview: PreviewData) -> None:
        state = self.file_states[file_key]
        if request_id != self.preview_request_ids[file_key]:
            return

        state.sheet_names = preview.sheet_names
        state.selected_sheet = preview.selected_sheet
        state.columns = preview.columns
        state.total_rows = preview.total_rows
        state.used_encoding = preview.used_encoding

        widgets = self.file_widgets[file_key]
        sheet_combo = widgets["sheet_combo"]
        sheet_combo.blockSignals(True)
        sheet_combo.clear()
        sheet_combo.addItems(preview.sheet_names)
        sheet_combo.setCurrentText(preview.selected_sheet)
        sheet_combo.setEnabled(len(preview.sheet_names) > 1)
        sheet_combo.blockSignals(False)

        from os.path import basename

        basename_text = basename(state.path)
        info = f"{basename_text} · {preview.total_rows:,} 行 · {len(preview.columns)} 列"
        if state.path.lower().endswith(".csv"):
            info += f" · 当前编码 {preview.used_encoding.upper()}"
        widgets["info_label"].setText(info)
        widgets["preview_empty"].setVisible(False)
        widgets["preview_table"].setVisible(True)
        self._update_preview_table(file_key, preview)

        if self.file_states[1].columns and self.file_states[2].columns and not self.match_rows:
            self._add_match_row()
        else:
            self._refresh_match_row_options()

        self._refresh_output_checkboxes()
        self.statusBar().showMessage(f"文件 {file_key} 预览已更新")

    def _on_preview_failed(self, file_key: int, request_id: int, message: str) -> None:
        if request_id != self.preview_request_ids[file_key]:
            return
        widgets = self.file_widgets[file_key]
        widgets["info_label"].setText(f"加载失败: {message}")
        widgets["model"].clear()
        widgets["preview_empty"].setText(f"加载失败：{message}")
        widgets["preview_empty"].setVisible(True)
        widgets["preview_table"].setVisible(False)
        self.statusBar().showMessage(f"文件 {file_key} 加载失败")

    def _update_preview_table(self, file_key: int, preview: PreviewData) -> None:
        widgets = self.file_widgets[file_key]
        model = widgets["model"]
        model.clear()
        model.setHorizontalHeaderLabels([column.display_name for column in preview.columns])
        for row in preview.rows:
            items = [QStandardItem(value) for value in row]
            model.appendRow(items)
        preview_table = widgets["preview_table"]
        preview_table.resizeColumnsToContents()
        header = preview_table.horizontalHeader()
        if len(preview.columns) <= 6:
            header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            header.setStretchLastSection(True)
        else:
            header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
            header.setStretchLastSection(False)

    def _add_match_row(self) -> None:
        row = MatchRowWidget()
        row.removed.connect(self._remove_match_row)
        row.changed.connect(self._refresh_output_checkboxes)
        self.match_rows.append(row)
        self.match_rows_layout.addWidget(row)
        self._refresh_match_row_options()
        self._refresh_output_checkboxes()
        self._update_match_empty_state()

    def _remove_match_row(self, row: MatchRowWidget) -> None:
        if row not in self.match_rows:
            return
        self.match_rows.remove(row)
        row.setParent(None)
        row.deleteLater()
        self._refresh_output_checkboxes()
        self._update_match_empty_state()

    def _update_match_empty_state(self) -> None:
        is_empty = not self.match_rows
        self.match_empty_state.setVisible(is_empty)
        self.match_scroll.setVisible(not is_empty)
        if is_empty:
            self.match_card.setMaximumHeight(158)
        else:
            preferred_height = min(112 + len(self.match_rows) * 52, 240)
            self.match_card.setMaximumHeight(preferred_height)

    def _refresh_match_row_options(self) -> None:
        for row in self.match_rows:
            self._populate_combo(row.left_combo, self.file_states[1].columns, row.current_left_raw())
            self._populate_combo(row.right_combo, self.file_states[2].columns, row.current_right_raw())

    def _populate_combo(self, combo: QComboBox, columns: list[ColumnOption], selected_raw: str) -> None:
        combo.blockSignals(True)
        combo.clear()
        for column in columns:
            combo.addItem(column.display_name, column.raw_name)
        if selected_raw:
            index = combo.findData(selected_raw)
            if index >= 0:
                combo.setCurrentIndex(index)
        elif combo.count():
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _refresh_output_checkboxes(self) -> None:
        preserved = {
            source: {
                checkbox.property("raw_name")
                for checkbox in self.output_checkboxes[source]
                if checkbox.isChecked()
            }
            for source in (LEFT_SOURCE, RIGHT_SOURCE)
        }
        blocked = {
            LEFT_SOURCE: {row.current_left_raw() for row in self.match_rows if row.current_left_raw()},
            RIGHT_SOURCE: {row.current_right_raw() for row in self.match_rows if row.current_right_raw()},
        }

        self._rebuild_checkbox_group(
            LEFT_SOURCE,
            self.file_states[1].columns,
            preserved[LEFT_SOURCE],
            blocked[LEFT_SOURCE],
            right_side=False,
        )
        self._rebuild_checkbox_group(
            RIGHT_SOURCE,
            self.file_states[2].columns,
            preserved[RIGHT_SOURCE],
            blocked[RIGHT_SOURCE],
            right_side=True,
        )

    def _rebuild_checkbox_group(
        self,
        source: str,
        columns: list[ColumnOption],
        checked_raw_names: set[str],
        blocked_raw_names: set[str],
        *,
        right_side: bool,
    ) -> None:
        layout = getattr(self, f"{source}_checkbox_layout")
        self._clear_layout(layout, keep_last_stretch=True)
        self.output_checkboxes[source] = []

        if not columns:
            placeholder = QLabel("加载文件后可选择输出列。")
            placeholder.setObjectName("EmptyState")
            placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(placeholder)
            return

        for column in columns:
            visible_label = column.display_name
            export_label = column.display_name if not right_side else f"{column.display_name} · 被匹配文件"
            checkbox = QCheckBox(visible_label)
            checkbox.setProperty("raw_name", column.raw_name)
            checkbox.setProperty("export_label", export_label)
            checkbox.setMinimumHeight(34)
            checkbox.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            checkbox.setEnabled(column.raw_name not in blocked_raw_names)
            checkbox.setChecked(column.raw_name in checked_raw_names and checkbox.isEnabled())
            checkbox.stateChanged.connect(lambda *_: self.statusBar().showMessage("输出列配置已更新"))
            layout.addWidget(checkbox)
            self.output_checkboxes[source].append(checkbox)

    def _start_export(self) -> None:
        if self.export_worker is not None and self.export_worker.isRunning():
            return
        config = self._build_export_config()
        if config is None:
            return

        worker = ExportWorker(config)
        worker.progress_changed.connect(self._on_export_progress)
        worker.succeeded.connect(self._on_export_succeeded)
        worker.failed.connect(self._on_export_failed)
        worker.cancelled.connect(self._on_export_cancelled)
        self.export_worker = worker

        self.progress_dialog = QProgressDialog("正在准备导出...", "取消", 0, 100, self)
        self.progress_dialog.setWindowTitle("导出中")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setMinimumDuration(0)
        self.progress_dialog.canceled.connect(worker.request_cancel)
        self.progress_dialog.show()

        self.export_button.setEnabled(False)
        self.statusBar().showMessage("开始导出结果")
        worker.start()

    def _build_export_config(self) -> ExportConfig | None:
        left_state = self.file_states[1]
        right_state = self.file_states[2]
        if not left_state.path or not right_state.path:
            QMessageBox.warning(self, "缺少文件", "请先选择两个文件。")
            return None

        match_pairs = []
        seen_left: set[str] = set()
        seen_right: set[str] = set()
        for row in self.match_rows:
            left_raw = row.current_left_raw()
            right_raw = row.current_right_raw()
            if not left_raw or not right_raw:
                continue
            if left_raw in seen_left or right_raw in seen_right:
                QMessageBox.warning(self, "匹配列重复", "同一列不能在匹配规则中重复使用。")
                return None
            seen_left.add(left_raw)
            seen_right.add(right_raw)
            match_pairs.append(
                MatchPair(
                    left_raw=left_raw,
                    right_raw=right_raw,
                    left_label=row.current_left_label(),
                    right_label=row.current_right_label(),
                )
            )

        if not match_pairs:
            QMessageBox.warning(self, "缺少匹配规则", "请至少添加一组匹配列。")
            return None

        output_columns = []
        for checkbox in self.output_checkboxes[LEFT_SOURCE]:
            if checkbox.isChecked():
                output_columns.append(
                    OutputColumn(LEFT_SOURCE, checkbox.property("raw_name"), checkbox.property("export_label"))
                )
        for checkbox in self.output_checkboxes[RIGHT_SOURCE]:
            if checkbox.isChecked():
                output_columns.append(
                    OutputColumn(RIGHT_SOURCE, checkbox.property("raw_name"), checkbox.property("export_label"))
                )

        return ExportConfig(
            left_path=left_state.path,
            right_path=right_state.path,
            left_sheet=left_state.selected_sheet,
            right_sheet=right_state.selected_sheet,
            encoding=self.selected_encoding(),
            match_pairs=match_pairs,
            output_columns=output_columns,
            output_filename=self.output_name_edit.text(),
            dedup_strategy=self.selected_dedup_strategy(),
        )

    def _on_export_progress(self, percent: int, message: str) -> None:
        dialog = self.progress_dialog
        if dialog is not None:
            dialog.setValue(percent)
            dialog.setLabelText(message)
        self.statusBar().showMessage(message)

    def _on_export_succeeded(self, output_path: str) -> None:
        self._finish_export()
        QMessageBox.information(self, "导出完成", f"结果已导出到：\n{output_path}")
        self.statusBar().showMessage(f"导出完成: {output_path}")

    def _on_export_failed(self, message: str) -> None:
        self._finish_export()
        QMessageBox.critical(self, "导出失败", message)
        self.statusBar().showMessage("导出失败")

    def _on_export_cancelled(self) -> None:
        self._finish_export()
        QMessageBox.information(self, "已取消", "导出任务已取消。")
        self.statusBar().showMessage("导出已取消")

    def _finish_export(self) -> None:
        if self.progress_dialog is not None:
            self.progress_dialog.close()
            self.progress_dialog = None
        self.export_button.setEnabled(True)
        self.export_worker = None

    def _cleanup_preview_worker(self, file_key: int, worker: PreviewWorker) -> None:
        if self.preview_workers.get(file_key) is worker:
            self.preview_workers.pop(file_key, None)

    def selected_encoding(self) -> str:
        return "gbk" if self.gbk_radio.isChecked() else "utf-8"

    def selected_dedup_strategy(self) -> str:
        return self.dedup_combo.currentData() or DEDUP_STRATEGY_LEFT_KEY_RIGHT_ROW

    def _clear_layout(self, layout: QVBoxLayout, *, keep_last_stretch: bool) -> None:
        while layout.count() > 0:
            item = layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()


def configure_application(app: QApplication | None) -> None:
    if app is None:
        return
    if app.property("_excel_tool_configured"):
        return

    font = QFont(COMMON_UI_FONT_FAMILY, 11)
    app.setFont(font)
    app.setStyle("Fusion")
    app.setProperty("_excel_tool_configured", True)


def main() -> None:
    app = QApplication(sys.argv)
    configure_application(app)
    window = ExcelCompareTool()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
