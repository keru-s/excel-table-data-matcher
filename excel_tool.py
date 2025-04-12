import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QGroupBox, QTableWidget, QTableWidgetItem, QFileDialog,
                             QComboBox, QCheckBox, QProgressDialog, QMessageBox,
                             QScrollArea)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import os
import xlrd
import chardet
import datetime

class ExcelCompareTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel文件处理工具")
        self.setGeometry(100, 100, 900, 340)
        
        # 创建中心部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(5)  # 减小主布局间距
        main_layout.setContentsMargins(5, 5, 5, 5)  # 设置主布局边距
        
        # 文件选择区域
        self.file1_path = QLineEdit()
        self.file2_path = QLineEdit()
        
        # 创建水平布局来容纳文件选择组件
        file_selection_layout = QHBoxLayout()
        file_selection_layout.setSpacing(10)  # 减小文件选择框之间的间距
        file_selection_layout.setContentsMargins(0, 0, 0, 0)  # 移除文件选择布局的边距
        
        # 创建文件选择组件
        file1_group = self.create_file_group("匹配源文件", self.file1_path, 1)
        file2_group = self.create_file_group("被匹配文件", self.file2_path, 2)
        
        file_selection_layout.addWidget(file1_group)
        file_selection_layout.addWidget(file2_group)
        main_layout.addLayout(file_selection_layout)
        
        # 预览区域
        preview_group = QGroupBox("数据预览")
        preview_layout = QHBoxLayout()
        preview_layout.setSpacing(2)  # 进一步减小水平布局的间距
        preview_layout.setContentsMargins(2, 2, 2, 2)  # 设置更小的边距
        
        # 创建两个垂直布局来容纳标签和表格
        preview1_layout = QVBoxLayout()
        preview1_layout.setSpacing(0)  # 移除垂直布局的间距
        preview2_layout = QVBoxLayout()
        preview2_layout.setSpacing(0)  # 移除垂直布局的间距
        
        # 创建文件名标签并设置样式
        self.preview1_label = QLabel("未选择文件")
        self.preview2_label = QLabel("未选择文件")
        self.preview1_label.setContentsMargins(0, 0, 0, 2)  # 设置标签底部小边距
        self.preview2_label.setContentsMargins(0, 0, 0, 2)  # 设置标签底部小边距
        
        # 创建表格预览组件
        self.preview1 = QTableWidget()
        self.preview2 = QTableWidget()
        
        # 设置表格的固定大小和滚动条策略
        for preview in [self.preview1, self.preview2]:
            preview.setFixedHeight(200)  # 减小固定高度
            preview.setFixedWidth(450)   # 设置固定宽度
            preview.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            preview.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            preview.setContentsMargins(0, 0, 0, 0)  # 移除表格的内边距
        # 将标签和表格添加到各自的垂直布局中
        preview1_layout.addWidget(self.preview1_label)
        preview1_layout.addWidget(self.preview1)
        preview2_layout.addWidget(self.preview2_label)
        preview2_layout.addWidget(self.preview2)
        
        # 将垂直布局添加到预览布局中
        preview_layout.addLayout(preview1_layout)
        preview_layout.addLayout(preview2_layout)
        
        preview_group.setLayout(preview_layout)
        main_layout.addWidget(preview_group)
        
        # 列匹配设置区域
        match_group = QGroupBox("列匹配设置")
        match_layout = QVBoxLayout()
        match_layout.setSpacing(5)
        match_layout.setContentsMargins(5, 5, 5, 5)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        add_button = QPushButton("增加")
        add_button.clicked.connect(self.add_match_row)
        delete_button = QPushButton("删除")
        delete_button.clicked.connect(self.delete_match_row)
        button_layout.addWidget(add_button)
        button_layout.addWidget(delete_button)
        button_layout.addStretch()
        match_layout.addLayout(button_layout)
        
        # 匹配行容器
        self.match_rows_widget = QWidget()
        self.match_rows_layout = QVBoxLayout(self.match_rows_widget)
        self.match_rows_layout.setSpacing(5)
        self.match_rows_layout.setContentsMargins(0, 0, 0, 0)
        match_layout.addWidget(self.match_rows_widget)
        
        match_group.setLayout(match_layout)
        main_layout.addWidget(match_group)
        
        # 输出结果设置区域
        output_group = QGroupBox("输出结果设置")
        output_layout = QVBoxLayout()
        output_layout.setSpacing(5)
        output_layout.setContentsMargins(5, 5, 5, 5)
        
        # 多选框区域
        checkboxes_layout = QHBoxLayout()
        
        # 文件1的多选框组
        file1_checkbox_group = QGroupBox("匹配源文件输出列")
        file1_checkbox_container = QWidget()
        self.file1_checkbox_layout = QVBoxLayout(file1_checkbox_container)
        file1_scroll = QScrollArea()
        file1_scroll.setWidget(file1_checkbox_container)
        file1_scroll.setWidgetResizable(True)
        file1_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        file1_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        file1_scroll.setMaximumHeight(450)
        file1_group_layout = QVBoxLayout(file1_checkbox_group)
        file1_group_layout.addWidget(file1_scroll)
        
        # 文件2的多选框组
        file2_checkbox_group = QGroupBox("被匹配文件输出列")
        file2_checkbox_container = QWidget()
        self.file2_checkbox_layout = QVBoxLayout(file2_checkbox_container)
        file2_scroll = QScrollArea()
        file2_scroll.setWidget(file2_checkbox_container)
        file2_scroll.setWidgetResizable(True)
        file2_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        file2_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        file2_scroll.setMaximumHeight(450)
        file2_group_layout = QVBoxLayout(file2_checkbox_group)
        file2_group_layout.addWidget(file2_scroll)
        
        checkboxes_layout.addWidget(file1_checkbox_group)
        checkboxes_layout.addWidget(file2_checkbox_group)
        output_layout.addLayout(checkboxes_layout)
        
        # 输出文件名设置
        filename_layout = QHBoxLayout()
        filename_layout.addWidget(QLabel("输出文件名:"))
        self.output_filename = QLineEdit("匹配结果")
        filename_layout.addWidget(self.output_filename)
        filename_layout.addWidget(QLabel(".xlsx"))
        output_layout.addLayout(filename_layout)
        
        # 添加导出按钮
        export_button = QPushButton("导出结果")
        export_button.clicked.connect(self.export_result)
        output_layout.addWidget(export_button)
        
        output_group.setLayout(output_layout)
        main_layout.addWidget(output_group)
        
        # 初始化匹配行和复选框列表
        self.match_rows = []
        self.file1_checkboxes = []
        self.file2_checkboxes = []
        
        # 添加默认的匹配行
        self.add_match_row()
    
    def create_file_group(self, title, line_edit, file_num):
        group = QGroupBox(title)
        layout = QHBoxLayout()
        
        line_edit.setReadOnly(True)
        layout.addWidget(line_edit)
        
        select_button = QPushButton("选择文件")
        select_button.clicked.connect(lambda: self.select_file(file_num))
        layout.addWidget(select_button)
        
        # 添加sheet选择下拉列表
        sheet_combo = QComboBox()
        sheet_combo.setEnabled(False)  # 初始禁用
        sheet_combo.currentIndexChanged.connect(lambda: self.on_sheet_changed(file_num))
        layout.addWidget(sheet_combo)
        
        # 保存sheet选择下拉列表的引用
        if file_num == 1:
            self.sheet1_combo = sheet_combo
        else:
            self.sheet2_combo = sheet_combo
        
        group.setLayout(layout)
        return group
    
    def select_file(self, file_num):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"选择Excel文件{file_num}",
            "",
            "Excel files (*.xlsx *.xls)"
        )
        
        if file_path:
            if file_num == 1:
                self.file1_path.setText(file_path)
                self.preview1_label.setText(f"匹配源文件: {file_path.split('/')[-1]}")
                self.show_preview(file_path, self.preview1)
            else:
                self.file2_path.setText(file_path)
                self.preview2_label.setText(f"被匹配文件: {file_path.split('/')[-1]}")
                self.show_preview(file_path, self.preview2)
    
    def on_sheet_changed(self, file_num):
        file_path = self.file1_path.text() if file_num == 1 else self.file2_path.text()
        sheet_combo = self.sheet1_combo if file_num == 1 else self.sheet2_combo
        preview_table = self.preview1 if file_num == 1 else self.preview2
        
        if file_path and sheet_combo.currentText():
            self.show_preview(file_path, preview_table, sheet_combo.currentText())
    
    def show_preview(self, file_path, preview_table, sheet_name=None):
        try:
            # 检查文件类型
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext == '.xls':
                # 使用xlrd处理xls文件
                # 尝试使用常见的中文编码
                encodings = ['gb18030', 'gbk', 'gb2312', 'utf-8']
                workbook = None
                
                for encoding in encodings:
                    try:
                        workbook = xlrd.open_workbook(file_path, encoding_override=encoding)
                        # 尝试读取一个单元格的内容来验证编码是否正确
                        sheet = workbook.sheet_by_index(0)
                        test_value = str(sheet.cell_value(0, 0))
                        if not any(c == '?' or c == '�' or ord(c) > 0xffff for c in test_value):
                            break
                    except Exception:
                        continue
                
                if not workbook:
                    # 如果所有编码都失败，使用默认编码
                    workbook = xlrd.open_workbook(file_path)
                
                # 如果是新选择的文件，更新sheet列表
                if not sheet_name:
                    sheet_names = workbook.sheet_names()
                    
                    # 更新对应的sheet下拉列表
                    sheet_combo = self.sheet1_combo if preview_table == self.preview1 else self.sheet2_combo
                    sheet_combo.clear()
                    sheet_combo.addItems(sheet_names)
                    sheet_combo.setEnabled(True)
                    sheet_name = sheet_names[0]  # 默认选择第一个sheet
                
                # 读取指定sheet的前5行
                sheet = workbook.sheet_by_name(sheet_name)
                
                # 创建DataFrame
                data = []
                headers = []
                
                # 获取表头
                for col in range(sheet.ncols):
                    headers.append(str(sheet.cell_value(0, col)))
                
                # 获取数据（前5行）
                for row in range(1, min(6, sheet.nrows)):
                    row_data = []
                    for col in range(sheet.ncols):
                        cell_value = sheet.cell_value(row, col)
                        # 处理日期类型
                        if sheet.cell_type(row, col) == xlrd.XL_CELL_DATE:
                            try:
                                cell_value = xlrd.xldate.xldate_as_datetime(cell_value, workbook.datemode).strftime('%Y-%m-%d')
                            except:
                                pass
                        row_data.append(cell_value)
                    data.append(row_data)
                
                # 创建DataFrame
                df = pd.DataFrame(data, columns=headers)
            else:
                # 使用openpyxl读取Excel文件
                # 如果是新选择的文件，更新sheet列表
                if not sheet_name:
                    # 使用openpyxl读取
                    excel_file = pd.ExcelFile(file_path, engine='openpyxl')
                    sheet_names = excel_file.sheet_names
                    
                    # 更新对应的sheet下拉列表
                    sheet_combo = self.sheet1_combo if preview_table == self.preview1 else self.sheet2_combo
                    sheet_combo.clear()
                    sheet_combo.addItems(sheet_names)
                    sheet_combo.setEnabled(True)
                    sheet_name = sheet_names[0]  # 默认选择第一个sheet
                
                # 读取指定sheet的前5行
                df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name, nrows=5)
            
            # 设置表格的行数和列数
            preview_table.setRowCount(len(df.index))
            preview_table.setColumnCount(len(df.columns))
            
            # 设置表头（Excel列标识与列名拼接）
            excel_columns = [chr(65 + i) if i < 26 else chr(64 + i//26) + chr(65 + i%26) for i in range(len(df.columns))]
            header_labels = [f"{col}-{str(val)}" for col, val in zip(excel_columns, df.columns)]
            preview_table.setHorizontalHeaderLabels(header_labels)
            
            # 填充数据
            for i in range(len(df.index)):
                for j in range(len(df.columns)):
                    item = QTableWidgetItem(str(df.iloc[i, j]))
                    preview_table.setItem(i, j, item)
            
            # 调整列宽以适应内容
            preview_table.resizeColumnsToContents()
            preview_table.resizeRowsToContents()
            
            # 更新列匹配下拉列表和输出列多选框
            if preview_table == self.preview1:
                self.update_file1_columns(header_labels)
            else:
                self.update_file2_columns(header_labels)
            
        except Exception as e:
            # 清空表格并显示错误信息
            preview_table.setRowCount(0)
            preview_table.setColumnCount(1)
            preview_table.setHorizontalHeaderLabels(["错误"])
            preview_table.setRowCount(1)
            preview_table.setItem(0, 0, QTableWidgetItem(str(e)))

    def update_file1_columns(self, columns):
        # 获取已选择的匹配列
        selected_columns1 = []
        for row in self.match_rows:
            text1 = row[0].currentText()
            if text1:
                selected_columns1.append(text1)
        
        # 更新文件1的下拉列表选项
        for row in self.match_rows:
            row[0].clear()
            row[0].addItems(columns)
        
        # 更新文件1的输出列多选框
        for checkbox in self.file1_checkboxes:
            self.file1_checkbox_layout.removeWidget(checkbox)
            checkbox.deleteLater()
        self.file1_checkboxes.clear()
        
        # 只添加未被选为匹配列的列到输出多选框
        for column in columns:
            if column not in selected_columns1:
                checkbox = QCheckBox(column)
                self.file1_checkboxes.append(checkbox)
                self.file1_checkbox_layout.addWidget(checkbox)

    def update_file2_columns(self, columns):
        # 获取已选择的匹配列
        selected_columns2 = []
        for row in self.match_rows:
            text2 = row[1].currentText()
            if text2:
                selected_columns2.append(text2)
        
        # 更新文件2的下拉列表选项
        for row in self.match_rows:
            row[1].clear()
            row[1].addItems(columns)
        
        # 更新文件2的输出列多选框
        for checkbox in self.file2_checkboxes:
            self.file2_checkbox_layout.removeWidget(checkbox)
            checkbox.deleteLater()
        self.file2_checkboxes.clear()
        
        # 只添加未被选为匹配列的列到输出多选框
        for column in columns:
            if column not in selected_columns2:
                checkbox = QCheckBox(column)
                self.file2_checkboxes.append(checkbox)
                self.file2_checkbox_layout.addWidget(checkbox)
    
    def add_match_row(self):
        # 创建一行匹配选择
        row_widget = QWidget()
        row_layout = QHBoxLayout(row_widget)
        combo1 = QComboBox()
        combo2 = QComboBox()
        
        # 获取已选择的列
        selected_columns1 = []
        selected_columns2 = []
        for row in self.match_rows:
            text1 = row[0].currentText()
            text2 = row[1].currentText()
            if text1:
                selected_columns1.append(text1)
            if text2:
                selected_columns2.append(text2)
        
        # 如果已经有文件1的数据，添加未被选择的列到下拉列表
        if self.preview1.columnCount() > 0:
            headers = [self.preview1.horizontalHeaderItem(i).text() 
                      for i in range(self.preview1.columnCount())]
            available_headers = [h for h in headers if h not in selected_columns1]
            combo1.addItems(available_headers)
            if available_headers:  # 如果有可用的列，设置第一个为默认值
                combo1.setCurrentText(available_headers[0])
        
        # 如果已经有文件2的数据，添加未被选择的列到下拉列表
        if self.preview2.columnCount() > 0:
            headers = [self.preview2.horizontalHeaderItem(i).text() 
                      for i in range(self.preview2.columnCount())]
            available_headers = [h for h in headers if h not in selected_columns2]
            combo2.addItems(available_headers)
            if available_headers:  # 如果有可用的列，设置第一个为默认值
                combo2.setCurrentText(available_headers[0])
        
        # 添加选择变化事件处理
        combo1.currentIndexChanged.connect(lambda index: self.on_match_column_changed())
        combo2.currentIndexChanged.connect(lambda index: self.on_match_column_changed())
        
        row_layout.addWidget(combo1)
        row_layout.addWidget(combo2)
        
        # 将下拉列表添加到匹配行列表中
        self.match_rows.append([combo1, combo2])
        
        # 将行布局添加到匹配行容器中
        self.match_rows_layout.addWidget(row_widget)
        
        # 更新所有下拉列表的可选项
        self.update_available_columns()
        
        # 手动触发更新输出列多选框
        self.on_match_column_changed()

    def on_match_column_changed(self):
        # 更新所有下拉列表的可选项
        self.update_available_columns()
        
        # 获取当前所有已选择的匹配列
        selected_columns1 = []
        selected_columns2 = []
        for row in self.match_rows:
            text1 = row[0].currentText()
            text2 = row[1].currentText()
            if text1:
                selected_columns1.append(text1)
            if text2:
                selected_columns2.append(text2)
        
        # 更新文件1的输出列多选框
        for checkbox in self.file1_checkboxes:
            self.file1_checkbox_layout.removeWidget(checkbox)
            checkbox.deleteLater()
        self.file1_checkboxes.clear()
        
        # 获取文件1的所有列
        if self.preview1.columnCount() > 0:
            columns1 = [self.preview1.horizontalHeaderItem(i).text() 
                       for i in range(self.preview1.columnCount())]
            # 只添加未被选为匹配列的列到输出多选框
            for column in columns1:
                if column not in selected_columns1:
                    checkbox = QCheckBox(column)
                    self.file1_checkboxes.append(checkbox)
                    self.file1_checkbox_layout.addWidget(checkbox)
        
        # 更新文件2的输出列多选框
        for checkbox in self.file2_checkboxes:
            self.file2_checkbox_layout.removeWidget(checkbox)
            checkbox.deleteLater()
        self.file2_checkboxes.clear()
        
        # 获取文件2的所有列
        if self.preview2.columnCount() > 0:
            columns2 = [self.preview2.horizontalHeaderItem(i).text() 
                       for i in range(self.preview2.columnCount())]
            # 只添加未被选为匹配列的列到输出多选框
            for column in columns2:
                if column not in selected_columns2:
                    checkbox = QCheckBox(column)
                    self.file2_checkboxes.append(checkbox)
                    self.file2_checkbox_layout.addWidget(checkbox)

    def update_available_columns(self):
        # 获取所有可用列
        all_columns1 = []
        all_columns2 = []
        
        if self.preview1.columnCount() > 0:
            all_columns1 = [self.preview1.horizontalHeaderItem(i).text() 
                           for i in range(self.preview1.columnCount())]
        if self.preview2.columnCount() > 0:
            all_columns2 = [self.preview2.horizontalHeaderItem(i).text() 
                           for i in range(self.preview2.columnCount())]
        
        # 获取所有已选择的列
        selected_columns1 = []
        selected_columns2 = []
        
        # 先收集所有已选择的列
        for combo1, combo2 in self.match_rows:
            text1 = combo1.currentText()
            text2 = combo2.currentText()
            if text1:
                selected_columns1.append(text1)
            if text2:
                selected_columns2.append(text2)
        
        # 更新每个下拉列表
        for combo1, combo2 in self.match_rows:
            current1 = combo1.currentText()
            current2 = combo2.currentText()
            
            # 暂时断开信号连接
            combo1.blockSignals(True)
            combo2.blockSignals(True)
            
            # 更新文件1的下拉列表
            combo1.clear()
            available_columns1 = [col for col in all_columns1 if col not in selected_columns1 or col == current1]
            combo1.addItems(available_columns1)
            if current1 in available_columns1:
                combo1.setCurrentText(current1)
            
            # 更新文件2的下拉列表
            combo2.clear()
            available_columns2 = [col for col in all_columns2 if col not in selected_columns2 or col == current2]
            combo2.addItems(available_columns2)
            if current2 in available_columns2:
                combo2.setCurrentText(current2)
            
            # 恢复信号连接
            combo1.blockSignals(False)
            combo2.blockSignals(False)
    
    def delete_match_row(self):
        if self.match_rows:
            # 获取最后一行的组件
            last_row = self.match_rows.pop()
            layout = last_row[0].parent().layout()
            
            # 删除组件
            for combo in last_row:
                layout.removeWidget(combo)
                combo.deleteLater()
            
            # 删除布局
            self.match_rows_layout.removeItem(layout)
            layout.deleteLater()
            
            # 更新所有下拉列表的可选项
            self.update_available_columns()

    def on_export_finished(self, output_path, loading):
        loading.close()
        QMessageBox.information(self, "导出成功", f"数据已成功匹配并导出到：\n{output_path}")

    def on_export_error(self, error_msg, loading):
        loading.close()
        QMessageBox.critical(self, "导出错误", f"导出过程中发生错误：\n{error_msg}")

    def export_result(self):
        if not self.file1_path.text() or not self.file2_path.text():
            return
        
        # 创建带动画的进度对话框
        loading = QProgressDialog(self)
        loading.setWindowTitle("处理中")
        loading.setLabelText("正在匹配数据，请稍候...")
        loading.setRange(0, 0)  # 设置为循环进度条
        loading.setWindowModality(Qt.WindowModality.WindowModal)
        loading.setMinimumDuration(0)  # 立即显示
        loading.setCancelButton(None)  # 移除取消按钮
        loading.setStyleSheet("""
            QProgressDialog {
                background-color: #f5f5f5;
                border-radius: 8px;
                padding: 20px;
            }
            QProgressDialog QLabel {
                color: #333333;
                font-size: 14px;
                margin-bottom: 10px;
            }
            QProgressBar {
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                text-align: center;
                background-color: #ffffff;
            }
            QProgressBar::chunk {
                background-color: #4a90e2;
                border-radius: 3px;
            }
        """)
        loading.show()
        
        # 获取匹配列
        match_columns = []
        for combo1, combo2 in self.match_rows:
            if combo1.currentText() and combo2.currentText():
                match_columns.append((combo1.currentText(), combo2.currentText()))
        
        if not match_columns:
            loading.close()
            return
        
        # 获取选中的输出列
        output_columns = []
        for checkbox in self.file1_checkboxes:
            if checkbox.isChecked():
                output_columns.append(checkbox.text().split('-')[1])
        for checkbox in self.file2_checkboxes:
            if checkbox.isChecked():
                col_name = checkbox.text().split('-')[1]
                if col_name not in output_columns:  # 避免重复列
                    output_columns.append(col_name)
        
        # 创建工作线程
        self.worker = ExportWorker(
            self.file1_path.text(),
            self.file2_path.text(),
            self.sheet1_combo.currentText(),
            self.sheet2_combo.currentText(),
            match_columns,
            output_columns,
            self.output_filename.text()
        )
        
        # 连接信号
        self.worker.finished.connect(lambda output_path: self.on_export_finished(output_path, loading))
        self.worker.error.connect(lambda error_msg: self.on_export_error(error_msg, loading))
        
        # 启动工作线程
        self.worker.start()


class ExportWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, file1_path, file2_path, sheet1_name, sheet2_name, 
                 match_columns, output_columns, output_filename):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.sheet1_name = sheet1_name
        self.sheet2_name = sheet2_name
        self.match_columns = match_columns
        self.output_columns = output_columns
        self.output_filename = output_filename
    
    def run(self):
        try:
            # 读取Excel文件
            # 根据文件扩展名选择合适的引擎
            file1_ext = os.path.splitext(self.file1_path)[1].lower()
            file2_ext = os.path.splitext(self.file2_path)[1].lower()
            
            # 读取第一个文件
            if file1_ext == '.xls':
                # 使用xlrd处理xls文件
                # 尝试使用常见的中文编码
                encodings = ['gb18030', 'gbk', 'gb2312', 'utf-8']
                workbook1 = None
                
                for encoding in encodings:
                    try:
                        workbook1 = xlrd.open_workbook(self.file1_path, encoding_override=encoding)
                        # 尝试读取一个单元格的内容来验证编码是否正确
                        sheet1 = workbook1.sheet_by_index(0)
                        test_value = str(sheet1.cell_value(0, 0))
                        if not any(c == '?' or c == '�' or ord(c) > 0xffff for c in test_value):
                            break
                    except Exception:
                        continue
                
                if not workbook1:
                    # 如果所有编码都失败，使用默认编码
                    workbook1 = xlrd.open_workbook(self.file1_path)
                
                sheet1 = workbook1.sheet_by_name(self.sheet1_name)
                
                # 创建DataFrame
                data1 = []
                headers1 = []
                
                # 获取表头
                for col in range(sheet1.ncols):
                    headers1.append(str(sheet1.cell_value(0, col)))
                
                # 获取数据
                for row in range(1, sheet1.nrows):
                    row_data = []
                    for col in range(sheet1.ncols):
                        cell_value = sheet1.cell_value(row, col)
                        # 处理日期类型
                        if sheet1.cell_type(row, col) == xlrd.XL_CELL_DATE:
                            try:
                                cell_value = xlrd.xldate.xldate_as_datetime(cell_value, workbook1.datemode).strftime('%Y-%m-%d')
                            except:
                                pass
                        row_data.append(cell_value)
                    data1.append(row_data)
                
                df1 = pd.DataFrame(data1, columns=headers1)
            else:
                df1 = pd.read_excel(self.file1_path, engine='openpyxl', sheet_name=self.sheet1_name)
            
            # 读取第二个文件
            if file2_ext == '.xls':
                # 使用xlrd处理xls文件
                # 尝试使用常见的中文编码
                encodings = ['gb18030', 'gbk', 'gb2312', 'utf-8']
                workbook2 = None
                
                for encoding in encodings:
                    try:
                        workbook2 = xlrd.open_workbook(self.file2_path, encoding_override=encoding)
                        # 尝试读取一个单元格的内容来验证编码是否正确
                        sheet2 = workbook2.sheet_by_index(0)
                        test_value = str(sheet2.cell_value(0, 0))
                        if not any(c == '?' or c == '�' or ord(c) > 0xffff for c in test_value):
                            break
                    except Exception:
                        continue
                
                if not workbook2:
                    # 如果所有编码都失败，使用默认编码
                    workbook2 = xlrd.open_workbook(self.file2_path)
                
                sheet2 = workbook2.sheet_by_name(self.sheet2_name)
                
                # 创建DataFrame
                data2 = []
                headers2 = []
                
                # 获取表头
                for col in range(sheet2.ncols):
                    headers2.append(str(sheet2.cell_value(0, col)))
                
                # 获取数据
                for row in range(1, sheet2.nrows):
                    row_data = []
                    for col in range(sheet2.ncols):
                        cell_value = sheet2.cell_value(row, col)
                        # 处理日期类型
                        if sheet2.cell_type(row, col) == xlrd.XL_CELL_DATE:
                            try:
                                cell_value = xlrd.xldate.xldate_as_datetime(cell_value, workbook2.datemode).strftime('%Y-%m-%d')
                            except:
                                pass
                        row_data.append(cell_value)
                    data2.append(row_data)
                
                df2 = pd.DataFrame(data2, columns=headers2)
            else:
                df2 = pd.read_excel(self.file2_path, engine='openpyxl', sheet_name=self.sheet2_name)
            
            # 创建匹配条件
            merge_conditions = []
            rename_dict = {}
            merge_on = []
            
            for col1, col2 in self.match_columns:
                col1_name = col1.split('-')[1]
                col2_name = col2.split('-')[1]
                merge_conditions.append(f"{col1_name} == {col2_name}")
                rename_dict[col2_name] = col1_name
                merge_on.append(col1_name)
            
            # 重命名df2的列以匹配df1
            df2 = df2.rename(columns=rename_dict)
            
            # 合并数据
            result = pd.merge(df1, df2, on=merge_on)
            
            # 合并匹配列和输出列
            final_columns = merge_on + [col for col in self.output_columns if col not in merge_on]
            
            # 选择最终输出的列
            result = result[final_columns]
            
            # 获取输出文件路径
            output_dir = os.path.dirname(self.file2_path)
            base_output_path = os.path.join(output_dir, f"{self.output_filename}.xlsx")
            
            # 检查是否存在同名文件，如果存在则添加时间后缀
            if os.path.exists(base_output_path):
                # 获取当前时间作为后缀
                current_time = datetime.datetime.now().strftime("%m%d_%H%M%S")  # 修改时间格式
                output_path = os.path.join(output_dir, f"{self.output_filename}_{current_time}.xlsx")
            else:
                output_path = base_output_path
            
            # 确保路径使用正确的分隔符
            output_path = os.path.normpath(output_path)
            
            # 导出结果
            result.to_excel(output_path, index=False, engine='openpyxl')  # 明确指定引擎
            
            self.finished.emit(output_path)
            
        except Exception as e:
            self.error.emit(str(e))

def main():
    app = QApplication(sys.argv)
    window = ExcelCompareTool()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()