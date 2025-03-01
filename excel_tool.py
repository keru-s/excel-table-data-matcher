import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QGroupBox, QTableWidget, QTableWidgetItem, QFileDialog,
                             QComboBox, QCheckBox, QProgressDialog, QMessageBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import os

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
        self.file1_checkbox_layout = QVBoxLayout()
        file1_checkbox_group.setLayout(self.file1_checkbox_layout)
        
        # 文件2的多选框组
        file2_checkbox_group = QGroupBox("被匹配文件输出列")
        self.file2_checkbox_layout = QVBoxLayout()
        file2_checkbox_group.setLayout(self.file2_checkbox_layout)
        
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
            "Excel files (*.xlsx)"
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
        # 更新文件1的下拉列表选项
        for row in self.match_rows:
            row[0].clear()
            row[0].addItems(columns)
        
        # 更新文件1的输出列多选框
        for checkbox in self.file1_checkboxes:
            self.file1_checkbox_layout.removeWidget(checkbox)
            checkbox.deleteLater()
        self.file1_checkboxes.clear()
        
        for column in columns:
            checkbox = QCheckBox(column)
            self.file1_checkboxes.append(checkbox)
            self.file1_checkbox_layout.addWidget(checkbox)
    
    def update_file2_columns(self, columns):
        # 更新文件2的下拉列表选项
        for row in self.match_rows:
            row[1].clear()
            row[1].addItems(columns)
        
        # 更新文件2的输出列多选框
        for checkbox in self.file2_checkboxes:
            self.file2_checkbox_layout.removeWidget(checkbox)
            checkbox.deleteLater()
        self.file2_checkboxes.clear()
        
        for column in columns:
            checkbox = QCheckBox(column)
            self.file2_checkboxes.append(checkbox)
            self.file2_checkbox_layout.addWidget(checkbox)
    
    def add_match_row(self):
        # 创建一行匹配选择
        row_layout = QHBoxLayout()
        combo1 = QComboBox()
        combo2 = QComboBox()
        
        # 如果已经有文件1的数据，添加到下拉列表
        if self.preview1.columnCount() > 0:
            headers = [self.preview1.horizontalHeaderItem(i).text() 
                      for i in range(self.preview1.columnCount())]
            combo1.addItems(headers)
        
        # 如果已经有文件2的数据，添加到下拉列表
        if self.preview2.columnCount() > 0:
            headers = [self.preview2.horizontalHeaderItem(i).text() 
                      for i in range(self.preview2.columnCount())]
            combo2.addItems(headers)
        
        row_layout.addWidget(combo1)
        row_layout.addWidget(combo2)
        
        # 将下拉列表添加到匹配行列表中
        self.match_rows.append([combo1, combo2])
        
        # 将行布局添加到匹配行容器中
        self.match_rows_layout.addLayout(row_layout)
    
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
            df1 = pd.read_excel(self.file1_path, engine='openpyxl', sheet_name=self.sheet1_name)
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
            output_path = os.path.join(output_dir, f"{self.output_filename}.xlsx")
            
            # 导出结果
            result.to_excel(output_path, index=False)
            
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