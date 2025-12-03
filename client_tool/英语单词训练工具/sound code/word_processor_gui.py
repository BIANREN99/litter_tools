import sys
import os
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                           QMessageBox, QProgressBar, QGroupBox, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# 导入现有的单词处理器类
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from slove_tool import WordProcessor

class ProcessingThread(QThread):
    """处理文件的线程，避免UI卡顿"""
    
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, processor, input_file, output_dir, formats):
        super().__init__()
        self.processor = processor
        self.input_file = input_file
        self.output_dir = output_dir
        self.formats = formats
    
    def run(self):
        try:
            # 加载文件
            self.progress.emit(10)
            if not self.processor.load_from_file(self.input_file):
                self.finished.emit(False, "加载文件失败，请检查文件格式")
                return
            
            self.progress.emit(30)
            
            # 创建输出目录
            output_path = Path(self.output_dir)
            output_path.mkdir(exist_ok=True)
            
            # 转换为不同格式
            success_count = 0
            total_formats = len(self.formats)
            
            if 'markdown' in self.formats:
                self.processor.to_markdown(output_path / 'words.md')
                success_count += 1
                self.progress.emit(30 + 70 * success_count // total_formats)
            
            if 'csv' in self.formats:
                self.processor.to_csv(output_path / 'words.csv')
                success_count += 1
                self.progress.emit(30 + 70 * success_count // total_formats)
            
            if 'json' in self.formats:
                self.processor.to_json(output_path / 'words.json')
                success_count += 1
                self.progress.emit(30 + 70 * success_count // total_formats)
            
            if 'txt' in self.formats:
                self.processor.to_txt(output_path / 'words.txt')
                success_count += 1
                self.progress.emit(100)
            
            self.finished.emit(True, f"成功转换 {success_count} 个格式的文件！")
        except Exception as e:
            self.finished.emit(False, f"处理过程中出错: {str(e)}")

class WordProcessorGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        # 先初始化属性，再调用init_ui
        self.processor = WordProcessor()
        self.input_file = ""
        self.output_dir = os.path.join(os.getcwd(), "output")
        self.answer_shown = False  # 添加这个属性以避免潜在的错误
        # 然后再调用init_ui
        self.init_ui()
    
    def init_ui(self):
        # 设置窗口标题和大小
        self.setWindowTitle("单词格式转换器")
        self.setGeometry(100, 100, 600, 400)
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        # 创建文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # 输入文件选择
        input_layout = QHBoxLayout()
        self.input_file_label = QLabel("未选择文件")
        self.input_file_label.setStyleSheet("color: gray")
        self.input_file_label.setAlignment(Qt.AlignCenter)
        self.input_file_label.setMinimumHeight(30)
        
        browse_button = QPushButton("浏览文件")
        browse_button.clicked.connect(self.browse_input_file)
        browse_button.setMinimumWidth(100)
        
        input_layout.addWidget(self.input_file_label, 1)
        input_layout.addWidget(browse_button)
        
        # 输出目录选择
        output_layout = QHBoxLayout()
        self.output_dir_label = QLabel(f"输出目录: {self.output_dir}")
        self.output_dir_label.setStyleSheet("color: blue")
        self.output_dir_label.setWordWrap(True)
        
        output_button = QPushButton("更改目录")
        output_button.clicked.connect(self.browse_output_dir)
        output_button.setMinimumWidth(100)
        
        output_layout.addWidget(self.output_dir_label, 1)
        output_layout.addWidget(output_button)
        
        file_layout.addLayout(input_layout)
        file_layout.addLayout(output_layout)
        file_group.setLayout(file_layout)
        
        # 创建格式选择区域
        format_group = QGroupBox("输出格式选择")
        format_layout = QHBoxLayout()
        
        self.md_checkbox = QCheckBox("Markdown (.md)")
        self.md_checkbox.setChecked(True)
        
        self.csv_checkbox = QCheckBox("CSV (.csv)")
        self.csv_checkbox.setChecked(True)
        
        self.json_checkbox = QCheckBox("JSON (.json)")
        self.json_checkbox.setChecked(True)
        
        self.txt_checkbox = QCheckBox("文本 (.txt)")
        self.txt_checkbox.setChecked(True)
        
        format_layout.addWidget(self.md_checkbox)
        format_layout.addWidget(self.csv_checkbox)
        format_layout.addWidget(self.json_checkbox)
        format_layout.addWidget(self.txt_checkbox)
        format_group.setLayout(format_layout)
        
        # 创建处理按钮
        self.process_button = QPushButton("开始转换")
        self.process_button.setMinimumHeight(40)
        self.process_button.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px;")
        self.process_button.clicked.connect(self.start_processing)
        
        # 创建进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        
        # 添加到主布局
        main_layout.addWidget(file_group)
        main_layout.addWidget(format_group)
        main_layout.addWidget(self.process_button)
        main_layout.addWidget(self.progress_bar)
        
        # 添加状态标签
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: green; font-weight: bold")
        main_layout.addWidget(self.status_label)
        
    def browse_input_file(self):
        """浏览并选择输入文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择单词文件", "", "文本文件 (*.txt *.csv *.md);;所有文件 (*)"
        )
        
        if file_path:
            self.input_file = file_path
            # 显示文件名而不是完整路径
            file_name = os.path.basename(file_path)
            self.input_file_label.setText(f"已选择: {file_name}")
            self.input_file_label.setStyleSheet("color: green; font-weight: bold")
            self.status_label.setText("")
    
    def browse_output_dir(self):
        """浏览并选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择输出目录", self.output_dir
        )
        
        if dir_path:
            self.output_dir = dir_path
            self.output_dir_label.setText(f"输出目录: {self.output_dir}")
    
    def get_selected_formats(self):
        """获取选中的输出格式"""
        formats = []
        if self.md_checkbox.isChecked():
            formats.append('markdown')
        if self.csv_checkbox.isChecked():
            formats.append('csv')
        if self.json_checkbox.isChecked():
            formats.append('json')
        if self.txt_checkbox.isChecked():
            formats.append('txt')
        return formats
    
    def start_processing(self):
        """开始处理文件"""
        # 验证输入
        if not self.input_file:
            QMessageBox.warning(self, "警告", "请先选择输入文件")
            return
        
        formats = self.get_selected_formats()
        if not formats:
            QMessageBox.warning(self, "警告", "请至少选择一种输出格式")
            return
        
        # 禁用按钮，显示进度条
        self.process_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_label.setText("正在处理...")
        
        # 创建处理线程
        self.thread = ProcessingThread(
            self.processor, self.input_file, self.output_dir, formats
        )
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.processing_finished)
        self.thread.start()
    
    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def processing_finished(self, success, message):
        """处理完成后的操作"""
        # 恢复界面状态
        self.process_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # 显示结果
        if success:
            self.status_label.setText(message)
            QMessageBox.information(self, "成功", f"转换完成！\n{message}\n文件已保存在：{self.output_dir}")
        else:
            self.status_label.setText("转换失败")
            self.status_label.setStyleSheet("color: red; font-weight: bold")
            QMessageBox.critical(self, "错误", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WordProcessorGUI()
    window.show()
    sys.exit(app.exec_())