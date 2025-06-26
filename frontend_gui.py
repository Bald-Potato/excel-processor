import sys
from pathlib import Path
import threading
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QCheckBox,
                             QScrollArea)
from PyQt6.QtCore import Qt
import process_backend

class SimpleExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # 设置窗口标题和大小
        self.setWindowTitle('Excel StartTime 处理工具')
        self.setGeometry(100, 100, 800, 550)
        
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 文件夹路径选择部分
        path_label = QLabel('选择主文件夹路径：')
        main_layout.addWidget(path_label)
        
        path_layout = QHBoxLayout()
        self.path_entry = QLineEdit()
        browse_button = QPushButton('浏览')
        browse_button.clicked.connect(self.browse_folder)
        path_layout.addWidget(self.path_entry)
        path_layout.addWidget(browse_button)
        main_layout.addLayout(path_layout)
        
        # 功能选项部分
        options_layout = QVBoxLayout()
        
        # 添加勾选框
        self.check_divisible = QCheckBox('判断数据合法性')
        self.check_divisible.setToolTip('判断提取出的starttime列内每个单元格的数据是否是40的整数倍')
        options_layout.addWidget(self.check_divisible)
        
        self.convert_time = QCheckBox('采用24小时制')
        self.convert_time.setToolTip('将提取出的starttime列内每个单元格的数据由毫秒单位转化为hh:mm:ss的格式')
        options_layout.addWidget(self.convert_time)
        
        # 自定义后缀部分
        suffix_label = QLabel('自定义后缀：')
        options_layout.addWidget(suffix_label)
        
        # 创建一个滚动区域来容纳自定义后缀输入框
        suffix_scroll_area = QScrollArea()
        suffix_scroll_area.setWidgetResizable(True)
        suffix_scroll_area.setMaximumHeight(250)  # 限制最大高度
        
        suffix_container = QWidget()
        self.suffix_layout = QVBoxLayout(suffix_container)
        suffix_scroll_area.setWidget(suffix_container)
        
        # 自定义后缀按钮布局
        suffix_buttons_layout = QHBoxLayout()
        add_suffix_button = QPushButton('添加自定义后缀')
        add_suffix_button.clicked.connect(self.add_custom_suffix)
        delete_suffix_button = QPushButton('删除自定义后缀')
        delete_suffix_button.clicked.connect(self.delete_custom_suffix)
        
        suffix_buttons_layout.addWidget(add_suffix_button)
        suffix_buttons_layout.addWidget(delete_suffix_button)
        
        options_layout.addLayout(suffix_buttons_layout)
        options_layout.addWidget(suffix_scroll_area)
        
        # 存储自定义后缀输入框的列表
        self.custom_suffix_entries = []
        
        main_layout.addLayout(options_layout)
        
        # 操作按钮部分
        button_layout = QHBoxLayout()
        self.start_button = QPushButton('开始处理')
        self.start_button.clicked.connect(self.start_processing)
        self.exit_button = QPushButton('退出程序')
        self.exit_button.clicked.connect(self.close)
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.exit_button)
        main_layout.addLayout(button_layout)
        
        # 日志输出部分
        log_label = QLabel('处理日志：')
        main_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)
        
        # 重定向标准输出到日志窗口
        sys.stdout = self
        
        # 显示初始消息
        self.write('程序已启动，请选择要处理的文件夹并点击"开始处理"按钮。\n')
    
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if folder:
            self.path_entry.setText(folder)
    
    def add_custom_suffix(self):
        # 限制最多添加10个自定义后缀
        if len(self.custom_suffix_entries) >= 10:
            self.write('⚠️  最多只能添加10个自定义后缀。\n')
            return
            
        # 创建新的自定义后缀输入框
        suffix_idx = len(self.custom_suffix_entries) + 1
        suffix_layout = QHBoxLayout()
        suffix_label = QLabel(f'自定义后缀{suffix_idx}:')
        suffix_entry = QLineEdit()
        suffix_layout.addWidget(suffix_label)
        suffix_layout.addWidget(suffix_entry)
        
        self.suffix_layout.addLayout(suffix_layout)
        self.custom_suffix_entries.append(suffix_entry)
    
    def delete_custom_suffix(self):
        # 删除最后一个自定义后缀输入框
        if self.custom_suffix_entries:
            # 获取最后一个输入框的布局
            layout_item = self.suffix_layout.itemAt(self.suffix_layout.count() - 1)
            if layout_item:
                # 获取水平布局
                h_layout = layout_item.layout()
                if h_layout:
                    # 清除水平布局中的所有控件
                    while h_layout.count():
                        item = h_layout.takeAt(0)
                        widget = item.widget()
                        if widget:
                            widget.deleteLater()
                    # 从垂直布局中移除水平布局
                    self.suffix_layout.removeItem(layout_item)
                    # 从列表中移除最后一个输入框
                    self.custom_suffix_entries.pop()
    
    def start_processing(self):
        folder = self.path_entry.text()
        if not folder:
            self.write('⚠️  请选择要处理的文件夹。\n')
            return
        path = Path(folder)
        if not path.exists():
            self.write('❌ 文件夹不存在。\n')
            return
        
        # 获取勾选框状态
        check_divisible = self.check_divisible.isChecked()
        convert_time = self.convert_time.isChecked()
        
        # 获取自定义后缀
        custom_suffixes = [entry.text() for entry in self.custom_suffix_entries if entry.text().strip()]
        
        # 显示选择的功能
        features = []
        if check_divisible:
            features.append("判断数据合法性")
        if convert_time:
            features.append("采用24小时制")
        if custom_suffixes:
            features.append(f"自定义后缀: {', '.join(custom_suffixes)}")
            
        if features:
            self.write(f"已启用功能: {', '.join(features)}\n")
        
        # 禁用所有按钮，除了退出按钮
        self.start_button.setEnabled(False)
        for i in range(self.suffix_layout.count()):
            layout_item = self.suffix_layout.itemAt(i)
            if layout_item and layout_item.layout():
                h_layout = layout_item.layout()
                for j in range(h_layout.count()):
                    widget = h_layout.itemAt(j).widget()
                    if isinstance(widget, QLineEdit):
                        widget.setEnabled(False)
        
        # 在单独的线程中运行处理任务，避免界面冻结
        def process_and_enable_buttons():
            process_backend.process_folder(path, check_divisible, convert_time, custom_suffixes)
            # 处理完成后恢复按钮
            self.start_button.setEnabled(True)
            for i in range(self.suffix_layout.count()):
                layout_item = self.suffix_layout.itemAt(i)
                if layout_item and layout_item.layout():
                    h_layout = layout_item.layout()
                    for j in range(h_layout.count()):
                        widget = h_layout.itemAt(j).widget()
                        if isinstance(widget, QLineEdit):
                            widget.setEnabled(True)
        
        threading.Thread(
            target=process_and_enable_buttons,
            daemon=True
        ).start()
    
    # 实现write和flush方法以重定向标准输出
    def write(self, text):
        if text:
            self.log_text.append(text.rstrip('\n'))
            # 滚动到底部
            self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
    
    def flush(self):
        pass

# 启动应用程序
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SimpleExcelApp()
    window.show()
    sys.exit(app.exec())