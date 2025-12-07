import sys
import os
import urllib.parse
import datetime
import traceback
import json
from PIL import Image
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QWidget, QSpinBox, 
                             QComboBox, QMessageBox, QTextEdit, QGroupBox, QSplitter,
                             QLineEdit, QMenuBar, QMenu, QStatusBar, QFrame, QScrollArea,
                             QProgressBar, QAction, QDialog, QTextBrowser, QDialogButtonBox,
                             QCheckBox, QFormLayout)
from PyQt5.QtCore import Qt, QUrl, QTimer
from PyQt5.QtGui import QPixmap, QImage, QIcon, QDesktopServices, QTextCursor, QColor

DEBUG_MODE = True

def get_resource_path(relative_path):
    """获取资源的绝对路径，支持打包后的exe文件"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def get_documents_path():
    """获取系统文档文件夹路径"""
    try:
        import ctypes
        from ctypes import wintypes, windll
        
        # 使用Windows API获取文档文件夹路径
        CSIDL_PERSONAL = 5  # My Documents
        SHGFP_TYPE_CURRENT = 0  # 获取当前路径，不是默认值
        
        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
        
        return buf.value
    except:
        # 如果失败，使用默认路径
        return os.path.join(os.path.expanduser('~'), 'Documents')

def parse_command_line():
    """解析命令行参数"""
    if len(sys.argv) > 1:
        # 第一个参数是程序自身，后面的参数是拖放的文件
        file_paths = []
        for arg in sys.argv[1:]:
            # 检查是否是有效的图片文件
            if os.path.exists(arg) and arg.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif', '.webp')):
                file_paths.append(arg)
        return file_paths
    return []

class AppConfig:
    def __init__(self):
        # 修改配置存储位置为 文档\E-IMG Slices\Config
        docs_path = get_documents_path()
        self.config_dir = os.path.join(docs_path, 'E-IMG Slices', 'Config')
        self.config_file = os.path.join(self.config_dir, 'config.json')
        
        # 默认配置值
        self.debug_mode = False 
        self.auto_create_folder = True 
        self.folder_name = "Slices" 
        
        # 新增：默认切片设置
        self.default_slice_name = ""  # 留空则使用图片名称
        self.default_slice_format = "JPG"
        self.default_slice_direction = "纵向"
        self.default_slice_method = "按大小切片"
        self.default_slice_size = 1080
        self.default_slice_count = 16
        
        os.makedirs(self.config_dir, exist_ok=True)
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                    self.auto_create_folder = config_data.get('auto_create_folder', True)
                    self.folder_name = config_data.get('folder_name', "Slices")
                    
                    # 加载新的默认设置
                    self.default_slice_name = config_data.get('default_slice_name', "")
                    self.default_slice_format = config_data.get('default_slice_format', "JPG")
                    self.default_slice_direction = config_data.get('default_slice_direction', "纵向")
                    self.default_slice_method = config_data.get('default_slice_method', "按大小切片")
                    self.default_slice_size = config_data.get('default_slice_size', 1080)
                    self.default_slice_count = config_data.get('default_slice_count', 16)
            except Exception as e:
                print(f"加载配置失败: {e}")
                self.auto_create_folder = True
                self.folder_name = "Slices"
    
    def save_config(self):
        """保存配置文件"""
        config_data = {
            'debug_mode': False,  
            'auto_create_folder': self.auto_create_folder,
            'folder_name': self.folder_name,
            
            # 保存新的默认设置
            'default_slice_name': self.default_slice_name,
            'default_slice_format': self.default_slice_format,
            'default_slice_direction': self.default_slice_direction,
            'default_slice_method': self.default_slice_method,
            'default_slice_size': self.default_slice_size,
            'default_slice_count': self.default_slice_count
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"保存配置失败: {e}")

class SettingsWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.config = parent.config if parent else AppConfig()
        self.setWindowTitle("设置")
        self.setModal(True)
        self.setFixedSize(400, 500)
        
        self.init_ui()
        self.load_settings()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # 默认切片设置组
        settings_group = QGroupBox("默认切片设置")
        form_layout = QFormLayout(settings_group)
        
        # 默认切片文件名称
        self.default_name_edit = QLineEdit()
        self.default_name_edit.setPlaceholderText("留空则使用图片名称")
        form_layout.addRow("默认切片文件名称:", self.default_name_edit)
        
        # 默认切片输出格式
        self.default_format_combo = QComboBox()
        self.default_format_combo.addItems(["JPG", "PNG"])
        form_layout.addRow("默认切片输出格式:", self.default_format_combo)
        
        # 默认切片方向
        self.default_direction_combo = QComboBox()
        self.default_direction_combo.addItems(["横向", "纵向"])
        form_layout.addRow("默认切片方向:", self.default_direction_combo)
        
        # 默认切片方式
        self.default_method_combo = QComboBox()
        self.default_method_combo.addItems(["按大小切片", "按数量切片"])
        form_layout.addRow("默认切片方式:", self.default_method_combo)
        
        # 默认切片大小
        self.default_size_spin = QSpinBox()
        self.default_size_spin.setMinimum(1)
        self.default_size_spin.setMaximum(10000)
        form_layout.addRow("默认切片大小:", self.default_size_spin)
        
        # 默认切片份数
        self.default_count_spin = QSpinBox()
        self.default_count_spin.setMinimum(1)
        self.default_count_spin.setMaximum(1000)
        form_layout.addRow("默认切片份数:", self.default_count_spin)
        
        layout.addWidget(settings_group)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        self.save_btn = QPushButton("保存")
        self.cancel_btn = QPushButton("取消")
        self.apply_btn = QPushButton("应用")
        
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.cancel_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.apply_btn)
        
        layout.addLayout(button_layout)
        
        # 连接信号
        self.save_btn.clicked.connect(self.save_and_close)
        self.cancel_btn.clicked.connect(self.reject)
        self.apply_btn.clicked.connect(self.apply_settings)
        
        # 连接切片方式变化事件
        self.default_method_combo.currentIndexChanged.connect(self.update_settings_visibility)
        self.update_settings_visibility()
    
    def update_settings_visibility(self):
        """根据切片方式显示/隐藏相关设置"""
        method = self.default_method_combo.currentText()
        is_size_method = (method == "按大小切片")
        
        # 显示/隐藏相关控件
        self.default_size_spin.setVisible(is_size_method)
        self.default_count_spin.setVisible(not is_size_method)
        
        # 更新表单布局的标签
        form_layout = self.layout().itemAt(0).widget().layout()
        if isinstance(form_layout, QFormLayout):
            # 获取标签项并设置可见性
            for i in range(form_layout.rowCount()):
                item = form_layout.itemAt(i, QFormLayout.LabelRole)
                if item and item.widget():
                    label_text = item.widget().text()
                    if "默认切片大小" in label_text:
                        item.widget().setVisible(is_size_method)
                    elif "默认切片份数" in label_text:
                        item.widget().setVisible(not is_size_method)
    
    def load_settings(self):
        """加载当前设置"""
        self.default_name_edit.setText(self.config.default_slice_name)
        self.default_format_combo.setCurrentText(self.config.default_slice_format)
        self.default_direction_combo.setCurrentText(self.config.default_slice_direction)
        self.default_method_combo.setCurrentText(self.config.default_slice_method)
        self.default_size_spin.setValue(self.config.default_slice_size)
        self.default_count_spin.setValue(self.config.default_slice_count)
    
    def apply_settings(self):
        """应用设置但不关闭窗口"""
        try:
            self.config.default_slice_name = self.default_name_edit.text().strip()
            self.config.default_slice_format = self.default_format_combo.currentText()
            self.config.default_slice_direction = self.default_direction_combo.currentText()
            self.config.default_slice_method = self.default_method_combo.currentText()
            self.config.default_slice_size = self.default_size_spin.value()
            self.config.default_slice_count = self.default_count_spin.value()
            
            self.config.save_config()
            
            # 更新主窗口的设置
            if self.parent:
                self.parent.apply_default_settings()
            
            QMessageBox.information(self, "成功", "设置已应用！")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"应用设置失败: {str(e)}")
    
    def save_and_close(self):
        """保存设置并关闭窗口"""
        self.apply_settings()
        self.accept()

class AboutWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于E-IMG Slices")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setModal(True)
        
        about_image_path = get_resource_path("about.jpg")
        
        layout = QVBoxLayout(self)
        
        if os.path.exists(about_image_path):
            pixmap = QPixmap(about_image_path)
            image_label = QLabel()
            image_label.setPixmap(pixmap)
            image_label.setAlignment(Qt.AlignCenter)
            image_label.setCursor(Qt.PointingHandCursor)  
            image_label.mousePressEvent = self.close  
            
            layout.addWidget(image_label)
                 
            self.resize(pixmap.width(), pixmap.height())
        else:
            error_label = QLabel("文件丢失：about哪去了？")
            error_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(error_label)

            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(self.close)
            layout.addWidget(close_btn)
    
    def close(self, event=None):
        """重写close方法以支持点击图片关闭"""
        super().accept()

class DebugLogWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Debug日志输出")
        self.setGeometry(100, 100, 800, 500)
        self.parent = parent
        self.is_task_interrupted = False

        layout = QVBoxLayout(self)
        
        self.log_text = QTextBrowser()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        button_layout = QHBoxLayout()
        self.open_log_dir_btn = QPushButton("打开日志文件夹")
        self.clear_log_btn = QPushButton("清空日志")
        self.interrupt_btn = QPushButton("中断当前任务")
        self.interrupt_btn.setEnabled(False)
        self.close_btn = QPushButton("关闭")
        
        button_layout.addWidget(self.open_log_dir_btn)
        button_layout.addWidget(self.clear_log_btn)
        button_layout.addWidget(self.interrupt_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)
        
        layout.addLayout(button_layout)

        self.open_log_dir_btn.clicked.connect(self.open_log_directory)
        self.clear_log_btn.clicked.connect(self.clear_log)
        self.interrupt_btn.clicked.connect(self.interrupt_task)
        self.close_btn.clicked.connect(self.accept)
        
        # 修改日志存储位置为 文档\E-IMG Slices\Log
        docs_path = get_documents_path()
        self.log_dir = os.path.join(docs_path, 'E-IMG Slices', 'Log')
        os.makedirs(self.log_dir, exist_ok=True)
        self.log_file = os.path.join(self.log_dir, f"debug_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        
    def append_log(self, message, log_type="INFO", color="black"):
        """添加日志信息"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp} {log_type}] {message}"
        
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        self.log_text.setFontWeight(75)
        self.log_text.setTextColor(QColor(color))
        self.log_text.insertPlainText(f"[{timestamp} {log_type}] ")
        
        self.log_text.setFontWeight(50)
        self.log_text.insertPlainText(message + "\n")
        
        self.log_text.setTextColor(QColor("black"))
        self.log_text.ensureCursorVisible()
        
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(formatted_message + '\n')
        except:
            pass
        
    def open_log_directory(self):
        """打开日志文件夹"""
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.log_dir))
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def interrupt_task(self):
        """中断当前任务"""
        self.is_task_interrupted = True
        self.interrupt_btn.setEnabled(False)
        self.append_log("用户请求中断当前任务", "WARNING", "orange")
        if self.parent:
            self.parent.set_progress_status("任务已中断", "orange")
    
    def reset_interrupt(self):
        """重置中断状态"""
        self.is_task_interrupted = False
        self.interrupt_btn.setEnabled(False)

def debug_print(*args, **kwargs):
    """调试输出函数"""
    if DEBUG_MODE:
        print("[DEBUG]", *args, **kwargs)
        if sys.stdout is not None:
            sys.stdout.flush()

class ImageSlicer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.image = None
        self.image_path = None
        self.last_save_dir = None
        self.config = AppConfig()
        self.debug_window = None
        self.is_slicing = False
        self.auto_export_on_startup = False  # 添加这个属性
        
        debug_print("程序启动，初始化界面...")
        self.initUI()
        self.apply_default_settings()  # 应用默认设置
        
    def apply_default_settings(self):
        """应用默认设置到界面"""
        # 设置默认值
        self.name_edit.setText(self.config.default_slice_name)
        self.format_combo.setCurrentText(self.config.default_slice_format)
        self.direction_combo.setCurrentText(self.config.default_slice_direction)
        self.method_combo.setCurrentText(self.config.default_slice_method)
        
        if self.config.default_slice_method == "按大小切片":
            self.param_spin.setValue(self.config.default_slice_size)
        else:
            self.param_spin.setValue(self.config.default_slice_count)
        
        self.update_param_hint()  # 更新参数标签
        
    def handle_dropped_file_at_startup(self, file_path):
        """处理程序启动时拖放的文件"""
        try:
            self.debug_log(f"处理启动时拖放的文件: {file_path}", "AUTO_EXPORT", "blue")
            
            # 加载图片
            self.load_image_from_path(file_path)
            
            if not self.image:
                self.debug_log("无法加载图片，取消自动导出", "AUTO_EXPORT", "red")
                return
            
            # 设置为自动导出模式
            self.auto_export_on_startup = True
            
            # 应用默认设置（使用配置中的按大小切片设置）
            self.apply_default_settings()
            
            # 确保使用按大小切片方式
            self.method_combo.setCurrentText(self.config.default_slice_method)
            
            self.debug_log("设置已应用，准备自动导出", "AUTO_EXPORT", "blue")
            
            # 延迟执行自动导出，确保UI完全更新
            QTimer.singleShot(1000, self.auto_quick_export)
            
        except Exception as e:
            self.debug_log(f"处理拖放文件失败: {str(e)}", "ERROR", "red")
            QMessageBox.critical(self, "错误", f"处理拖放文件失败: {str(e)}")
    
    def auto_quick_export(self):
        """自动执行快速导出"""
        try:
            if not self.image or not self.auto_export_on_startup:
                return
            
            self.debug_log("开始自动快速导出", "AUTO_EXPORT", "blue")
            
            # 获取当前设置
            direction = self.direction_combo.currentText()
            method = self.method_combo.currentText()
            param = self.param_spin.value()
            base_name = self.name_edit.text().strip() or os.path.splitext(os.path.basename(self.image_path))[0]
            file_format = self.format_combo.currentText().lower()
            
            self.debug_log(f"自动导出设置 - 方向: {direction}, 方法: {method}, 参数: {param}, 名称: {base_name}, 格式: {file_format}", "AUTO_EXPORT", "blue")
            
            # 使用图片所在目录作为保存目录
            image_dir = os.path.dirname(self.image_path)
            save_dir = image_dir
            
            self.debug_log(f"自动导出目录: {save_dir}", "AUTO_EXPORT", "blue")
            
            # 自动创建文件夹
            if self.config.auto_create_folder:
                folder_name = self.config.folder_name.strip() or "Slices"
                save_dir = os.path.join(save_dir, folder_name)
                os.makedirs(save_dir, exist_ok=True)
                self.debug_log(f"已创建输出文件夹: {save_dir}", "AUTO_EXPORT", "green")
            
            # 检查文件冲突
            conflict_files = self.check_all_file_conflicts(save_dir, base_name, file_format, direction, method, param)
            
            if conflict_files:
                self.debug_log(f"发现 {len(conflict_files)} 个文件冲突", "WARNING", "orange")
                # 自动模式下自动覆盖
                self.debug_log("自动模式：自动覆盖现有文件", "AUTO_EXPORT", "orange")
            
            self.set_progress_status("正在自动导出...", "blue")
            
            # 执行切片
            if method == "按大小切片":
                result = self.slice_by_size(direction, param, save_dir, base_name, file_format, conflict_files)
            else:
                result = self.slice_by_count(direction, param, save_dir, base_name, file_format, conflict_files)
            
            if result:
                self.debug_log("自动导出完成", "AUTO_EXPORT", "green")
                self.set_progress_status("自动导出完成", "green")
                
                # 显示完成提示，使用三个按钮
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("完成")
                msg_box.setText(f"图片已自动切片并保存到:\n{save_dir}")
                msg_box.setIcon(QMessageBox.Information)

                # 添加三个按钮
                ok_button = msg_box.addButton("OK", QMessageBox.AcceptRole)
                exit_button = msg_box.addButton("退出", QMessageBox.RejectRole)
                view_button = msg_box.addButton("查看", QMessageBox.ActionRole)
                
                # 设置默认按钮
                msg_box.setDefaultButton(ok_button)
                
                msg_box.exec_()
                
                clicked_button = msg_box.clickedButton()
                
                if clicked_button == view_button:
                    self.debug_log("用户点击查看按钮，打开输出目录", "AUTO_EXPORT", "blue")
                    QDesktopServices.openUrl(QUrl.fromLocalFile(save_dir))
                elif clicked_button == exit_button:
                    self.debug_log("用户点击退出按钮，关闭程序", "AUTO_EXPORT", "blue")
                    self.close()  # 关闭程序
            
            # 重置自动导出标志
            self.auto_export_on_startup = False
            
        except Exception as e:
            self.debug_log(f"自动导出失败: {str(e)}", "ERROR", "red")
            self.set_progress_status("自动导出失败", "red")
            QMessageBox.critical(self, "错误", f"自动导出失败: {str(e)}")
            self.auto_export_on_startup = False
        
    def initUI(self):
        self.setWindowTitle('E-IMG 图片切片工具')
        self.setGeometry(100, 100, 900, 600)  
        
        logo_path = get_resource_path("logo.ico")
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))
            if hasattr(QApplication, 'setWindowIcon'):
                QApplication.setWindowIcon(QIcon(logo_path))

        self.createMenuBar()
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(10)
        
        self.drop_label = QLabel()
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setText("拖放图片到这里或点击\"加载图片\"")
        self.drop_label.setMinimumSize(400, 300)  
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 2px dashed gray; 
                background-color: #f0f0f0;
                border-radius: 5px;
                padding: 20px;
                font-size: 14px;
                color: #666;
            }
            QLabel:hover {
                background-color: #e8e8e8;
                border-color: #0071bc;
            }
        """)
        left_layout.addWidget(self.drop_label)
        
        info_group = QGroupBox("图片信息")
        info_layout = QVBoxLayout(info_group)
        
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        info_layout.addWidget(self.info_text)
        
        left_layout.addWidget(info_group)
        
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setSpacing(10)
        
        settings_group = QGroupBox("切片设置")
        settings_layout = QVBoxLayout(settings_group)

        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("文件命名:"))
        self.name_edit = QLineEdit("")
        self.name_edit.setPlaceholderText("自动使用图片名称")
        name_layout.addWidget(self.name_edit)
        
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("输出格式:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["JPG", "PNG"])
        format_layout.addWidget(self.format_combo)
        
        direction_layout = QHBoxLayout()
        direction_layout.addWidget(QLabel("切片方向:"))
        self.direction_combo = QComboBox()
        self.direction_combo.addItems(["横向", "纵向"])
        direction_layout.addWidget(self.direction_combo)
        
        method_layout = QHBoxLayout()
        method_layout.addWidget(QLabel("切片方式:"))
        self.method_combo = QComboBox()
        self.method_combo.addItems(["按大小切片", "按数量切片"])
        method_layout.addWidget(self.method_combo)
        
        param_layout = QHBoxLayout()
        self.param_label = QLabel("参数值:")  # 初始文本，后续会更新
        param_layout.addWidget(self.param_label)
        self.param_spin = QSpinBox()
        self.param_spin.setMinimum(1)
        self.param_spin.setMaximum(10000)
        self.param_spin.setValue(500)  
        param_layout.addWidget(self.param_spin)
        
        folder_layout = QHBoxLayout()
        self.auto_folder_check = QCheckBox("输出时自动创建文件夹")
        self.auto_folder_check.setChecked(self.config.auto_create_folder)
        self.auto_folder_check.stateChanged.connect(self.toggle_auto_folder)
        folder_layout.addWidget(self.auto_folder_check)
        
        folder_name_layout = QHBoxLayout()
        folder_name_layout.addWidget(QLabel("文件夹名称:"))
        self.folder_name_edit = QLineEdit(self.config.folder_name)
        self.folder_name_edit.textChanged.connect(self.update_folder_name)
        folder_name_layout.addWidget(self.folder_name_edit)
        
        button_layout = QHBoxLayout()
        self.load_btn = QPushButton("加载图片")
        self.slice_btn = QPushButton("开始切片")
        self.slice_btn.setEnabled(False)
        self.quick_export_btn = QPushButton("快速导出")
        self.quick_export_btn.setEnabled(False)
        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.slice_btn)
        button_layout.addWidget(self.quick_export_btn)
        
        settings_layout.addLayout(name_layout)
        settings_layout.addLayout(format_layout)
        settings_layout.addLayout(direction_layout)
        settings_layout.addLayout(method_layout)
        settings_layout.addLayout(param_layout)
        settings_layout.addLayout(folder_layout)
        settings_layout.addLayout(folder_name_layout)
        settings_layout.addLayout(button_layout)
        
        right_layout.addWidget(settings_group)
        
        preview_group = QGroupBox("切片预览信息")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        
        right_layout.addWidget(preview_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.set_progress_status("就绪", "gray")
        right_layout.addWidget(self.progress_bar)
        
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 500])
        
        main_layout.addWidget(splitter)
        
        self.createStatusBar()
        
        self.load_btn.clicked.connect(self.load_image)
        self.slice_btn.clicked.connect(self.slice_image)
        self.quick_export_btn.clicked.connect(self.quick_export)
        self.method_combo.currentIndexChanged.connect(self.update_param_hint)
        self.direction_combo.currentIndexChanged.connect(self.update_preview_if_enabled)
        self.param_spin.valueChanged.connect(self.update_preview_if_enabled)
        
        self.update_param_hint()
        
        self.setAcceptDrops(True)
        self.drop_label.dragEnterEvent = self.dragEnterEvent
        self.drop_label.dropEvent = self.dropEvent
        
        # 为快速导出按钮单独设置样式
        self.quick_export_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFA726;
                border: none;
                color: white;
                padding: 8px 16px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #FFCC80;
            }
            QPushButton:pressed {
                background-color: #FFD54F;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        debug_print("界面初始化完成")
    
    def closeEvent(self, event):
        """重写关闭事件，确保debug窗口也关闭"""
        if self.debug_window:
            self.debug_window.close()
        self.auto_export_on_startup = False  # 添加这行
        event.accept()
        
    def set_progress_status(self, text, color="gray"):
        """设置进度条状态"""
        self.progress_bar.setFormat(text)
        if color == "gray":
            self.progress_bar.setStyleSheet("QProgressBar { background-color: #f0f0f0; } QProgressBar::chunk { background-color: #d0d0d0; }")
        elif color == "green":
            self.progress_bar.setStyleSheet("QProgressBar { background-color: #f0f0f0; } QProgressBar::chunk { background-color: #4CAF50; }")
        elif color == "orange":
            self.progress_bar.setStyleSheet("QProgressBar { background-color: #f0f0f0; } QProgressBar::chunk { background-color: #FF9800; }")
        elif color == "red":
            self.progress_bar.setStyleSheet("QProgressBar { background-color: #f0f0f0; } QProgressBar::chunk { background-color: #F44336; }")
        elif color == "blue":
            self.progress_bar.setStyleSheet("QProgressBar { background-color: #f0f0f0; } QProgressBar::chunk { background-color: #2196F3; }")
    
    def update_progress(self, value, text=None):
        """更新进度条"""
        self.progress_bar.setValue(value)
        if text:
            self.progress_bar.setFormat(text)
    
    def createMenuBar(self):
        menubar = self.menuBar()

        function_menu = menubar.addMenu('功能')
        
        self.debug_action = QAction('Debug', self)
        self.debug_action.setCheckable(True)
        self.debug_action.setChecked(False)  
        self.debug_action.triggered.connect(self.toggle_debug)
        function_menu.addAction(self.debug_action)
        
        # 添加设置菜单
        settings_action = QAction('设置', self)
        settings_action.triggered.connect(self.open_settings)
        function_menu.addAction(settings_action)

        about_menu = menubar.addMenu('帮助')
        
        about_action = about_menu.addAction('关于E-IMG Slices')
        about_action.triggered.connect(self.openAboutWindow)
        
        github_action = about_menu.addAction('GitHub项目页')
        github_action.triggered.connect(self.openGithubUrl)
    
    def open_settings(self):
        """打开设置窗口"""
        settings_window = SettingsWindow(self)
        settings_window.exec_()
        
    def toggle_debug(self, checked):
        """切换Debug模式"""
        if checked:
            self.debug_window = DebugLogWindow(self)
            self.debug_window.show()
            self.debug_window.append_log("Debug模式已启用", "INFO", "black")
            self.debug_log("Debug窗口已打开")
        else:
            if self.debug_window:
                self.debug_window.append_log("Debug模式已禁用", "INFO", "black")
                self.debug_window.close()
                self.debug_window = None
    
    def debug_log(self, message, log_type="INFO", color="black"):
        """记录Debug日志"""
        if self.debug_window:
            self.debug_window.append_log(message, log_type, color)
    
    def toggle_auto_folder(self, state):
        """切换自动创建文件夹功能"""
        self.config.auto_create_folder = (state == Qt.Checked)
        self.config.save_config()
        self.debug_log(f"自动创建文件夹设置已更新: {self.config.auto_create_folder}", "SETTING", "blue")
    
    def update_folder_name(self, text):
        """更新文件夹名称"""
        self.config.folder_name = text.strip() or "Slices"
        self.config.save_config()
        self.debug_log(f"文件夹名称已更新: {self.config.folder_name}", "SETTING", "blue")
        
    def openAboutWindow(self):
        """打开关于E-IMG Slices窗口"""
        about_window = AboutWindow(self)
        about_window.exec_()
        
    def createStatusBar(self):
        self.statusbar = QStatusBar()
        self.statusbar.showMessage("E-IMG Slices | V1.4-Beta")
        self.setStatusBar(self.statusbar)
        
    def openGithubUrl(self):
        QDesktopServices.openUrl(QUrl('https://github.com/404Solved/E-IMG-Slices'))
        self.restoreStatusBar()
        
    def restoreStatusBar(self):
        QTimer.singleShot(100, lambda: self.statusbar.showMessage("E-IMG Slices | V1.4-Beta"))
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.debug_log("拖放事件开始: 检测到文件", "DRAG", "blue")
            
    def dropEvent(self, event):
        try:
            self.debug_log("处理拖放事件", "DRAG", "blue")
            if event.mimeData().hasUrls():
                url = event.mimeData().urls()[0]
                file_path = url.toLocalFile()
                self.debug_log(f"拖放文件路径: {file_path}", "DRAG", "blue")

                if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif', '.webp')):
                    self.debug_log("文件类型验证通过，开始加载图片", "DRAG", "green")
                    self.load_image_from_path(file_path)
                else:
                    self.debug_log("文件类型验证失败: 不是支持的图片格式", "DRAG", "red")
                    QMessageBox.warning(self, "错误", "请拖放有效的图片文件")
        except Exception as e:
            self.debug_log(f"拖放事件异常: {str(e)}", "ERROR", "red")
            QMessageBox.critical(self, "错误", f"拖放操作失败: {str(e)}")
                
    def load_image_from_path(self, file_path):
        try:
            self.debug_log(f"开始加载图片: {file_path}", "LOAD", "blue")
            self.set_progress_status("正在导入...", "blue")
            QApplication.processEvents()  

            try:
                self.debug_log("尝试打开图片文件...", "LOAD", "blue")
                self.image = Image.open(file_path)
                self.debug_log("图片打开成功，开始验证...", "LOAD", "blue")

                self.image.verify()
                self.debug_log("图片验证成功", "LOAD", "green")

                self.image = Image.open(file_path)
                self.debug_log("图片重新打开成功", "LOAD", "green")
                
            except Exception as e:
                self.debug_log(f"图片验证失败: {str(e)}", "ERROR", "red")
                raise Exception(f"图片文件损坏或格式不受支持: {str(e)}")
            
            self.image_path = file_path
            self.debug_log("图片基本信息设置完成", "LOAD", "green")

            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.name_edit.setText(base_name)
            self.debug_log(f"自动设置文件命名前缀: {base_name}", "SETTING", "blue")

            self.drop_label.setText(f"已加载图片: {os.path.basename(file_path)}\n\n拖放新图片替换当前图片")
            self.drop_label.setStyleSheet("""
                QLabel {
                    border: 2px dashed #0071bc; 
                    background-color: #eceff1;
                    border-radius: 5px;
                    padding: 20px;
                    font-size: 14px;
                    color: #0059A8;
                }
                QLabel:hover {
                    background-color: #C7D9E2;
                    border-color: #0071bc;
                }
            """)
            
            self.show_image_info()
            self.slice_btn.setEnabled(True)
            self.quick_export_btn.setEnabled(True)
            self.debug_log(f"图片加载完成: {os.path.basename(file_path)}", "LOAD", "green")

            try:
                self.debug_log("开始计算切片预览信息", "PREVIEW", "blue")
                self.preview_slice_info()
                self.debug_log("切片预览计算成功", "PREVIEW", "green")
            except Exception as e:
                self.debug_log(f"切片预览失败: {str(e)}", "WARNING", "orange")
                self.preview_text.clear()
                self.append_preview("预览失败: 无法计算切片信息", "red")
                self.append_preview("您仍然可以尝试继续切片操作", "orange")
            
            self.set_progress_status("就绪", "gray")
            self.debug_log("图片加载流程完成", "LOAD", "green")
            
        except Exception as e:
            self.debug_log(f"图片加载过程中出现严重错误: {str(e)}", "ERROR", "red")
            error_msg = f"无法加载图片: {str(e)}"
            QMessageBox.critical(self, "错误", error_msg)
            self.set_progress_status("导入失败", "red")

            self.image = None
            self.image_path = None
            self.slice_btn.setEnabled(False)
            self.quick_export_btn.setEnabled(False)
            self.drop_label.setText("拖放图片到这里或点击\"加载图片\"")
            self.drop_label.setStyleSheet("""
                QLabel {
                    border: 2px dashed gray; 
                    background-color: #f0f0f0;
                    border-radius: 5px;
                    padding: 20px;
                    font-size: 14px;
                    color: #666;
                }
                QLabel:hover {
                    background-color: #e8e8e8;
                    border-color: #4CAF50;
                }
            """)
            self.info_text.clear()
            self.preview_text.clear()
    
    def update_param_hint(self):
        """更新参数标签和提示"""
        if self.method_combo.currentText() == "按大小切片":
            self.param_spin.setSuffix(" 像素")
            self.param_label.setText("切片大小:")  # 修改标签文本
            self.debug_log("切片方式切换为: 按大小切片", "SETTING", "blue")
        else:
            self.param_spin.setSuffix(" 份")
            self.param_label.setText("切片份数:")  # 修改标签文本
            self.debug_log("切片方式切换为: 按数量切片", "SETTING", "blue")
        self.update_preview_if_enabled()
    
    def update_preview_if_enabled(self):
        if self.image:  
            try:
                self.debug_log("更新切片预览信息", "PREVIEW", "blue")
                self.preview_slice_info()
            except Exception as e:
                self.debug_log(f"实时预览失败: {str(e)}", "WARNING", "orange")
                self.preview_text.clear()
                self.append_preview("预览失败: 无法计算切片信息", "red")
                self.append_preview("您仍然可以尝试继续切片操作", "orange")
    
    def load_image(self):
        try:
            self.debug_log("打开文件对话框选择图片", "LOAD", "blue")
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择图片", "", 
                "图片文件 (*.png *.jpg *.jpeg *.bmp *.tiff *.gif *.webp)"
            )
            
            if file_path:
                self.debug_log(f"用户选择文件: {file_path}", "LOAD", "blue")
                self.load_image_from_path(file_path)
            else:
                self.debug_log("用户取消文件选择", "LOAD", "orange")
        except Exception as e:
            self.debug_log(f"文件对话框异常: {str(e)}", "ERROR", "red")
    
    def show_image_info(self):
        if self.image:
            try:
                self.debug_log("开始获取图片信息", "INFO", "blue")
                width, height = self.image.size
                mode = self.image.mode
                info = f"文件名: {os.path.basename(self.image_path)}\n"
                info += f"尺寸: {width} × {height} 像素\n"
                info += f"颜色模式: {mode}\n"
                
                dpi = self.image.info.get('dpi', (72, 72))
                info += f"分辨率: {dpi[0]} PPI\n"
                
                file_size = os.path.getsize(self.image_path)
                info += f"文件大小: {file_size / 1024:.2f} KB"
                
                self.info_text.setPlainText(info)
                self.debug_log("图片信息显示完成", "INFO", "green")
                
            except Exception as e:
                self.debug_log(f"获取图片信息失败: {str(e)}", "WARNING", "orange")
                self.info_text.setPlainText(f"文件名: {os.path.basename(self.image_path)}\n无法获取完整图片信息")
    
    def append_preview(self, message, color="black"):
        cursor = self.preview_text.textCursor()
        cursor.movePosition(QTextCursor.End)

        self.preview_text.setTextColor(QColor(color))
        self.preview_text.insertPlainText(message + "\n")

        self.preview_text.setTextColor(QColor("black"))
        self.preview_text.ensureCursorVisible()
    
    def preview_slice_info(self):
        if not self.image:
            return
            
        try:
            self.debug_log("开始计算切片预览", "PREVIEW", "blue")
            direction = self.direction_combo.currentText()
            method = self.method_combo.currentText()
            param = self.param_spin.value()
            width, height = self.image.size
            
            self.debug_log(f"切片参数 - 方向: {direction}, 方法: {method}, 参数: {param}, 尺寸: {width}x{height}", "PREVIEW", "blue")
            
            self.preview_text.clear()
            
            if method == "按大小切片":
                if direction == "横向":
                    total_slices = (width + param - 1) // param  
                    remainder = width % param
                    
                    self.append_preview(f"将切成 {total_slices} 份", "black")
                    if remainder > 0:
                        self.append_preview(f"前 {total_slices-1} 份尺寸: {param}×{height} 像素", "black")
                        self.append_preview(f"最后 1 份尺寸: {remainder}×{height} 像素", "black")
                        self.append_preview("末尾切片不满足要求，将直接输出", "orange")
                    else:
                        self.append_preview(f"每份尺寸: {param}×{height} 像素", "black")
                else:  
                    total_slices = (height + param - 1) // param  
                    remainder = height % param
                    
                    self.append_preview(f"将切成 {total_slices} 份", "black")
                    if remainder > 0:
                        self.append_preview(f"前 {total_slices-1} 份尺寸: {width}×{param} 像素", "black")
                        self.append_preview(f"最后 1 份尺寸: {width}×{remainder} 像素", "black")
                        self.append_preview("末尾切片不满足要求，将直接输出", "orange")
                    else:
                        self.append_preview(f"每份尺寸: {width}×{param} 像素", "black")
            else:  
                if direction == "横向":
                    base_width = width // param
                    remainder = width % param
                    
                    self.append_preview(f"将切成 {param} 份", "black")
                    if remainder > 0:
                        self.append_preview(f"前 {remainder} 份尺寸: {base_width+1}×{height} 像素", "black")
                        if param - remainder > 0:
                            self.append_preview(f"后 {param-remainder} 份尺寸: {base_width}×{height} 像素", "black")
                        self.append_preview("已采用余数分散分配处理", "orange")
                    else:
                        self.append_preview(f"每份尺寸: {base_width}×{height} 像素", "black")
                else:  
                    base_height = height // param
                    remainder = height % param
                    
                    self.append_preview(f"将切成 {param} 份", "black")
                    if remainder > 0:
                        self.append_preview(f"前 {remainder} 份尺寸: {width}×{base_height+1} 像素", "black")
                        if param - remainder > 0:
                            self.append_preview(f"后 {param-remainder} 份尺寸: {width}×{base_height} 像素", "black")
                        self.append_preview("已采用余数分散分配处理", "orange")
                    else:
                        self.append_preview(f"每份尺寸: {width}×{base_height} 像素", "black")
            self.debug_log("切片预览计算完成", "PREVIEW", "green")
        except Exception as e:
            self.debug_log(f"切片预览计算失败: {str(e)}", "ERROR", "red")
            raise Exception(f"预览计算失败: {str(e)}")
    
    def slice_image(self):
        if not self.image:
            return
            
        try:
            self.debug_log("开始切片操作", "SLICE", "blue")
            self.is_slicing = True
            if self.debug_window:
                self.debug_window.interrupt_btn.setEnabled(True)
                self.debug_window.is_task_interrupted = False
            
            direction = self.direction_combo.currentText()
            method = self.method_combo.currentText()
            param = self.param_spin.value()
            base_name = self.name_edit.text().strip() or os.path.splitext(os.path.basename(self.image_path))[0]
            file_format = self.format_combo.currentText().lower()
            
            self.debug_log(f"切片设置 - 方向: {direction}, 方法: {method}, 参数: {param}, 名称: {base_name}, 格式: {file_format}", "SLICE", "blue")

            save_dir = QFileDialog.getExistingDirectory(self, "选择保存目录", self.last_save_dir or "")
            if not save_dir:
                self.debug_log("用户取消选择目录", "SLICE", "orange")
                self.is_slicing = False
                if self.debug_window:
                    self.debug_window.interrupt_btn.setEnabled(False)
                return
            
            self.last_save_dir = save_dir
            self.debug_log(f"保存目录: {save_dir}", "SLICE", "blue")
            
            if self.config.auto_create_folder:
                folder_name = self.config.folder_name.strip() or "Slices"
                save_dir = os.path.join(save_dir, folder_name)
                os.makedirs(save_dir, exist_ok=True)
                self.debug_log(f"已创建输出文件夹: {save_dir}", "SLICE", "green")

            conflict_files = self.check_all_file_conflicts(save_dir, base_name, file_format, direction, method, param)
            
            if conflict_files:
                self.debug_log(f"发现 {len(conflict_files)} 个文件冲突: {conflict_files}", "WARNING", "orange")
                reply = QMessageBox.question(self, "文件冲突", 
                                            f"发现 {len(conflict_files)} 个文件已存在，是否全部覆盖？",
                                            QMessageBox.Yes | QMessageBox.No,
                                            QMessageBox.No)
                if reply != QMessageBox.Yes:
                    self.debug_log("用户取消覆盖操作", "SLICE", "orange")
                    self.set_progress_status("操作取消", "orange")
                    self.is_slicing = False
                    if self.debug_window:
                        self.debug_window.interrupt_btn.setEnabled(False)
                    return
                else:
                    self.debug_log("用户确认覆盖现有文件", "SLICE", "orange")
            else:
                self.debug_log("无文件冲突", "SLICE", "green")
                
            self.set_progress_status("正在切片...", "blue")
            QApplication.processEvents()  
            
            if method == "按大小切片":
                self.debug_log("使用按大小切片方法", "SLICE", "blue")
                result = self.slice_by_size(direction, param, save_dir, base_name, file_format, conflict_files)
            else:
                self.debug_log("使用按数量切片方法", "SLICE", "blue")
                result = self.slice_by_count(direction, param, save_dir, base_name, file_format, conflict_files)
            
            if result:
                self.debug_log("切片操作完成", "SLICE", "green")
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("完成")
                msg_box.setText("图片切片已完成！")
                msg_box.setIcon(QMessageBox.Information)

                view_button = msg_box.addButton("查看", QMessageBox.ActionRole)
                msg_box.addButton(QMessageBox.Ok)
                
                msg_box.exec_()
                
                if msg_box.clickedButton() == view_button:
                    self.debug_log("用户点击查看按钮，打开输出目录", "SLICE", "blue")
                    QDesktopServices.openUrl(QUrl.fromLocalFile(save_dir))
            
            self.is_slicing = False
            if self.debug_window:
                self.debug_window.interrupt_btn.setEnabled(False)
            
        except Exception as e:
            self.debug_log(f"切片过程中出现严重错误: {str(e)}", "ERROR", "red")
            error_msg = f"切片过程中出错: {str(e)}"
            QMessageBox.critical(self, "错误", error_msg)
            self.set_progress_status("切片失败", "red")
            self.is_slicing = False
            if self.debug_window:
                self.debug_window.interrupt_btn.setEnabled(False)
    
    def quick_export(self):
        """快速导出功能 - 将切片导出到加载图片所在路径"""
        if not self.image or not self.image_path:
            QMessageBox.warning(self, "警告", "请先加载图片")
            return
            
        try:
            self.debug_log("开始快速导出操作", "QUICK_EXPORT", "blue")
            self.is_slicing = True
            if self.debug_window:
                self.debug_window.interrupt_btn.setEnabled(True)
                self.debug_window.is_task_interrupted = False
            
            direction = self.direction_combo.currentText()
            method = self.method_combo.currentText()
            param = self.param_spin.value()
            base_name = self.name_edit.text().strip() or os.path.splitext(os.path.basename(self.image_path))[0]
            file_format = self.format_combo.currentText().lower()
            
            self.debug_log(f"快速导出设置 - 方向: {direction}, 方法: {method}, 参数: {param}, 名称: {base_name}, 格式: {file_format}", "QUICK_EXPORT", "blue")

            # 使用图片所在目录作为保存目录
            image_dir = os.path.dirname(self.image_path)
            save_dir = image_dir
            
            self.debug_log(f"快速导出目录: {save_dir}", "QUICK_EXPORT", "blue")
            
            # 自动创建文件夹
            if self.config.auto_create_folder:
                folder_name = self.config.folder_name.strip() or "Slices"
                save_dir = os.path.join(save_dir, folder_name)
                os.makedirs(save_dir, exist_ok=True)
                self.debug_log(f"已创建输出文件夹: {save_dir}", "QUICK_EXPORT", "green")

            conflict_files = self.check_all_file_conflicts(save_dir, base_name, file_format, direction, method, param)
            
            if conflict_files:
                self.debug_log(f"发现 {len(conflict_files)} 个文件冲突: {conflict_files}", "WARNING", "orange")
                reply = QMessageBox.question(self, "文件冲突", 
                                            f"发现 {len(conflict_files)} 个文件已存在，是否全部覆盖？",
                                            QMessageBox.Yes | QMessageBox.No,
                                            QMessageBox.No)
                if reply != QMessageBox.Yes:
                    self.debug_log("用户取消覆盖操作", "QUICK_EXPORT", "orange")
                    self.set_progress_status("操作取消", "orange")
                    self.is_slicing = False
                    if self.debug_window:
                        self.debug_window.interrupt_btn.setEnabled(False)
                    return
                else:
                    self.debug_log("用户确认覆盖现有文件", "QUICK_EXPORT", "orange")
            else:
                self.debug_log("无文件冲突", "QUICK_EXPORT", "green")
                
            self.set_progress_status("正在快速导出...", "blue")
            QApplication.processEvents()  
            
            if method == "按大小切片":
                self.debug_log("使用按大小切片方法", "QUICK_EXPORT", "blue")
                result = self.slice_by_size(direction, param, save_dir, base_name, file_format, conflict_files)
            else:
                self.debug_log("使用按数量切片方法", "QUICK_EXPORT", "blue")
                result = self.slice_by_count(direction, param, save_dir, base_name, file_format, conflict_files)
            
            if result:
                self.debug_log("快速导出完成", "QUICK_EXPORT", "green")
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("完成")
                msg_box.setText("图片快速导出已完成！")
                msg_box.setIcon(QMessageBox.Information)

                view_button = msg_box.addButton("查看", QMessageBox.ActionRole)
                msg_box.addButton(QMessageBox.Ok)
                
                msg_box.exec_()
                
                if msg_box.clickedButton() == view_button:
                    self.debug_log("用户点击查看按钮，打开输出目录", "QUICK_EXPORT", "blue")
                    QDesktopServices.openUrl(QUrl.fromLocalFile(save_dir))
            
            self.is_slicing = False
            if self.debug_window:
                self.debug_window.interrupt_btn.setEnabled(False)
            
        except Exception as e:
            self.debug_log(f"快速导出过程中出现严重错误: {str(e)}", "ERROR", "red")
            error_msg = f"快速导出过程中出错: {str(e)}"
            QMessageBox.critical(self, "错误", error_msg)
            self.set_progress_status("快速导出失败", "red")
            self.is_slicing = False
            if self.debug_window:
                self.debug_window.interrupt_btn.setEnabled(False)
    
    def check_all_file_conflicts(self, save_dir, base_name, file_format, direction, method, param):
        """检查所有可能产生的文件冲突"""
        if not self.image:
            return []
            
        try:
            self.debug_log("开始检查文件冲突", "CHECK", "blue")
            width, height = self.image.size
            conflict_files = []
            
            if method == "按大小切片":
                if direction == "横向":
                    current_x = 0
                    i = 1
                    while current_x < width:
                        slice_width = min(param, width - current_x)
                        filename = f"{base_name}_{i}_{current_x}.{file_format}"
                        save_path = os.path.join(save_dir, filename)
                        if os.path.exists(save_path):
                            conflict_files.append(filename)
                        current_x += slice_width
                        i += 1
                else:  
                    current_y = 0
                    i = 1
                    while current_y < height:
                        slice_height = min(param, height - current_y)
                        filename = f"{base_name}_{i}_{current_y}.{file_format}"
                        save_path = os.path.join(save_dir, filename)
                        if os.path.exists(save_path):
                            conflict_files.append(filename)
                        current_y += slice_height
                        i += 1
            else:  
                if direction == "横向":
                    base_width = width // param
                    remainder = width % param
                    current_x = 0
                    for i in range(param):
                        slice_width = base_width + 1 if i < remainder else base_width
                        filename = f"{base_name}_{i+1}_{current_x}.{file_format}"
                        save_path = os.path.join(save_dir, filename)
                        if os.path.exists(save_path):
                            conflict_files.append(filename)
                        current_x += slice_width
                else:  
                    base_height = height // param
                    remainder = height % param
                    current_y = 0
                    for i in range(param):
                        slice_height = base_height + 1 if i < remainder else base_height
                        filename = f"{base_name}_{i+1}_{current_y}.{file_format}"
                        save_path = os.path.join(save_dir, filename)
                        if os.path.exists(save_path):
                            conflict_files.append(filename)
                        current_y += slice_height
            
            self.debug_log(f"文件冲突检查完成，发现 {len(conflict_files)} 个冲突文件", "CHECK", "green" if not conflict_files else "orange")
            return conflict_files
        except Exception as e:
            self.debug_log(f"文件冲突检查失败: {str(e)}", "ERROR", "red")
            return []
    
    def slice_by_size(self, direction, size, save_dir, base_name, file_format, conflict_files):
        """按大小切片"""
        try:
            self.debug_log("开始按大小切片", "SLICE", "blue")
            width, height = self.image.size
            total_slices = 0
            
            if direction == "横向":
                self.debug_log("横向切片", "SLICE", "blue")
                current_x = 0
                i = 1
                while current_x < width:
                    if self.debug_window and self.debug_window.is_task_interrupted:
                        self.debug_log("切片操作被用户中断", "WARNING", "orange")
                        return False
                        
                    slice_width = min(size, width - current_x)
                    box = (current_x, 0, current_x + slice_width, height)
                    slice_img = self.image.crop(box)
                    
                    filename = f"{base_name}_{i}_{current_x}.{file_format}"
                    save_path = os.path.join(save_dir, filename)

                    is_overwrite = os.path.exists(save_path)
                    self.debug_log(f"保存切片 {i}: {filename} {'(替换)' if is_overwrite else ''}", "SAVE", "orange" if is_overwrite else "green")
                    
                    self.save_slice_image(slice_img, save_path, file_format)
                    
                    current_x += slice_width
                    i += 1
                    total_slices += 1

                    progress = int((current_x / width) * 100)
                    self.update_progress(progress, f"正在切片... {progress}%")
                    QApplication.processEvents()
            else:  
                self.debug_log("纵向切片", "SLICE", "blue")
                current_y = 0
                i = 1
                while current_y < height:
                    if self.debug_window and self.debug_window.is_task_interrupted:
                        self.debug_log("切片操作被用户中断", "WARNING", "orange")
                        return False
                        
                    slice_height = min(size, height - current_y)
                    box = (0, current_y, width, current_y + slice_height)
                    slice_img = self.image.crop(box)
                    
                    filename = f"{base_name}_{i}_{current_y}.{file_format}"
                    save_path = os.path.join(save_dir, filename)

                    is_overwrite = os.path.exists(save_path)
                    self.debug_log(f"保存切片 {i}: {filename} {'(替换)' if is_overwrite else ''}", "SAVE", "orange" if is_overwrite else "green")
                    
                    self.save_slice_image(slice_img, save_path, file_format)
                    
                    current_y += slice_height
                    i += 1
                    total_slices += 1

                    progress = int((current_y / height) * 100)
                    self.update_progress(progress, f"正在切片... {progress}%")
                    QApplication.processEvents()
            
            self.debug_log(f"切片完成，共 {total_slices} 个文件", "SLICE", "green")
            self.set_progress_status("切片完成", "green")
            return True
        except Exception as e:
            self.debug_log(f"按大小切片失败: {str(e)}", "ERROR", "red")
            raise Exception(f"按大小切片失败: {str(e)}")
    
    def slice_by_count(self, direction, count, save_dir, base_name, file_format, conflict_files):
        """按数量切片"""
        try:
            self.debug_log("开始按数量切片", "SLICE", "blue")
            width, height = self.image.size
            
            if direction == "横向":
                self.debug_log("横向切片", "SLICE", "blue")
                base_width = width // count
                remainder = width % count
                current_x = 0
                
                for i in range(count):
                    if self.debug_window and self.debug_window.is_task_interrupted:
                        self.debug_log("切片操作被用户中断", "WARNING", "orange")
                        return False
                        
                    slice_width = base_width + 1 if i < remainder else base_width
                    box = (current_x, 0, current_x + slice_width, height)
                    slice_img = self.image.crop(box)
                    
                    filename = f"{base_name}_{i+1}_{current_x}.{file_format}"
                    save_path = os.path.join(save_dir, filename)

                    is_overwrite = os.path.exists(save_path)
                    self.debug_log(f"保存切片 {i+1}: {filename} {'(替换)' if is_overwrite else ''}", "SAVE", "orange" if is_overwrite else "green")
                    
                    self.save_slice_image(slice_img, save_path, file_format)
                    
                    current_x += slice_width

                    progress = int(((i + 1) / count) * 100)
                    self.update_progress(progress, f"正在切片... {progress}%")
                    QApplication.processEvents()
            else:  
                self.debug_log("纵向切片", "SLICE", "blue")
                base_height = height // count
                remainder = height % count
                current_y = 0
                
                for i in range(count):
                    if self.debug_window and self.debug_window.is_task_interrupted:
                        self.debug_log("切片操作被用户中断", "WARNING", "orange")
                        return False
                        
                    slice_height = base_height + 1 if i < remainder else base_height
                    box = (0, current_y, width, current_y + slice_height)
                    slice_img = self.image.crop(box)
                    
                    filename = f"{base_name}_{i+1}_{current_y}.{file_format}"
                    save_path = os.path.join(save_dir, filename)

                    is_overwrite = os.path.exists(save_path)
                    self.debug_log(f"保存切片 {i+1}: {filename} {'(替换)' if is_overwrite else ''}", "SAVE", "orange" if is_overwrite else "green")
                    
                    self.save_slice_image(slice_img, save_path, file_format)
                    
                    current_y += slice_height
                    
                    progress = int(((i + 1) / count) * 100)
                    self.update_progress(progress, f"正在切片... {progress}%")
                    QApplication.processEvents()
            
            self.debug_log(f"切片完成，共 {count} 个文件", "SLICE", "green")
            self.set_progress_status("切片完成", "green")
            return True
        except Exception as e:
            self.debug_log(f"按数量切片失败: {str(e)}", "ERROR", "red")
            raise Exception(f"按数量切片失败: {str(e)}")
    
    def save_slice_image(self, image, save_path, file_format):
        """保存切片图片"""
        try:
            if file_format == "jpg":
                if image.mode != "RGB":
                    image = image.convert("RGB")
                image.save(save_path, "JPEG", quality=95)
            else:
                image.save(save_path, "PNG")
        except Exception as e:
            self.debug_log(f"图片保存失败: {str(e)}", "ERROR", "red")
            raise Exception(f"保存切片失败: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # 检查命令行参数
    dropped_files = parse_command_line()
    
    app.setApplicationName("E-IMG Slices")
    app.setApplicationVersion("V1.4-Beta")
    app.setOrganizationName("E-IMG")

    app.setStyleSheet("""
        QMainWindow {
            background-color: #f5f5f5;
        }
        QGroupBox {
            font-weight: bold;
            border: 1px solid #cccccc;
            border-radius: 5px;
            margin-top: 1ex;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QPushButton {
            background-color: #0071bc;
            border: none;
            color: white;
            padding: 8px 16px;
            text-align: center;
            text-decoration: none;
            font-size: 14px;
            margin: 4px 2px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #29abe2;
        }
        QPushButton:pressed {
            background-color: #0059a8;
        }
        QPushButton:disabled {
            background-color: #cccccc;
            color: #666666;
        }
        /* 快速导出按钮特殊样式 - 通过对象名称选择器 */
        QPushButton#quick_export_btn {
            background-color: #FFA726;
        }
        QPushButton#quick_export_btn:hover {
            background-color: #FFCC80;
        }
        QPushButton#quick_export_btn:pressed {
            background-color: #FFD54F;
        }
        QPushButton#quick_export_btn:disabled {
            background-color: #cccccc;
            color: #666666;
        }
        QTextEdit, QLineEdit {
            padding: 5px;
            border: 1px solid #cccccc;
            border-radius: 3px;
            background-color: white;
        }
        QSpinBox, QComboBox {
            padding: 5px;
            border: 1px solid #cccccc;
            border-radius: 3px;
            background-color: white;
        }
        QSpinBox::up-button, QSpinBox::down-button {
            border: none;
            border-radius: 2px;
            width: 16px;
        }
        QSpinBox::up-button:hover, QSpinBox::down-button:hover {
        }
        QSpinBox::up-arrow, QSpinBox::down-arrow {
            width: 8px;
            height: 8px;
        }
        QComboBox::drop-down {
            border: none;
            width: 20px;
        }
        QComboBox::down-arrow {
            width: 10px;
            height: 10px;
        }
        QLabel {
            color: #333333;
        }
        QProgressBar {
            border: 1px solid #cccccc;
            border-radius: 3px;
            text-align: center;
        }
        QProgressBar::chunk {
            border-radius: 2px;
        }
    """)
    
    window = ImageSlicer()
    window.show()
    
    debug_print("应用程序启动完成")
    
    # 设置快速导出按钮的对象名称，以便样式表选择器匹配
    window.quick_export_btn.setObjectName("quick_export_btn")
    
    # 如果有拖放的文件，则使用第一个文件并执行快速导出
    if dropped_files:
        debug_print(f"检测到拖放文件: {dropped_files[0]}")
        # 延迟执行，确保窗口完全加载
        QTimer.singleShot(500, lambda: window.handle_dropped_file_at_startup(dropped_files[0]))
    
    sys.exit(app.exec_())