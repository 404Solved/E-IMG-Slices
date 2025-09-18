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
                             QProgressBar, QAction, QDialog, QTextBrowser, QDialogButtonBox)
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QPixmap, QImage, QIcon, QDesktopServices, QTextCursor, QColor

# 全局调试标志
DEBUG_MODE = True

def get_resource_path(relative_path):
    """获取资源的绝对路径，支持打包后的exe文件"""
    try:
        # 打包后的资源路径
        base_path = sys._MEIPASS
    except Exception:
        # 开发时的资源路径
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# 配置类
class AppConfig:
    def __init__(self):
        self.config_dir = os.path.join(os.getenv('APPDATA'), 'E-IMG Slices')
        self.config_file = os.path.join(self.config_dir, 'config.json')
        self.debug_mode = False  # 强制debug模式为关闭状态
        self.load_large_preview = False
        
        # 确保配置目录存在
        os.makedirs(self.config_dir, exist_ok=True)
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                    # 不加载debug_mode，始终保持关闭状态
                    self.load_large_preview = config_data.get('load_large_preview', False)
            except:
                # 如果配置文件损坏，使用默认值
                self.load_large_preview = False
    
    def save_config(self):
        """保存配置文件"""
        config_data = {
            'debug_mode': False,  # 始终保存为关闭状态
            'load_large_preview': self.load_large_preview
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
        except:
            pass

# 关于窗口
class AboutWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于E-IMG Slices")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setModal(True)
        
        # 尝试加载about.jpg
        about_image_path = get_resource_path("about.jpg")
        
        layout = QVBoxLayout(self)
        
        if os.path.exists(about_image_path):
            # 显示图片
            pixmap = QPixmap(about_image_path)
            image_label = QLabel()
            image_label.setPixmap(pixmap)
            image_label.setAlignment(Qt.AlignCenter)
            image_label.setCursor(Qt.PointingHandCursor)  # 设置手型光标
            image_label.mousePressEvent = self.close  # 点击图片关闭窗口
            
            layout.addWidget(image_label)
            
            # 设置窗口大小适应图片
            self.resize(pixmap.width(), pixmap.height())
        else:
            # 如果图片不存在，显示错误信息
            error_label = QLabel("文件丢失：about哪去了？")
            error_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(error_label)
            
            # 添加关闭按钮
            close_btn = QPushButton("关闭")
            close_btn.clicked.connect(self.close)
            layout.addWidget(close_btn)
    
    def close(self, event=None):
        """重写close方法以支持点击图片关闭"""
        super().accept()

# Debug日志窗口
class DebugLogWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Debug日志输出")
        self.setGeometry(100, 100, 800, 500)
        
        # 创建布局
        layout = QVBoxLayout(self)
        
        # 日志显示区域
        self.log_text = QTextBrowser()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        self.open_log_dir_btn = QPushButton("打开日志文件夹")
        self.clear_log_btn = QPushButton("清空日志")
        self.close_btn = QPushButton("关闭")
        
        button_layout.addWidget(self.open_log_dir_btn)
        button_layout.addWidget(self.clear_log_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)
        
        layout.addLayout(button_layout)
        
        # 连接信号
        self.open_log_dir_btn.clicked.connect(self.open_log_directory)
        self.clear_log_btn.clicked.connect(self.clear_log)
        self.close_btn.clicked.connect(self.accept)
        
        # 日志文件路径
        self.log_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'E-IMG Slices Log')
        os.makedirs(self.log_dir, exist_ok=True)
        self.log_file = os.path.join(self.log_dir, f"debug_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        
    def append_log(self, message):
        """添加日志信息"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}"
        
        # 在窗口中显示
        self.log_text.append(log_message)
        
        # 写入文件
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_message + '\n')
        except:
            pass
        
        # 自动滚动到底部
        self.log_text.moveCursor(QTextCursor.End)
    
    def open_log_directory(self):
        """打开日志文件夹"""
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.log_dir))
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()

def debug_print(*args, **kwargs):
    """调试输出函数"""
    if DEBUG_MODE:
        print("[DEBUG]", *args, **kwargs)
        # 检查stdout是否存在再调用flush
        if sys.stdout is not None:
            sys.stdout.flush()

class ImageSlicer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.image = None
        self.image_path = None
        self.last_save_dir = None
        self.max_display_size = 4000  # 最大显示尺寸限制
        self.config = AppConfig()
        self.debug_window = None
        
        debug_print("程序启动，初始化界面...")
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle('E-IMG 图片切片工具')
        self.setGeometry(100, 100, 900, 700)
        
        # 设置程序图标
        logo_path = get_resource_path("logo.ico")
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))
            # 设置任务栏图标
            if hasattr(QApplication, 'setWindowIcon'):
                QApplication.setWindowIcon(QIcon(logo_path))
        
        # 创建顶部菜单栏
        self.createMenuBar()
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QHBoxLayout(central_widget)
        
        # 左侧图片显示区域
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        # 创建滚动区域用于显示大图片
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("border: 2px dashed gray; background-color: #f0f0f0;")
        
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setText("拖放图片到这里或点击\"加载图片\"")
        self.image_label.setMinimumSize(400, 400)
        
        self.scroll_area.setWidget(self.image_label)
        left_layout.addWidget(self.scroll_area)
        
        # 右侧控制区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 图片信息显示
        info_group = QGroupBox("图片信息")
        info_layout = QVBoxLayout(info_group)
        
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setMaximumHeight(120)
        info_layout.addWidget(self.info_text)
        
        # 切片设置
        settings_group = QGroupBox("切片设置")
        settings_layout = QVBoxLayout(settings_group)
        
        # 文件命名
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("文件命名:"))
        self.name_edit = QLineEdit("E-IMG slices")
        name_layout.addWidget(self.name_edit)
        
        # 文件格式
        format_layout = QHBoxLayout()
        format_layout.addWidget(QLabel("输出格式:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["JPG", "PNG"])
        format_layout.addWidget(self.format_combo)
        
        # 方向选择
        direction_layout = QHBoxLayout()
        direction_layout.addWidget(QLabel("切片方向:"))
        self.direction_combo = QComboBox()
        self.direction_combo.addItems(["横向", "纵向"])
        self.direction_combo.setCurrentIndex(1)  # 默认选择纵向
        direction_layout.addWidget(self.direction_combo)
        
        # 切片方式选择
        method_layout = QHBoxLayout()
        method_layout.addWidget(QLabel("切片方式:"))
        self.method_combo = QComboBox()
        self.method_combo.addItems(["按大小切片", "按数量切片"])
        method_layout.addWidget(self.method_combo)
        
        # 参数输入
        param_layout = QHBoxLayout()
        param_layout.addWidget(QLabel("参数值:"))
        self.param_spin = QSpinBox()
        self.param_spin.setMinimum(1)
        self.param_spin.setMaximum(10000)
        param_layout.addWidget(self.param_spin)
        
        # 按钮
        button_layout = QHBoxLayout()
        self.load_btn = QPushButton("加载图片")
        self.slice_btn = QPushButton("开始切片")
        self.slice_btn.setEnabled(False)
        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.slice_btn)
        
        # 添加到设置组
        settings_layout.addLayout(name_layout)
        settings_layout.addLayout(format_layout)
        settings_layout.addLayout(direction_layout)
        settings_layout.addLayout(method_layout)
        settings_layout.addLayout(param_layout)
        settings_layout.addLayout(button_layout)
        
        # 预览信息
        preview_group = QGroupBox("切片预览信息")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMaximumHeight(100)
        preview_layout.addWidget(self.preview_text)
        
        # 日志输出和进度条
        log_group = QGroupBox("操作日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.set_progress_status("就绪", "gray")
        log_layout.addWidget(self.progress_bar)
        
        # 添加到右侧布局
        right_layout.addWidget(info_group, 1)
        right_layout.addWidget(settings_group, 2)
        right_layout.addWidget(preview_group, 1)
        right_layout.addWidget(log_group, 2)
        
        # 添加到主布局
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 500])
        
        main_layout.addWidget(splitter)
        
        # 创建底部状态栏
        self.createStatusBar()
        
        # 连接信号和槽
        self.load_btn.clicked.connect(self.load_image)
        self.slice_btn.clicked.connect(self.slice_image)
        self.method_combo.currentIndexChanged.connect(self.update_param_hint)
        self.direction_combo.currentIndexChanged.connect(self.update_preview_if_enabled)
        self.param_spin.valueChanged.connect(self.update_preview_if_enabled)
        
        # 初始化参数提示
        self.update_param_hint()
        
        # 启用拖放功能
        self.setAcceptDrops(True)
        self.image_label.dragEnterEvent = self.dragEnterEvent
        self.image_label.dropEvent = self.dropEvent
        
        debug_print("界面初始化完成")
        
    def closeEvent(self, event):
        """重写关闭事件，确保debug窗口也关闭"""
        if self.debug_window:
            self.debug_window.close()
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
        
        # 功能菜单
        function_menu = menubar.addMenu('功能')
        
        # Debug功能
        self.debug_action = QAction('Debug', self)
        self.debug_action.setCheckable(True)
        self.debug_action.setChecked(False)  # 始终设置为关闭状态
        self.debug_action.triggered.connect(self.toggle_debug)
        function_menu.addAction(self.debug_action)
        
        # 加载大尺寸图片预览图功能
        self.large_preview_action = QAction('加载大尺寸图片预览图', self)
        self.large_preview_action.setCheckable(True)
        self.large_preview_action.setChecked(self.config.load_large_preview)
        self.large_preview_action.triggered.connect(self.toggle_large_preview)
        function_menu.addAction(self.large_preview_action)
        
        # 关于菜单（下拉列表）
        about_menu = menubar.addMenu('帮助')
        
        # 关于
        about_action = about_menu.addAction('关于E-IMG Slices')
        about_action.triggered.connect(self.openAboutWindow)
        
        # GitHub
        github_action = about_menu.addAction('GitHub项目页')
        github_action.triggered.connect(self.openGithubUrl)
        
    def toggle_debug(self, checked):
        """切换Debug模式"""
        if checked:
            self.debug_window = DebugLogWindow(self)
            self.debug_window.show()
            self.debug_window.append_log("Debug模式已启用")
        else:
            if self.debug_window:
                self.debug_window.append_log("Debug模式已禁用")
                self.debug_window.close()
                self.debug_window = None
    
    def toggle_large_preview(self, checked):
        """切换大尺寸图片预览功能"""
        if checked:
            # 显示警告对话框
            reply = QMessageBox.warning(
                self, 
                "注意！", 
                "继续启用\"显示大文件预览图\"功能？\n此功能尚未完善，部分图片将无法导入或导致程序崩溃\n如果遇到问题请关闭此功能",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.config.load_large_preview = True
                self.large_preview_action.setChecked(True)
            else:
                self.config.load_large_preview = False
                self.large_preview_action.setChecked(False)
        else:
            self.config.load_large_preview = False
        
        self.config.save_config()
        
    def openAboutWindow(self):
        """打开关于E-IMG Slices窗口"""
        about_window = AboutWindow(self)
        about_window.exec_()
        
    def createStatusBar(self):
        self.statusbar = QStatusBar()
        self.statusbar.showMessage("E-IMG Slices | V1.0-Release")
        self.setStatusBar(self.statusbar)
        
    def openGithubUrl(self):
        QDesktopServices.openUrl(QUrl('https://github.com/404Solved/E-IMG-Slices'))
        self.restoreStatusBar()
        
    def restoreStatusBar(self):
        # 使用QTimer延迟恢复状态栏文字，确保菜单点击完成后恢复
        from PyQt5.QtCore import QTimer
        QTimer.singleShot(100, lambda: self.statusbar.showMessage("E-IMG Slices | V1.0-Release"))
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            
    def dropEvent(self, event):
        try:
            debug_print("拖放事件开始")
            if event.mimeData().hasUrls():
                url = event.mimeData().urls()[0]
                file_path = url.toLocalFile()
                debug_print(f"拖放文件路径: {file_path}")
                
                # 检查文件是否为图片
                if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.gif', '.webp')):
                    self.load_image_from_path(file_path)
                else:
                    QMessageBox.warning(self, "错误", "请拖放有效的图片文件")
                    debug_print("拖放的文件不是图片格式")
        except Exception as e:
            debug_print(f"拖放事件异常: {traceback.format_exc()}")
            QMessageBox.critical(self, "错误", f"拖放操作失败: {str(e)}")
                
    def load_image_from_path(self, file_path):
        try:
            debug_print(f"开始加载图片: {file_path}")
            self.set_progress_status("正在导入...", "blue")
            QApplication.processEvents()  # 更新UI
            
            # 使用更安全的图片加载方式
            try:
                debug_print("尝试打开图片文件...")
                self.image = Image.open(file_path)
                debug_print("图片打开成功，开始验证...")
                
                # 尝试读取图片信息来验证图片是否有效
                self.image.verify()
                debug_print("图片验证成功")
                
                # 重新打开图片，因为verify()会关闭文件
                self.image = Image.open(file_path)
                debug_print("图片重新打开成功")
                
            except Exception as e:
                debug_print(f"图片验证失败: {traceback.format_exc()}")
                raise Exception(f"图片文件损坏或格式不受支持: {str(e)}")
            
            self.image_path = file_path
            debug_print("图片基本信息设置完成")
            
            # 检查图片尺寸是否过大
            width, height = self.image.size
            if (width > 20000 or height > 20000) and not self.config.load_large_preview:
                debug_print(f"图片尺寸过大: {width}x{height}")
                self.append_log("警告: 图片尺寸较大，预览功能受限，但仍可进行切片操作", "WARNING", "orange")
                self.image_label.setText("图片尺寸较大，预览功能受限\n但仍可进行切片操作")
                self.image_label.setStyleSheet("color: orange; font-size: 14px;")
            else:
                # 尝试显示图片，如果失败则继续处理
                try:
                    debug_print("尝试显示图片...")
                    self.display_image()
                    debug_print("图片显示成功")
                except Exception as e:
                    debug_print(f"图片显示失败: {traceback.format_exc()}")
                    self.append_log(f"图片预览失败: {str(e)}", "WARNING", "orange")
                    self.image_label.setText("图片预览失败，但仍可进行切片操作")
                    self.image_label.setStyleSheet("color: orange;")
            
            self.show_image_info()
            self.slice_btn.setEnabled(True)
            self.append_log(f"已加载图片: {os.path.basename(file_path)}", "INFO", "black")
            
            # 尝试预览切片信息
            try:
                debug_print("尝试预览切片信息...")
                self.preview_slice_info()
                debug_print("切片预览成功")
            except Exception as e:
                debug_print(f"切片预览失败: {traceback.format_exc()}")
                self.preview_text.clear()
                self.append_preview("预览失败: 无法计算切片信息", "red")
                self.append_preview("您仍然可以尝试继续切片操作", "orange")
                self.append_log(f"切片预览失败: {str(e)}", "WARNING", "orange")
            
            self.set_progress_status("就绪", "gray")
            debug_print("图片加载流程完成")
            
        except Exception as e:
            debug_print(f"图片加载过程中出现严重错误: {traceback.format_exc()}")
            error_msg = f"无法加载图片: {str(e)}"
            QMessageBox.critical(self, "错误", error_msg)
            self.append_log(error_msg, "ERROR", "red")
            self.set_progress_status("导入失败", "red")
            
            # 重置图片状态
            self.image = None
            self.image_path = None
            self.slice_btn.setEnabled(False)
            self.image_label.setText("拖放图片到这里或点击\"加载图片\"")
            self.image_label.setStyleSheet("")
            self.info_text.clear()
            self.preview_text.clear()
    
    def update_param_hint(self):
        if self.method_combo.currentText() == "按大小切片":
            self.param_spin.setSuffix(" 像素")
        else:
            self.param_spin.setSuffix(" 份")
        self.update_preview_if_enabled()
    
    def update_preview_if_enabled(self):
        if self.image:  # 只要有图片就预览，不需要按钮启用状态
            try:
                self.preview_slice_info()
            except Exception as e:
                debug_print(f"实时预览失败: {traceback.format_exc()}")
                self.preview_text.clear()
                self.append_preview("预览失败: 无法计算切片信息", "red")
                self.append_preview("您仍然可以尝试继续切片操作", "orange")
    
    def load_image(self):
        try:
            debug_print("打开文件对话框...")
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择图片", "", 
                "图片文件 (*.png *.jpg *.jpeg *.bmp *.tiff *.gif *.webp)"
            )
            
            if file_path:
                debug_print(f"用户选择文件: {file_path}")
                self.load_image_from_path(file_path)
        except Exception as e:
            debug_print(f"文件对话框异常: {traceback.format_exc()}")
    
    def display_image(self):
        if self.image:
            try:
                debug_print("开始处理图片显示...")
                width, height = self.image.size
                
                # 检查图片尺寸，如果过大则显示缩略图
                if (width > self.max_display_size or height > self.max_display_size) and not self.config.load_large_preview:
                    debug_print(f"图片尺寸过大 ({width}x{height})，显示缩略图")
                    # 计算缩放比例
                    scale = min(self.max_display_size / width, self.max_display_size / height)
                    new_width = int(width * scale)
                    new_height = int(height * scale)
                    
                    # 创建缩略图
                    thumbnail = self.image.copy()
                    thumbnail.thumbnail((new_width, new_height), Image.Resampling.LANCZOS)
                    
                    # 转换为QImage
                    if thumbnail.mode == "RGB":
                        qimage = QImage(thumbnail.tobytes(), thumbnail.width, thumbnail.height, QImage.Format_RGB888)
                    elif thumbnail.mode == "RGBA":
                        qimage = QImage(thumbnail.tobytes(), thumbnail.width, thumbnail.height, QImage.Format_RGBA8888)
                    else:
                        thumbnail = thumbnail.convert("RGB")
                        qimage = QImage(thumbnail.tobytes(), thumbnail.width, thumbnail.height, QImage.Format_RGB888)
                    
                    debug_print(f"缩略图尺寸: {thumbnail.width}x{thumbnail.height}")
                else:
                    # 创建高质量的图像显示
                    img = self.image.copy()
                    debug_print(f"图片复制完成，模式: {img.mode}, 尺寸: {img.size}")
                    
                    # 转换为QImage
                    if img.mode == "RGB":
                        debug_print("转换为RGB格式")
                        qimage = QImage(img.tobytes(), img.width, img.height, QImage.Format_RGB888)
                    elif img.mode == "RGBA":
                        debug_print("转换为RGBA格式")
                        qimage = QImage(img.tobytes(), img.width, img.height, QImage.Format_RGBA8888)
                    elif img.mode == "L":  # 灰度图像
                        debug_print("转换为灰度格式")
                        qimage = QImage(img.tobytes(), img.width, img.height, QImage.Format_Grayscale8)
                    elif img.mode == "P":  # 调色板模式
                        debug_print("调色板模式，转换为RGB")
                        img = img.convert("RGB")
                        qimage = QImage(img.tobytes(), img.width, img.height, QImage.Format_RGB888)
                    else:
                        debug_print(f"未知模式 {img.mode}，尝试转换为RGB")
                        img = img.convert("RGB")
                        qimage = QImage(img.tobytes(), img.width, img.height, QImage.Format_RGB888)
                
                debug_print("QImage创建成功，创建QPixmap...")
                # 创建高质量的pixmap
                pixmap = QPixmap.fromImage(qimage)
                debug_print("QPixmap创建成功")
                
                # 设置图片大小策略，保持原始比例但允许滚动查看
                self.image_label.setPixmap(pixmap)
                self.image_label.setScaledContents(False)  # 不自动缩放
                
                # 如果显示的是缩略图，添加提示信息
                if (width > self.max_display_size or height > self.max_display_size) and not self.config.load_large_preview:
                    self.image_label.setText(f"缩略图预览\n原始尺寸: {width}×{height}")
                    self.image_label.setStyleSheet("color: gray; font-size: 12px;")
                else:
                    self.image_label.setMinimumSize(pixmap.width(), pixmap.height())
                    self.image_label.setStyleSheet("")  # 清除之前的样式
                
                debug_print("图片标签设置完成")
                
                # 如果图片太大，调整滚动区域
                if pixmap.width() > 800 or pixmap.height() > 600:
                    self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
                    self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
                    debug_print("启用滚动条")
                else:
                    self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                    self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                    debug_print("禁用滚动条")
                    
            except Exception as e:
                debug_print(f"图片显示过程中出错: {traceback.format_exc()}")
                raise Exception(f"无法显示图片: {str(e)}")
    
    def show_image_info(self):
        if self.image:
            try:
                debug_print("开始获取图片信息...")
                width, height = self.image.size
                mode = self.image.mode
                info = f"文件名: {os.path.basename(self.image_path)}\n"
                info += f"尺寸: {width} × {height} 像素\n"
                info += f"颜色模式: {mode}\n"
                
                # 获取DPI信息（只显示一个值）
                dpi = self.image.info.get('dpi', (72, 72))
                info += f"分辨率: {dpi[0]} PPI\n"
                
                # 文件大小
                file_size = os.path.getsize(self.image_path)
                info += f"文件大小: {file_size / 1024:.2f} KB"
                
                self.info_text.setPlainText(info)
                debug_print("图片信息显示完成")
                
            except Exception as e:
                debug_print(f"获取图片信息失败: {traceback.format_exc()}")
                self.info_text.setPlainText(f"文件名: {os.path.basename(self.image_path)}\n无法获取完整图片信息")
                self.append_log(f"图片信息获取失败: {str(e)}", "WARNING", "orange")
    
    def append_log(self, message, log_type="INFO", color="black"):
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        timestamp = f"[{current_time} {log_type}]"
        
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        # 设置时间戳样式（加粗）
        self.log_text.setFontWeight(75)  # 使用数值而不是QFont.Bold
        self.log_text.setTextColor(QColor(color))
        self.log_text.insertPlainText(timestamp + " ")
        
        # 设置消息样式（正常字体）
        self.log_text.setFontWeight(50)  # 使用数值而不是QFont.Normal
        self.log_text.insertPlainText(message + "\n")
        
        # 恢复默认样式
        self.log_text.setTextColor(QColor("black"))
        self.log_text.ensureCursorVisible()
        
        # 如果Debug模式开启，同时输出到Debug窗口
        if self.debug_window:
            self.debug_window.append_log(f"{log_type}: {message}")
    
    def append_preview(self, message, color="black"):
        cursor = self.preview_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        # 设置文本颜色
        self.preview_text.setTextColor(QColor(color))
        self.preview_text.insertPlainText(message + "\n")
        
        # 恢复默认颜色
        self.preview_text.setTextColor(QColor("black"))
        self.preview_text.ensureCursorVisible()
    
    def preview_slice_info(self):
        if not self.image:
            return
            
        try:
            debug_print("开始计算切片预览...")
            direction = self.direction_combo.currentText()
            method = self.method_combo.currentText()
            param = self.param_spin.value()
            width, height = self.image.size
            
            debug_print(f"切片参数 - 方向: {direction}, 方法: {method}, 参数: {param}, 尺寸: {width}x{height}")
            
            self.preview_text.clear()
            
            if method == "按大小切片":
                if direction == "横向":
                    total_slices = (width + param - 1) // param  # 向上取整
                    remainder = width % param
                    
                    self.append_preview(f"将切成 {total_slices} 份", "black")
                    if remainder > 0:
                        self.append_preview(f"前 {total_slices-1} 份尺寸: {param}×{height} 像素", "black")
                        self.append_preview(f"最后 1 份尺寸: {remainder}×{height} 像素", "black")
                        self.append_preview("末尾切片不满足要求，将直接输出", "orange")
                    else:
                        self.append_preview(f"每份尺寸: {param}×{height} 像素", "black")
                else:  # 纵向
                    total_slices = (height + param - 1) // param  # 向上取整
                    remainder = height % param
                    
                    self.append_preview(f"将切成 {total_slices} 份", "black")
                    if remainder > 0:
                        self.append_preview(f"前 {total_slices-1} 份尺寸: {width}×{param} 像素", "black")
                        self.append_preview(f"最后 1 份尺寸: {width}×{remainder} 像素", "black")
                        self.append_preview("末尾切片不满足要求，将直接输出", "orange")
                    else:
                        self.append_preview(f"每份尺寸: {width}×{param} 像素", "black")
            else:  # 按数量切片
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
                else:  # 纵向
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
            debug_print("切片预览计算完成")
        except Exception as e:
            debug_print(f"切片预览计算失败: {traceback.format_exc()}")
            raise Exception(f"预览计算失败: {str(e)}")
    
    def slice_image(self):
        if not self.image:
            return
            
        try:
            debug_print("开始切片操作...")
            direction = self.direction_combo.currentText()
            method = self.method_combo.currentText()
            param = self.param_spin.value()
            base_name = self.name_edit.text().strip() or "E-IMG slices"
            file_format = self.format_combo.currentText().lower()
            
            debug_print(f"切片设置 - 方向: {direction}, 方法: {method}, 参数: {param}, 名称: {base_name}, 格式: {file_format}")
            
            # 选择保存目录
            save_dir = QFileDialog.getExistingDirectory(self, "选择保存目录", self.last_save_dir or "")
            if not save_dir:
                debug_print("用户取消选择目录")
                return
            
            self.last_save_dir = save_dir
            debug_print(f"保存目录: {save_dir}")
            
            # 检查是否有文件冲突
            conflict_files = self.check_all_file_conflicts(save_dir, base_name, file_format, direction, method, param)
            
            if conflict_files:
                debug_print(f"发现 {len(conflict_files)} 个文件冲突")
                reply = QMessageBox.question(self, "文件冲突", 
                                            f"发现 {len(conflict_files)} 个文件已存在，是否全部覆盖？",
                                            QMessageBox.Yes | QMessageBox.No,
                                            QMessageBox.No)
                if reply != QMessageBox.Yes:
                    self.append_log("操作已取消", "WARNING", "orange")
                    self.set_progress_status("操作取消", "orange")
                    debug_print("用户取消覆盖操作")
                    return
                
            self.set_progress_status("正在切片...", "blue")
            QApplication.processEvents()  # 更新UI
            
            if method == "按大小切片":
                debug_print("使用按大小切片方法")
                result = self.slice_by_size(direction, param, save_dir, base_name, file_format, conflict_files)
            else:
                debug_print("使用按数量切片方法")
                result = self.slice_by_count(direction, param, save_dir, base_name, file_format, conflict_files)
            
            if result:
                debug_print("切片操作完成")
                # 创建完成对话框
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("完成")
                msg_box.setText("图片切片已完成！")
                msg_box.setIcon(QMessageBox.Information)
                
                # 添加查看按钮
                view_button = msg_box.addButton("查看", QMessageBox.ActionRole)
                msg_box.addButton(QMessageBox.Ok)
                
                msg_box.exec_()
                
                if msg_box.clickedButton() == view_button:
                    debug_print("用户点击查看按钮")
                    # 打开文件资源管理器
                    QDesktopServices.openUrl(QUrl.fromLocalFile(save_dir))
            
        except Exception as e:
            debug_print(f"切片过程中出现严重错误: {traceback.format_exc()}")
            error_msg = f"切片过程中出错: {str(e)}"
            QMessageBox.critical(self, "错误", error_msg)
            self.append_log(error_msg, "ERROR", "red")
            self.set_progress_status("切片失败", "red")
    
    def check_all_file_conflicts(self, save_dir, base_name, file_format, direction, method, param):
        """检查所有可能产生的文件冲突"""
        if not self.image:
            return []
            
        try:
            debug_print("开始检查文件冲突...")
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
                else:  # 纵向
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
            else:  # 按数量切片
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
                else:  # 纵向
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
            
            debug_print(f"文件冲突检查完成，发现 {len(conflict_files)} 个冲突文件")
            return conflict_files
        except Exception as e:
            debug_print(f"文件冲突检查失败: {traceback.format_exc()}")
            self.append_log(f"文件冲突检查失败: {str(e)}", "WARNING", "orange")
            return []
    
    def slice_by_size(self, direction, size, save_dir, base_name, file_format, conflict_files):
        if not self.image:
            return False
            
        try:
            debug_print(f"开始按大小切片，方向: {direction}, 大小: {size}")
            width, height = self.image.size
            has_warning = False
            
            if direction == "横向":
                debug_print("水平切片")
                # 水平切片
                slices = []
                current_x = 0
                
                while current_x < width:
                    slice_width = min(size, width - current_x)
                    box = (current_x, 0, current_x + slice_width, height)
                    slice_img = self.image.crop(box)
                    slices.append((slice_img, current_x))
                    current_x += slice_width
                    
                debug_print(f"生成 {len(slices)} 个切片")
                # 保存切片
                total_slices = len(slices)
                for i, (img, x_pos) in enumerate(slices):
                    filename = f"{base_name}_{i+1}_{x_pos}.{file_format}"
                    save_path = os.path.join(save_dir, filename)
                    
                    if file_format == "jpg":
                        img.convert("RGB").save(save_path, "JPEG", quality=95)
                    else:
                        img.save(save_path)
                        
                    self.append_log(f"已保存: {filename} ({img.size[0]}×{img.size[1]})", "INFO", "black")
                    
                    # 更新进度
                    progress = int((i + 1) / total_slices * 100)
                    self.update_progress(progress, f"切片中... {i+1}/{total_slices}")
                    QApplication.processEvents()
                    
                    # 检查最后一片是否小于指定尺寸
                    if slice_width < size and len(slices) > 1 and not has_warning:
                        has_warning = True
                    
            else:  # 纵向
                debug_print("垂直切片")
                # 垂直切片
                slices = []
                current_y = 0
                
                while current_y < height:
                    slice_height = min(size, height - current_y)
                    box = (0, current_y, width, current_y + slice_height)
                    slice_img = self.image.crop(box)
                    slices.append((slice_img, current_y))
                    current_y += slice_height
                    
                debug_print(f"生成 {len(slices)} 个切片")
                # 保存切片
                total_slices = len(slices)
                for i, (img, y_pos) in enumerate(slices):
                    filename = f"{base_name}_{i+1}_{y_pos}.{file_format}"
                    save_path = os.path.join(save_dir, filename)
                    
                    if file_format == "jpg":
                        img.convert("RGB").save(save_path, "JPEG", quality=95)
                    else:
                        img.save(save_path)
                        
                    self.append_log(f"已保存: {filename} ({img.size[0]}×{img.size[1]})", "INFO", "black")
                    
                    # 更新进度
                    progress = int((i + 1) / total_slices * 100)
                    self.update_progress(progress, f"切片中... {i+1}/{total_slices}")
                    QApplication.processEvents()
                    
                    # 检查最后一片是否小于指定尺寸
                    if slice_height < size and len(slices) > 1 and not has_warning:
                        has_warning = True
            
            # 显示警告信息（只显示一次）
            if has_warning:
                self.append_log("警告: 末尾切片不满足要求，已输出，请检查", "WARNING", "orange")
                debug_print("检测到末尾切片不满足要求")
            
            # 显示文件冲突警告
            if conflict_files:
                self.append_log("警告: 以下重复文件已被覆盖: " + ", ".join(conflict_files), "WARNING", "orange")
                debug_print(f"覆盖了 {len(conflict_files)} 个文件")
            
            self.append_log("切片成功！", "SUCCESS", "green")
            self.set_progress_status("操作完成", "green")
            debug_print("按大小切片完成")
            return True
            
        except Exception as e:
            debug_print(f"按大小切片失败: {traceback.format_exc()}")
            raise Exception(f"按大小切片失败: {str(e)}")
    
    def slice_by_count(self, direction, count, save_dir, base_name, file_format, conflict_files):
        if not self.image:
            return False
            
        try:
            debug_print(f"开始按数量切片，方向: {direction}, 数量: {count}")
            width, height = self.image.size
            has_warning = False
            
            if direction == "横向":
                debug_print("水平切片")
                # 计算每片宽度（考虑余数分散）
                base_width = width // count
                remainder = width % count
                debug_print(f"基础宽度: {base_width}, 余数: {remainder}")
                
                slice_widths = [base_width + 1 if i < remainder else base_width for i in range(count)]
                debug_print(f"切片宽度分布: {slice_widths}")
                
                # 生成切片
                slices = []
                current_x = 0
                
                for i, sw in enumerate(slice_widths):
                    box = (current_x, 0, current_x + sw, height)
                    slice_img = self.image.crop(box)
                    slices.append((slice_img, current_x, i+1))
                    current_x += sw
                    
                debug_print(f"生成 {len(slices)} 个切片")
                # 保存切片
                total_slices = len(slices)
                for i, (img, x_pos, idx) in enumerate(slices):
                    filename = f"{base_name}_{idx}_{x_pos}.{file_format}"
                    save_path = os.path.join(save_dir, filename)
                    
                    if file_format == "jpg":
                        img.convert("RGB").save(save_path, "JPEG", quality=95)
                    else:
                        img.save(save_path)
                        
                    self.append_log(f"已保存: {filename} ({img.size[0]}×{img.size[1]})", "INFO", "black")
                    
                    # 更新进度
                    progress = int((i + 1) / total_slices * 100)
                    self.update_progress(progress, f"切片中... {i+1}/{total_slices}")
                    QApplication.processEvents()
                    
                if remainder > 0 and not has_warning:
                    has_warning = True
                    
            else:  # 纵向
                debug_print("垂直切片")
                # 计算每片高度（考虑余数分散")
                base_height = height // count
                remainder = height % count
                debug_print(f"基础高度: {base_height}, 余数: {remainder}")
                
                slice_heights = [base_height + 1 if i < remainder else base_height for i in range(count)]
                debug_print(f"切片高度分布: {slice_heights}")
                
                # 生成切片
                slices = []
                current_y = 0
                
                for i, sh in enumerate(slice_heights):
                    box = (0, current_y, width, current_y + sh)
                    slice_img = self.image.crop(box)
                    slices.append((slice_img, current_y, i+1))
                    current_y += sh
                    
                debug_print(f"生成 {len(slices)} 个切片")
                # 保存切片
                total_slices = len(slices)
                for i, (img, y_pos, idx) in enumerate(slices):
                    filename = f"{base_name}_{idx}_{y_pos}.{file_format}"
                    save_path = os.path.join(save_dir, filename)
                    
                    if file_format == "jpg":
                        img.convert("RGB").save(save_path, "JPEG", quality=95)
                    else:
                        img.save(save_path)
                        
                    self.append_log(f"已保存: {filename} ({img.size[0]}×{img.size[1]})", "INFO", "black")
                    
                    # 更新进度
                    progress = int((i + 1) / total_slices * 100)
                    self.update_progress(progress, f"切片中... {i+1}/{total_slices}")
                    QApplication.processEvents()
                    
                if remainder > 0 and not has_warning:
                    has_warning = True
            
            # 显示警告信息（只显示一次）
            if has_warning:
                self.append_log("提示: 已采用余数分散分配处理", "INFO", "orange")
                debug_print("采用余数分散分配处理")
            
            # 显示文件冲突警告
            if conflict_files:
                self.append_log("警告: 以下重复文件已被覆盖: " + ", ".join(conflict_files), "WARNING", "orange")
                debug_print(f"覆盖了 {len(conflict_files)} 个文件")
            
            self.append_log("切片成功！", "SUCCESS", "green")
            self.set_progress_status("操作完成", "green")
            debug_print("按数量切片完成")
            return True
            
        except Exception as e:
            debug_print(f"按数量切片失败: {traceback.format_exc()}")
            raise Exception(f"按数量切片失败: {str(e)}")

if __name__ == '__main__':
    # 设置异常钩子来捕获所有未处理的异常
    def exception_hook(exctype, value, traceback_obj):
        """
        全局异常处理函数
        """
        print("=" * 60)
        print("发生未处理的异常:")
        print(f"异常类型: {exctype}")
        print(f"异常值: {value}")
        print("跟踪信息:")
        import traceback
        traceback.print_tb(traceback_obj)
        print("=" * 60)
        
        # 调用默认的异常钩子
        sys.__excepthook__(exctype, value, traceback_obj)
    
    # 设置全局异常钩子
    sys.excepthook = exception_hook
    
    debug_print("启动应用程序...")
    app = QApplication(sys.argv)
    window = ImageSlicer()
    window.show()
    debug_print("进入主事件循环")
    sys.exit(app.exec_())