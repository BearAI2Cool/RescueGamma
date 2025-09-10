import json
import logging
import os

from PySide6.QtCore import QThread
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QAction, QFont, QIcon
from PySide6.QtGui import QPainter, QLinearGradient, QColor, QBrush, QPen
from PySide6.QtWidgets import (QMainWindow, QVBoxLayout, QLineEdit, QComboBox,
                               QTextEdit, QFileDialog, QMessageBox, QMenu, QProgressBar, QFrame, QScrollArea, QDialog,
                               QSizePolicy, QGroupBox)
from PySide6.QtWidgets import QPushButton
from PySide6.QtWidgets import QWidget, QHBoxLayout, QLabel

from modules.config_manager import ConfigManager
from modules.ppt_processor import PPTProcessor
from ui.color_picker import ColorPicker
from ui.font_config import FontConfig


def init_logger():

    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(log_dir, 'main_window.log')

    logger = logging.getLogger('MainWindow')
    logger.setLevel(logging.INFO)

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    from logging.handlers import RotatingFileHandler
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=3,
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)

    return logger

logger = init_logger()


class GradientButton(QPushButton):
    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        gradient = QLinearGradient(0, 0, self.width(), 0)
        gradient.setColorAt(0, QColor("#9A6FDC"))
        gradient.setColorAt(1, QColor("#73C6E1"))

        painter.setPen(QPen(QBrush(gradient), 0))
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.text())


class GradientLabel(QLabel):
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        gradient = QLinearGradient(0, 0, self.width(), 0)
        gradient.setColorAt(0, QColor("#9A6FDC"))
        gradient.setColorAt(1, QColor("#73C6E1"))

        painter.setPen(QPen(QBrush(gradient), 0))
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, self.text())


class ProcessThread(QThread):
    progress = Signal(str)
    finished = Signal(bool)

    def __init__(self, processor, input_path, output_path, gradient_config, font_size):
        super().__init__()
        self.processor = processor
        self.input_path = input_path
        self.output_path = output_path
        self.gradient_config = gradient_config
        self.font_size = font_size

    def run(self):
        try:
            self.progress.emit("开始处理PPT文件...")
            success = self.processor.process_ppt(
                self.input_path,
                self.output_path,
                self.gradient_config,
                self.font_size
            )
            self.finished.emit(success)
        except Exception as e:
            self.progress.emit(f"处理失败: {str(e)}")
            self.finished.emit(False)


class PreviewLabel(QLabel):
    def __init__(self):
        super().__init__()
        self.preview_texts = [
            "渐变预览"
        ]
        self.setText("\n".join(self.preview_texts))
        self.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        self.gradient_config = []
        self.font_size = "12"
        self.setStyleSheet("""
            background-color: #f8f8f8; 
            border: 1px solid #ddd;
            padding: 10px;
        """)
        self.setWordWrap(True)
        self.setFixedHeight(240)
        self.update_preview([], "12")

    def update_preview(self, gradient_config, font_size):
        self.gradient_config = gradient_config
        self.font_size = font_size

        try:
            size = int(float(font_size)) if font_size else 12
        except:
            size = 12

        font = QFont("SimHei", size, QFont.Bold)
        self.setFont(font)

        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        painter.fillRect(self.rect(), QColor("#f8f8f8"))

        try:
            size = int(float(self.font_size)) if self.font_size else 12
        except:
            size = 12
        font = QFont("Microsoft YaHei", size)
        painter.setFont(font)

        if self.gradient_config and len(self.gradient_config) >= 2:
            gradient = QLinearGradient(0, 0, self.width(), 0)

            for config in self.gradient_config:
                pos = config['position'] / 100000.0
                color = config['color']

                if color.startswith('#'):
                    gradient.setColorAt(pos, QColor(color))
                elif color == 'accent1':
                    gradient.setColorAt(pos, QColor("#5B9BD5"))
                else:
                    gradient.setColorAt(pos, QColor("#333333"))

            painter.setPen(QPen(QBrush(gradient), 0))
        else:
            painter.setPen(QColor("#333333"))

        text_rect = self.rect().adjusted(10, 10, -10, -10)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter, self.text())


class SortButton(QPushButton):

    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        layout.setAlignment(Qt.AlignCenter)

        self.asc_label = QLabel("升序")
        self.asc_label.setStyleSheet("border: none; background: transparent;")

        separator = QLabel("/")
        separator.setStyleSheet("border: none; background: transparent;")

        self.desc_label = QLabel("降序")
        self.desc_label.setStyleSheet("border: none; background: transparent;")

        layout.addWidget(self.asc_label)
        layout.addWidget(separator)
        layout.addWidget(self.desc_label)

        self.setLayout(layout)


    def update_style(self, sort_order, is_dark_mode):
        if is_dark_mode:
            bold_color = "white"
            normal_color = "gray"
        else:
            bold_color = "black"
            normal_color = "gray"

        base_style = "border: none; background: transparent;"

        if sort_order == 0:
            self.asc_label.setStyleSheet(f"font-weight: bold; color: {bold_color}; {base_style}")
            self.desc_label.setStyleSheet(f"font-weight: normal; color: {normal_color}; {base_style}")
        else:
            self.asc_label.setStyleSheet(f"font-weight: normal; color: {normal_color}; {base_style}")
            self.desc_label.setStyleSheet(f"font-weight: bold; color: {bold_color}; {base_style}")


class ConfigListDialog(QDialog):
    config_deleted = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配置清单")
        self.setMinimumSize(450, 300)

        from modules.config_manager import ConfigManager
        self.config_manager = ConfigManager()

        self.is_dark_mode = self.parent().is_dark_mode if self.parent() else False

        self.configs = {}
        self.scroll_layout = None

        self.load_configs()
        self.setup_ui()
        self.update_sort_button_style()
        self.populate_config_list()
        self.resize_to_content()

    def setup_ui(self):

        main_layout = QVBoxLayout(self)

        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("border: none;")
        scroll_area.setWidgetResizable(True)

        scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(scroll_widget)
        self.scroll_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

        bottom_layout = QHBoxLayout()

        self.sort_button = SortButton()

        self.close_btn = QPushButton("关闭")

        bottom_layout.addWidget(self.sort_button)
        bottom_layout.addWidget(self.close_btn)

        self.sort_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.close_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        self.sort_button.setMinimumHeight(35)
        self.close_btn.setMinimumHeight(35)

        self.sort_button.clicked.connect(self.toggle_sort_order)
        self.close_btn.clicked.connect(self.accept)

        main_layout.addLayout(bottom_layout)
        self.setLayout(main_layout)

    def clear_layout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

    def populate_config_list(self):
        self.clear_layout(self.scroll_layout)

        if not self.configs:
            no_config_label = QLabel("尚未保存任何配置")
            no_config_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            no_config_label.setStyleSheet("border: none; padding: 50px; font-size: 14px;")
            self.scroll_layout.addWidget(no_config_label)
        else:
            for i, (font_size, config_data) in enumerate(self.configs.items()):
                scheme_frame = self.create_scheme_widget(i + 1, font_size, config_data)
                self.scroll_layout.addWidget(scheme_frame)

        self.resize_to_content()

    def create_scheme_widget(self, scheme_num, font_size, config_data):
        scheme_frame = QFrame()
        scheme_frame.setFrameStyle(QFrame.Shape.Box)
        scheme_frame.setStyleSheet("QFrame { border: 1px solid #ccc; margin: 5px; padding: 8px; }")
        scheme_layout = QVBoxLayout(scheme_frame)
        scheme_layout.setSpacing(3)

        title_layout = QHBoxLayout()
        scheme_title = QLabel(f"# 方案{scheme_num}：")
        scheme_title.setStyleSheet("font-weight: bold; font-size: 14px; border: none;")
        title_layout.addWidget(scheme_title)
        title_layout.addStretch()

        delete_btn = QPushButton("删除")
        delete_btn.setFixedSize(45, 25)
        delete_btn.setStyleSheet("background-color: #ff4444; color: white; border: none; font-size: 14px;")
        delete_btn.clicked.connect(lambda checked, fs=font_size: self.delete_scheme(fs))
        title_layout.addWidget(delete_btn)
        scheme_layout.addLayout(title_layout)

        font_size_info = QLabel(f"字号：{font_size}")
        font_size_info.setStyleSheet("border: none; margin-left: 10px;")
        scheme_layout.addWidget(font_size_info)

        font_name_info = QLabel(f"字体：{config_data.get('font_name', 'N/A')}")
        font_name_info.setStyleSheet("border: none; margin-left: 10px;")
        scheme_layout.addWidget(font_name_info)

        gradient_title = QLabel("渐变配置方案：")
        gradient_title.setStyleSheet("border: none; margin-left: 10px;")
        scheme_layout.addWidget(gradient_title)

        for item in config_data.get('gradient_config', []):
            config_label = QLabel(f"位置：{item['position']:>6}    颜色：{item['color']}")
            config_label.setStyleSheet(
                "border: none; margin-left: 20px; font-family: 'Consolas', 'Courier New', monospace;")
            scheme_layout.addWidget(config_label)

        return scheme_frame

    def update_sort_button_style(self):
        sort_order = self.config_manager.get_sort_order()
        self.sort_button.update_style(sort_order, self.is_dark_mode)

    def toggle_sort_order(self):
        current_order = self.config_manager.get_sort_order()
        new_order = 1 - current_order

        self.config_manager.set_sort_order(new_order)

        self.update_sort_button_style()

        self.load_configs()
        self.populate_config_list()

    def load_configs(self):
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 获取排序状态 (0=升序, 1=降序)
                    sort_order = self.config_manager.get_sort_order()
                    # 关键修复：使用 float 作为 key 进行数值排序
                    sorted_keys = sorted(data.keys(), key=lambda x: float(x), reverse=bool(sort_order))
                    self.configs = {k: data[k] for k in sorted_keys}
            else:
                self.configs = {}
        except (json.JSONDecodeError, ValueError, Exception) as e:
            logging.error(f"加载或排序配置失败: {str(e)}")
            self.configs = {}

    def resize_to_content(self):
        content_height = 100
        for config_data in self.configs.values():
            content_height += 80
            content_height += len(config_data.get('gradient_config', [])) * 20
        max_height = min(600, max(300, content_height))
        self.resize(self.width(), max_height)

    def delete_scheme(self, font_size):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.NoIcon)
        msg.setWindowTitle('确认删除')
        msg.setText(f'确定要删除字号为 {font_size} 的方案吗？')
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if msg.exec() == QMessageBox.StandardButton.Yes:
            if font_size in self.configs:
                del self.configs[font_size]
                try:
                    with open('config.json', 'w', encoding='utf-8') as f:
                        json.dump(self.configs, f, ensure_ascii=False, indent=2)
                    self.config_deleted.emit()
                    self.accept()
                    success_msg = QMessageBox(self)
                    success_msg.setWindowTitle("成功")
                    success_msg.setText("方案删除成功")
                    success_msg.exec()
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"删除方案失败: {str(e)}")


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于")
        self.setFixedSize(320, 250)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(20, 20, 20, 20)

        title = GradientLabel("- RescueGamma -")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 10px; border: none;")
        layout.addWidget(title)

        layout.addSpacing(15)

        developer = QLabel("开发者: Moss_Go")
        developer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        developer.setStyleSheet("border: none;")
        layout.addWidget(developer)

        version = QLabel("版本号: 3.0")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version.setStyleSheet("border: none;")
        layout.addWidget(version)

        layout.addSpacing(10)

        date = QLabel("发布日期: 2025年9月10日")
        date.setAlignment(Qt.AlignmentFlag.AlignCenter)
        date.setStyleSheet("border: none;")
        layout.addWidget(date)

        layout.addSpacing(15)

        description = QLabel("懂知识付费的AI课研助理")
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setStyleSheet("font-style: italic; border: none;")
        layout.addWidget(description)

        layout.addSpacing(25)

        ok_btn = QPushButton("确定")
        ok_btn.setMinimumHeight(30)
        ok_btn.setStyleSheet("""
            QPushButton {
                color: #000000;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                padding: 5px 15px;  # 增加内边距（上下5px，左右15px）
                font-size: 12px;
                min-width: 45px;  # 设置最小宽度
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """)
        ok_btn.clicked.connect(self.accept)
        layout.addWidget(ok_btn)

        self.setLayout(layout)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.processor = PPTProcessor(status_callback=self.add_status_message)
        self.config_manager = ConfigManager()
        self.is_dark_mode = self.config_manager.get_dark_mode()
        self.gradient_config = [
            {'position': 0, 'color': '#9A6FDC'},
            {'position': 74000, 'color': 'accent1'},
            {'position': 83000, 'color': 'accent1'},
            {'position': 100000, 'color': '#73C6E1'}
        ]
        self.setup_ui()
        self.load_config()
        self.apply_theme()
        last_ppt_path = self.config_manager.get_last_ppt_path()
        if last_ppt_path:
            self.gamma_path.setText(last_ppt_path)

    def setup_ui(self):
        self.setWindowTitle("RescueGamma 3.0")
        self.setFixedSize(800, 700)

        icon_path = "ico/RescueGamma_32.ico"
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.create_menu_bar()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)

        layout.addSpacing(20)

        title_label = GradientLabel("Gamma PPT 修复")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; border: none;")
        layout.addWidget(title_label)

        file_layout = QVBoxLayout()
        file_layout.setSpacing(10)

        gamma_layout = QHBoxLayout()
        gamma_label = QLabel("Gamma PPT路径 :")
        gamma_label.setFixedWidth(120)
        gamma_label.setStyleSheet("border: none; font-weight: bold;")
        gamma_layout.addWidget(gamma_label)

        self.gamma_path = QLineEdit()
        self.gamma_path.setPlaceholderText("仅支持.pptx格式")
        gamma_layout.addWidget(self.gamma_path)

        browse_gamma_btn = QPushButton("浏览")
        browse_gamma_btn.setFixedWidth(60)
        browse_gamma_btn.clicked.connect(self.browse_gamma_file)
        gamma_layout.addWidget(browse_gamma_btn)
        file_layout.addLayout(gamma_layout)

        output_layout = QHBoxLayout()
        output_label = QLabel("PPT保存路径 :")
        output_label.setFixedWidth(120)
        output_label.setStyleSheet("border: none; font-weight: bold;")
        output_layout.addWidget(output_label)

        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("仅支持.pptx格式")
        output_layout.addWidget(self.output_path)

        browse_output_btn = QPushButton("浏览")
        browse_output_btn.setFixedWidth(60)
        browse_output_btn.clicked.connect(self.browse_output_path)
        output_layout.addWidget(browse_output_btn)
        file_layout.addLayout(output_layout)

        layout.addLayout(file_layout)
        layout.addSpacing(15)

        config_layout = QHBoxLayout()
        config_layout.setSpacing(20)

        left_group = QGroupBox()
        left_group.setStyleSheet("QGroupBox { border: 1px solid #ccc; }")
        left_group.setFixedWidth(350)
        left_layout = QVBoxLayout(left_group)
        left_layout.setSpacing(15)
        left_layout.setContentsMargins(10, 10, 10, 10)

        Gamma_label = QLabel("原Gamma文本")
        Gamma_label.setStyleSheet("font-weight: bold; border: none;")
        left_layout.addWidget(Gamma_label)

        h_layout = QHBoxLayout()

        font_label = QLabel("字体")
        font_label.setStyleSheet("border: none;")
        self.font_name_edit = QLineEdit()
        self.font_name_edit.setPlaceholderText("试试填：Sora")
        h_layout.addWidget(font_label)
        h_layout.addWidget(self.font_name_edit)

        size_label = QLabel("字号")
        size_label.setStyleSheet("border: none;")

        h_layout.addWidget(size_label)

        self.font_size_combo = QComboBox()
        self.font_size_combo.setEditable(True)
        self.font_size_combo.setMinimumWidth(120)
        font_sizes = ["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "32", "36", "44", "48", "72"]
        self.font_size_combo.addItems(font_sizes)
        self.font_size_combo.setCurrentText("44.5")
        self.font_size_combo.currentTextChanged.connect(self.update_preview)

        h_layout.addWidget(self.font_size_combo)

        left_layout.addLayout(h_layout)
        left_layout.addSpacing(10)

        gradient_label = QLabel("渐变配置")
        gradient_label.setStyleSheet("font-weight: bold; border: none;")
        left_layout.addWidget(gradient_label)

        gradient_frame = QFrame()
        gradient_frame.setFrameStyle(QFrame.Shape.Box)
        gradient_frame_layout = QVBoxLayout(gradient_frame)
        gradient_frame_layout.setSpacing(8)

        self.gradient_entries = []
        self.position_inputs = []
        self.color_inputs = []

        for i, config in enumerate(self.gradient_config):
            entry_widget = QWidget()
            entry_layout = QHBoxLayout(entry_widget)
            entry_layout.setContentsMargins(5, 2, 5, 2)

            pos_label = QLabel("位置:")
            pos_label.setFixedWidth(30)
            pos_label.setStyleSheet("border: none;")
            entry_layout.addWidget(pos_label)

            pos_input = QLineEdit(str(config['position']))
            pos_input.setFixedWidth(60)
            pos_input.textChanged.connect(self.on_gradient_config_changed)

            pos_input.setStyleSheet("border: 1px solid #ccc; padding: 2px;")
            entry_layout.addWidget(pos_input)
            self.position_inputs.append(pos_input)

            color_label = QLabel("颜色:")
            color_label.setFixedWidth(30)
            color_label.setStyleSheet("border: none;")
            entry_layout.addWidget(color_label)

            color_input = QLineEdit(str(config['color']))
            color_input.setFixedWidth(80)
            color_input.textChanged.connect(self.on_gradient_config_changed)

            color_input.setStyleSheet("border: 1px solid #ccc; padding: 2px;")
            entry_layout.addWidget(color_input)
            self.color_inputs.append(color_input)

            color_preview = QLabel()
            color_preview.setFixedSize(20, 20)
            if config['color'].startswith('#'):
                color_preview.setStyleSheet(f"background-color: {config['color']}; border: 1px solid black;")
            else:
                color_preview.setStyleSheet("background-color: lightgray; border: 1px solid black;")
            entry_layout.addWidget(color_preview)

            entry_layout.addStretch()
            gradient_frame_layout.addWidget(entry_widget)

            self.gradient_entries.append({
                'widget': entry_widget,
                'preview': color_preview,
                'pos_input': pos_input,
                'color_input': color_input
            })

        left_layout.addWidget(gradient_frame)

        config_btn_layout = QHBoxLayout()
        save_config_btn = QPushButton("保存配置")
        save_config_btn.clicked.connect(self.save_config)
        config_btn_layout.addWidget(save_config_btn)

        show_config_btn = QPushButton("配置清单")
        show_config_btn.clicked.connect(self.show_config_list)
        config_btn_layout.addWidget(show_config_btn)

        left_layout.addLayout(config_btn_layout)
        left_layout.addStretch()

        config_layout.addWidget(left_group)

        right_group = QGroupBox()
        right_group.setStyleSheet("QGroupBox { border: 1px solid #ccc; }")
        right_layout = QVBoxLayout(right_group)
        right_layout.setSpacing(15)
        right_layout.setContentsMargins(10, 10, 10, 10)

        preview_label = QLabel("效果预览")
        preview_label.setStyleSheet("font-weight: bold; border: none;")
        right_layout.addWidget(preview_label)

        self.preview_label = PreviewLabel()
        self.preview_label.setMinimumHeight(250)
        right_layout.addWidget(self.preview_label)

        right_layout.addStretch()
        config_layout.addWidget(right_group)

        layout.addLayout(config_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_text = QTextEdit()
        self.status_text.setMaximumHeight(100)
        self.status_text.setPlaceholderText("处理状态将在这里显示...")
        layout.addWidget(self.status_text)

        button_layout = QHBoxLayout()
        button_layout.addStretch()

        start_btn = GradientButton("一键渐变")
        start_btn.setStyleSheet(
            "QPushButton { font-size: 18px; font-weight: bold; padding: 10px 25px; }")
        start_btn.clicked.connect(self.start_processing)
        button_layout.addWidget(start_btn)

        button_layout.addStretch()
        layout.addLayout(button_layout)

        self.update_preview()

    def create_menu_bar(self):
        menubar = self.menuBar()

        menu = QMenu("菜单", self)

        color_action = QAction("颜色查询", self)
        color_action.triggered.connect(self.show_color_picker)
        menu.addAction(color_action)

        font_action = QAction("字体配置", self)
        font_action.triggered.connect(self.show_font_config)
        menu.addAction(font_action)

        gradient_action = QAction("渐变提取", self)
        gradient_action.triggered.connect(self.show_gradient_extractor)
        menu.addAction(gradient_action)

        menu.addSeparator()

        self.dark_mode_action = QAction("切换深色模式", self)
        self.dark_mode_action.triggered.connect(self.toggle_dark_mode)
        menu.addAction(self.dark_mode_action)

        reset_action = QAction("恢复默认设置", self)
        reset_action.triggered.connect(self.restore_default_settings)
        menu.addAction(reset_action)

        menubar.addMenu(menu)

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        menubar.addAction(about_action)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        menubar.setCornerWidget(spacer, Qt.Corner.TopRightCorner)

    def show_gradient_extractor(self):
        from ui.gradient_extractor import GradientExtractor

        self.gradient_window = GradientExtractor(self)
        self.gradient_window.setWindowFlags(Qt.Window)
        self.gradient_window.setWindowTitle("渐变手动提取器")
        self.gradient_window.setWindowModality(Qt.WindowModal)

        self.gradient_window.apply_theme(self.is_dark_mode)
        self.gradient_window.resize(800, 600)
        self.gradient_window.center_window()
        self.gradient_window.show()

    def restore_default_settings(self):
        if self.config_manager.restore_default_settings():
            self.load_config()
            self.is_dark_mode = self.config_manager.get_dark_mode()
            self.apply_theme()
            self.gamma_path.clear()
            self.add_status_message("已恢复默认设置，配置文件已备份到Backup目录")
        else:
            self.add_status_message("恢复默认设置失败")

    def show_about(self):
        dialog = AboutDialog(self)
        dialog.exec()

    def toggle_dark_mode(self):
        self.is_dark_mode = not self.is_dark_mode
        self.apply_theme()
        self.config_manager.set_dark_mode(self.is_dark_mode)

    def apply_theme(self):
        if self.is_dark_mode:
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #2b2b2b;
                    color: #ffffff;
                }
                QWidget {
                    background-color: #2b2b2b;
                    color: #ffffff;
                }
                QLabel {
                    color: #ffffff;
                }
                QLineEdit {
                    background-color: #3c3c3c;
                    border: 1px solid #555555;
                    color: #ffffff;
                    padding: 5px;
                }
                QPushButton {
                    background-color: #404040;
                    border: 1px solid #555555;
                    color: #ffffff;
                    padding: 5px 10px;
                }
                QPushButton:hover {
                    background-color: #505050;
                }
                QPushButton:pressed {
                    background-color: #353535;
                }
                QComboBox {
                    background-color: #3c3c3c;
                    border: 1px solid #555555;
                    color: #ffffff;
                    padding: 5px;
                }
                QComboBox::drop-down {
                    border: none;
                }
                QComboBox::down-arrow {
                    border: none;
                }
                QTextEdit {
                    background-color: #3c3c3c;
                    border: 1px solid #555555;
                    color: #ffffff;
                }
                QFrame {
                    border: 1px solid #555555;
                }
                QGroupBox {
                    border: 1px solid #555555;
                }
                QMenuBar {
                    background-color: #2b2b2b;
                    color: #ffffff;
                }
                QMenuBar::item {
                    background-color: transparent;
                    padding: 4px 8px;
                }
                QMenuBar::item:selected {
                    background-color: #404040;
                }
                QMenu {
                    background-color: #3c3c3c;
                    color: #ffffff;
                    border: 1px solid #555555;
                }
                QMenu::item {
                    padding: 6px 20px;
                    text-align: center;
                }
                QMenu::item:selected {
                    background-color: #505050;
                }
                QCheckBox {
                    color: #ffffff;
                }
                QCheckBox::indicator {
                    width: 15px;
                    height: 15px;
                    border: 1px solid #555555;
                    border-radius: 3px;
                    background-color: #ffffff;
                }
                QCheckBox::indicator:checked {
                    background-color: #555555;
                }
            """)
            self.dark_mode_action.setText("切换亮色模式")
        else:
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #ffffff;
                    color: #000000;
                }
                QWidget {
                    background-color: #ffffff;
                    color: #000000;
                }
                QLabel {
                    color: #000000;
                }
                QLineEdit {
                    background-color: #ffffff;
                    border: 1px solid #cccccc;
                    color: #000000;
                    padding: 5px;
                }
                QPushButton {
                    background-color: #f0f0f0;
                    border: 1px solid #cccccc;
                    color: #000000;
                    padding: 5px 10px;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
                QPushButton:pressed {
                    background-color: #d0d0d0;
                }
                QComboBox {
                    background-color: #ffffff;
                    border: 1px solid #cccccc;
                    color: #000000;
                    padding: 5px;
                }
                QTextEdit {
                    background-color: #ffffff;
                    border: 1px solid #cccccc;
                    color: #000000;
                }
                QFrame {
                    border: 1px solid #cccccc;
                }
                QGroupBox {
                    border: 1px solid #cccccc;
                }
                QMenuBar {
                    background-color: #f0f0f0;
                    color: #000000;
                }
                QMenuBar::item {
                    background-color: transparent;
                    padding: 4px 8px;
                }
                QMenuBar::item:selected {
                    background-color: #e0e0e0;
                }
                QMenu {
                    background-color: #ffffff;
                    color: #000000;
                    border: 1px solid #cccccc;
                }
                QMenu::item {
                    padding: 6px 20px;
                    text-align: center;
                }
                QMenu::item:selected {
                    background-color: #e0e0e0;
                }
                QCheckBox {
                    color: #000000;
                }
                QCheckBox::indicator {
                    width: 15px;
                    height: 15px;
                    border: 1px solid #555555;
                    border-radius: 3px;
                }
                QCheckBox::indicator:checked {
                    background-color: #555555;
                }
            """)
            self.dark_mode_action.setText("切换深色模式")

    def on_gradient_config_changed(self):
        try:
            for i, entry in enumerate(self.gradient_entries):
                position = entry['pos_input'].text()
                color = entry['color_input'].text()

                if i < len(self.gradient_config):
                    try:
                        self.gradient_config[i]['position'] = int(position) if position.isdigit() else 0
                    except:
                        pass
                    self.gradient_config[i]['color'] = color

                if color.startswith('#') and len(color) == 7:
                    entry['preview'].setStyleSheet(f"background-color: {color}; border: 1px solid black;")
                else:
                    entry['preview'].setStyleSheet("background-color: lightgray; border: 1px solid black;")

            self.update_preview()
        except Exception as e:
            print(f"更新渐变配置时出错: {e}")

    def show_color_picker(self):
        dialog = ColorPicker(self)
        dialog.exec()

    def show_font_config(self):
        dialog = FontConfig(self)
        dialog.font_changed.connect(self.on_font_config_changed)
        dialog.exec()

    def on_font_config_changed(self, config):
        self.add_status_message(
            f"字体配置: {config['old_font']} {config['old_size']} -> {config['new_font']} {config['new_size']}")

    def show_gradient_extractor(self):
        from ui.gradient_extractor import GradientExtractor

        self.gradient_window = GradientExtractor(self)
        self.gradient_window.setWindowFlags(Qt.Window)
        self.gradient_window.setWindowTitle("渐变手动提取器")
        self.gradient_window.setWindowModality(Qt.WindowModal)

        self.gradient_window.apply_theme(self.is_dark_mode)

        self.gradient_window.resize(800, 600)
        self.gradient_window.center_window()

        self.gradient_window.show()

    def browse_gamma_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Gamma PPT文件", "", "PowerPoint文件 (*.pptx)"
        )
        if file_path:
            self.gamma_path.setText(file_path)
            self.config_manager.set_last_ppt_path(file_path)

    def browse_output_path(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "选择保存路径", "", "PowerPoint文件 (*.pptx)"
        )
        if file_path:
            self.output_path.setText(file_path)

    def update_preview(self):
        font_size = self.font_size_combo.currentText()
        self.preview_label.update_preview(self.gradient_config, font_size)

    def save_config(self):
        font_size = self.font_size_combo.currentText().strip()
        if not font_size:
            msg = QMessageBox(self)
            msg.setWindowTitle("错误")
            msg.setText("请指定字体大小")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 150px;
                    min-height: 50px;
                }
            """)
            msg.exec()
            return

        configs = {}
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    configs = data
        except Exception as e:
            print(f"读取配置失败: {str(e)}")

        configs[font_size] = {
            'gradient_config': self.gradient_config.copy(),
            'font_name': self.font_name_edit.text()
        }

        try:
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(configs, f, ensure_ascii=False, indent=2)

            msg = QMessageBox(self)
            msg.setWindowTitle("成功")
            if font_size in configs and len([k for k in configs.keys() if k == font_size]) > 0:
                msg.setText(f"字号 {font_size} 的配置已更新")
            else:
                msg.setText(f"字号 {font_size} 的新方案已保存")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    padding: 20px;
                    min-width: 200px;
                    min-height: 50px;
                }
            """)
            msg.exec()
        except Exception as e:
            msg = QMessageBox(self)
            msg.setWindowTitle("错误")
            msg.setText(f"保存配置失败: {str(e)}")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 200px;
                    min-height: 50px;
                }
            """)
            msg.exec()

    def load_config(self):
        try:
            if os.path.exists('config.json'):
                with open('config.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)

                    if data:
                        first_font_size = list(data.keys())[0]
                        first_config = data[first_font_size]
                        self.font_size_combo.setCurrentText(first_font_size)
                        self.gradient_config = first_config.get('gradient_config', self.gradient_config)
                        self.font_name_edit.setText(first_config.get('font_name', 'Sora'))

                    for i, grad_config in enumerate(self.gradient_config):
                        if i < len(self.gradient_entries):
                            self.gradient_entries[i]['pos_input'].setText(str(grad_config['position']))
                            self.gradient_entries[i]['color_input'].setText(str(grad_config['color']))

                    self.update_preview()
        except Exception as e:
            logging.error(f"加载配置失败: {str(e)}")

    def show_config_list(self):
        dialog = ConfigListDialog(self)
        dialog.config_deleted.connect(self.on_config_deleted)
        dialog.exec()

    def on_config_deleted(self):
        self.load_config()

    def add_status_message(self, message):
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.append(f"[{timestamp}] {message}")

    def start_processing(self):
        gamma_file = self.gamma_path.text().strip()
        output_file = self.output_path.text().strip()
        font_size = self.font_size_combo.currentText().strip()

        if not gamma_file or not os.path.exists(gamma_file):
            msg = QMessageBox(self)
            msg.setWindowTitle("错误")
            msg.setText("请选择有效的Gamma PPT文件")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 220px;
                    min-height: 50px;
                }
            """)
            msg.exec()
            return

        if not output_file:
            msg = QMessageBox(self)
            msg.setWindowTitle("错误")
            msg.setText("请指定输出文件路径")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 180px;
                    min-height: 50px;
                }
            """)
            msg.exec()
            return

        if not font_size:
            msg = QMessageBox(self)
            msg.setWindowTitle("错误")
            msg.setText("请指定字体大小")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 150px;
                    min-height: 50px;
                }
            """)
            msg.exec()
            return

        font_config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'font_config.json')
        try:
            if not os.path.exists(font_config_path):
                raise FileNotFoundError("字体配置文件不存在")
            with open(font_config_path, 'r', encoding='utf-8') as f:
                font_config = json.load(f)
            if not font_config:
                raise ValueError("字体配置为空")
        except Exception:
            msg = QMessageBox(self)
            msg.setWindowTitle("警告")
            msg.setText("未检查到有效字体配置")
            msg.setStyleSheet(""
                              "QMessageBox {\n"
                              "    background-color: white;\n"
                              "    border: none;\n"
                              "}\
                  "
                              "QMessageBox QLabel {\n"
                              "    border: none;\n"
                              "    qproperty-alignment: 'AlignCenter';\n"
                              "    padding: 20px;\n"
                              "}\
                  "
                              "QMessageBox QPushButton {\n"
                              "    border: none;\n"
                              "    background-color: #f0f0f0;\n"
                              "    padding: 5px 15px;\n"
                              "}\
                  "
                              "QMessageBox QPushButton:hover {\n"
                              "    background-color: #e0e0e0;\n"
                              "}\
                  "
                              )
            msg.exec()
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.process_thread = ProcessThread(
            self.processor, gamma_file, output_file, self.gradient_config, font_size
        )
        self.process_thread.progress.connect(self.add_status_message)
        self.process_thread.finished.connect(self.on_processing_finished)
        self.process_thread.start()

    def on_processing_finished(self, success):
        self.progress_bar.setVisible(False)

        if success:
            self.add_status_message("PPT处理完成！")
            msg = QMessageBox(self)
            msg.setWindowTitle("成功")
            msg.setText("PPT渐变效果应用完成！")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 200px;
                    min-height: 50px;
                }
            """)
            msg.exec()
        else:
            self.add_status_message("PPT处理失败")
            msg = QMessageBox(self)
            msg.setWindowTitle("失败")
            msg.setText("PPT处理过程中出现错误，请检查文件和配置")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter;
                    padding: 20px;
                    min-width: 300px;
                    min-height: 50px;
                    }
                    """)
            msg.exec()
