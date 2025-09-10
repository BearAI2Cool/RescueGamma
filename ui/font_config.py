import json
import os

from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel,
                               QLineEdit, QPushButton, QMessageBox,
                               QFrame, QScrollArea, QWidget, QCheckBox)


class FontConfigListDialog(QDialog):
    config_deleted = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("字体配置清单")
        self.setMinimumSize(500, 300)
        self.configs = {}
        self.load_configs()
        self.setup_ui()
        self.resize_to_content()

    def load_configs(self):
        try:
            if os.path.exists('font_config.json'):
                with open('font_config.json', 'r', encoding='utf-8') as f:
                    self.configs = json.load(f)
        except Exception as e:
            print(f"加载字体配置失败: {str(e)}")
            self.configs = {}

    def setup_ui(self):
        layout = QVBoxLayout()

        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        if not self.configs:
            no_config_label = QLabel("尚未保存任何字体配置")
            no_config_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            no_config_label.setStyleSheet("border: none; padding: 50px; font-size: 14px;")
            scroll_layout.addWidget(no_config_label)
        else:
            scheme_num = 1
            for config_key, config_data in self.configs.items():
                scheme_frame = QFrame()
                scheme_frame.setFrameStyle(QFrame.Shape.Box)
                scheme_frame.setStyleSheet("QFrame { border: 1px solid #ccc; margin: 5px; padding: 6px; }")
                scheme_layout = QVBoxLayout(scheme_frame)
                scheme_layout.setSpacing(2)

                title_layout = QHBoxLayout()

                scheme_title = QLabel(f"# 方案{scheme_num}：")
                scheme_title.setStyleSheet("font-weight: bold; font-size: 14px; border: none;")
                title_layout.addWidget(scheme_title)

                title_layout.addStretch()

                delete_btn = QPushButton("删除")
                delete_btn.setFixedSize(50, 25)
                delete_btn.setStyleSheet("background-color: #ff4444; color: white; border: none; font-size: 10px;")
                delete_btn.clicked.connect(lambda checked, key=config_key: self.delete_scheme(key))
                title_layout.addWidget(delete_btn)

                scheme_layout.addLayout(title_layout)

                if config_data.get('old_font'):
                    old_font_label = QLabel(f"原字体：{config_data['old_font']}")
                    old_font_label.setStyleSheet(
                        "border: none; margin-left: 10px; margin-top: 1px; margin-bottom: 1px;")
                    scheme_layout.addWidget(old_font_label)

                if config_data.get('old_size'):
                    old_size_label = QLabel(f"原字号：{config_data['old_size']}")
                    old_size_label.setStyleSheet(
                        "border: none; margin-left: 10px; margin-top: 1px; margin-bottom: 1px;")
                    scheme_layout.addWidget(old_size_label)

                if config_data.get('new_font'):
                    new_font_label = QLabel(f"新字体：{config_data['new_font']}")
                    new_font_label.setStyleSheet(
                        "border: none; margin-left: 10px; margin-top: 1px; margin-bottom: 1px;")
                    scheme_layout.addWidget(new_font_label)

                if config_data.get('new_size'):
                    new_size_label = QLabel(f"新字号：{config_data['new_size']}")
                    new_size_label.setStyleSheet(
                        "border: none; margin-left: 10px; margin-top: 1px; margin-bottom: 1px;")
                    scheme_layout.addWidget(new_size_label)
                lang_status = []
                if config_data.get('latin'):
                    lang_status.append("拉丁文")
                if config_data.get('ea'):
                    lang_status.append("象形文")
                if config_data.get('cs'):
                    lang_status.append("复杂文")

                if lang_status:
                    lang_label = QLabel(f"语言类型：{'，'.join(lang_status)}")
                    lang_label.setStyleSheet(
                        "border: none; margin-left: 10px; margin-top: 1px; margin-bottom: 1px;")
                    scheme_layout.addWidget(lang_label)
                else:
                    lang_label = QLabel("语言类型：无")
                    lang_label.setStyleSheet(
                        "border: none; margin-left: 10px; margin-top: 1px; margin-bottom: 1px; color: #999;")
                    scheme_layout.addWidget(lang_label)

                scroll_layout.addWidget(scheme_frame)
                scheme_num += 1

        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)

        close_btn = QPushButton("关闭")
        close_btn.setMinimumHeight(35)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

        self.setLayout(layout)

    def resize_to_content(self):
        content_height = 100

        for config_key, config_data in self.configs.items():
            content_height += 60
            config_items = sum(1 for key in ['old_font', 'old_size', 'new_font', 'new_size', 'latin', 'ea', 'cs'] if
                               config_data.get(key))
            content_height += config_items * 12

        max_height = min(600, max(300, content_height))

        max_width = 500
        for config_key, config_data in self.configs.items():
            for key in ['old_font', 'old_size', 'new_font', 'new_size']:
                if config_data.get(key):
                    text_width = len(f"{key}：{config_data[key]}") * 10 + 100
                    max_width = max(max_width, text_width)

        self.resize(min(800, max_width), max_height)

    def delete_scheme(self, config_key):
        msg = QMessageBox(self)
        msg.setWindowTitle('确认删除')
        msg.setText(f'确定要删除这个字体配置方案吗？')
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
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if msg.exec() == QMessageBox.StandardButton.Yes:
            if config_key in self.configs:
                del self.configs[config_key]

                try:
                    with open('font_config.json', 'w', encoding='utf-8') as f:
                        json.dump(self.configs, f, ensure_ascii=False, indent=2)

                    self.config_deleted.emit()

                    self.accept()

                    success_msg = QMessageBox(self)
                    success_msg.setWindowTitle("成功")
                    success_msg.setText("字体配置方案删除成功")
                    success_msg.setStyleSheet("""
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
                    success_msg.exec()
                except Exception as e:
                    error_msg = QMessageBox(self)
                    error_msg.setWindowTitle("错误")
                    error_msg.setText(f"删除方案失败: {str(e)}")
                    error_msg.setStyleSheet("""
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
                    error_msg.exec()


class FontConfig(QDialog):
    font_changed = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("字体配置")
        self.setFixedSize(400, 250)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)

        old_group_label = QLabel("原字体配置:")
        old_group_label.setStyleSheet("font-weight: bold; border: none;")
        layout.addWidget(old_group_label)

        old_layout = QHBoxLayout()

        old_font_label = QLabel("字体:")
        old_font_label.setStyleSheet("border: none;")
        old_layout.addWidget(old_font_label)

        self.old_font = QLineEdit()
        self.old_font.setPlaceholderText("Arial（可空）")
        old_layout.addWidget(self.old_font)

        old_size_label = QLabel("字号:")
        old_size_label.setStyleSheet("border: none;")
        old_layout.addWidget(old_size_label)

        self.old_size = QLineEdit()
        self.old_size.setPlaceholderText("12（可空）")
        self.old_size.setFixedWidth(85)
        old_layout.addWidget(self.old_size)

        layout.addLayout(old_layout)

        new_group_label = QLabel("新字体配置:")
        new_group_label.setStyleSheet("font-weight: bold; border: none;")
        layout.addWidget(new_group_label)

        new_layout = QHBoxLayout()

        new_font_label = QLabel("字体:")
        new_font_label.setStyleSheet("border: none;")
        new_layout.addWidget(new_font_label)

        self.new_font = QLineEdit()
        self.new_font.setPlaceholderText("微软雅黑（可空）")
        new_layout.addWidget(self.new_font)

        new_size_label = QLabel("字号:")
        new_size_label.setStyleSheet("border: none;")
        new_layout.addWidget(new_size_label)

        self.new_size = QLineEdit()
        self.new_size.setPlaceholderText("14（可空）")
        self.new_size.setFixedWidth(85)
        new_layout.addWidget(self.new_size)

        layout.addLayout(new_layout)

        scope_layout = QHBoxLayout()

        self.latin_check = QCheckBox("拉丁文")
        self.ea_check = QCheckBox("象形文")
        self.cs_check = QCheckBox("复杂文")

        scope_layout.addWidget(self.latin_check)
        scope_layout.addWidget(self.ea_check)
        scope_layout.addWidget(self.cs_check)
        layout.addLayout(scope_layout)

        help_label = QLabel("提示：原配置和新配置至少要有一项不为空")
        help_label.setStyleSheet("color: #666; font-size: 11px; border: none;")
        layout.addWidget(help_label)

        button_layout = QHBoxLayout()

        apply_btn = QPushButton("应用配置")
        apply_btn.clicked.connect(self.apply_config)
        button_layout.addWidget(apply_btn)

        config_list_btn = QPushButton("配置清单")
        config_list_btn.clicked.connect(self.show_config_list)
        button_layout.addWidget(config_list_btn)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def apply_config(self):
        old_font = self.old_font.text().strip()
        old_size = self.old_size.text().strip()
        new_font = self.new_font.text().strip()
        new_size = self.new_size.text().strip()

        if not any([old_font, old_size, new_font, new_size]):
            msg = QMessageBox(self)
            msg.setWindowTitle("警告")
            msg.setText("原配置和新配置至少要有一项不为空")
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QLabel {
                    border: none;
                    qproperty-alignment: 'AlignCenter';
                    padding: 20px;
                    min-width: 250px;
                    min-height: 50px;
                }
            """)
            msg.exec()
            return

        scope_tags = []
        if self.latin_check.isChecked():
            scope_tags.append("latin")
        if self.ea_check.isChecked():
            scope_tags.append("ea")
        if self.cs_check.isChecked():
            scope_tags.append("cs")

        for size_text, size_name in [(old_size, "原字号"), (new_size, "新字号")]:
            if size_text:
                try:
                    float(size_text)
                except ValueError:
                    msg = QMessageBox(self)
                    msg.setWindowTitle("警告")
                    msg.setText(f"{size_name}必须是数字")
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

        config = {
            'old_font': old_font if old_font else None,
            'old_size': old_size if old_size else None,
            'new_font': new_font if new_font else None,
            'new_size': new_size if new_size else None,
            'latin': self.latin_check.isChecked(),
            'ea': self.ea_check.isChecked(),
            'cs': self.cs_check.isChecked()
        }

        config_parts = []
        if old_font:
            config_parts.append(f"字体{old_font}")
        if old_size:
            config_parts.append(f"字号{old_size}")
        if new_font:
            config_parts.append(f"改为{new_font}")
        if new_size:
            config_parts.append(f"改为{new_size}")

        config_key = "_".join(config_parts)

        try:
            font_configs = {}
            if os.path.exists('font_config.json'):
                with open('font_config.json', 'r', encoding='utf-8') as f:
                    font_configs = json.load(f)

            font_configs[config_key] = config

            with open('font_config.json', 'w', encoding='utf-8') as f:
                json.dump(font_configs, f, ensure_ascii=False, indent=2)

            self.font_changed.emit(config)

            msg = QMessageBox(self)
            msg.setWindowTitle("成功")
            msg.setText("字体配置已保存并应用")
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

            self.old_font.clear()
            self.old_size.clear()
            self.new_font.clear()
            self.new_size.clear()

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

    def show_config_list(self):
        dialog = FontConfigListDialog(self)
        dialog.config_deleted.connect(self.on_config_deleted)
        dialog.exec()

    def on_config_deleted(self):
        pass
