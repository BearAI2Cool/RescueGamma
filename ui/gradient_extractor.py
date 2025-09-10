import json

from PySide6.QtCore import Qt, Signal, QPoint, QRect
from PySide6.QtGui import QPixmap, QGuiApplication, QFontMetrics, QPainter, QPen, QCursor, QColor, QPainterPath
from PySide6.QtWidgets import (QApplication, QAbstractSpinBox)
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGroupBox, QLabel, QPushButton, QTextEdit, QSpinBox,
                               QFileDialog, QMessageBox, QCheckBox)


class ClickableImageLabel(QLabel):
    color_picked = Signal(QColor, QPoint)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent
        self.picking_mode = False
        self.mouse_pos = None
        self.background_color = QColor(255, 255, 255)
        self.setMouseTracking(True)

    def set_picking_mode(self, enabled):
        self.picking_mode = enabled
        if enabled:
            self.setCursor(QCursor(Qt.CrossCursor))
        else:
            self.setCursor(QCursor(Qt.ArrowCursor))

    def calculate_background_color(self, pixmap):
        if not pixmap or pixmap.isNull():
            return QColor(255, 255, 255)

        image = pixmap.toImage()
        width = image.width()
        height = image.height()

        colors = []
        corners = [
            (0, 0),
            (width - 1, 0),
            (0, height - 1),
            (width - 1, height - 1)
        ]

        for x, y in corners:
            color = image.pixelColor(x, y)
            colors.append((color.red(), color.green(), color.blue()))

        avg_r = sum(c[0] for c in colors) // len(colors)
        avg_g = sum(c[1] for c in colors) // len(colors)
        avg_b = sum(c[2] for c in colors) // len(colors)

        return QColor(avg_r, avg_g, avg_b)

    def get_color_at_position(self, img_x, img_y, pixmap):
        if not pixmap or pixmap.isNull():
            return self.background_color

        image = pixmap.toImage()
        width = image.width()
        height = image.height()

        if 0 <= img_x < width and 0 <= img_y < height:
            return image.pixelColor(img_x, img_y)
        else:
            return self.background_color

    def mouseMoveEvent(self, event):
        self.mouse_pos = event.pos()
        self.update()
        super().mouseMoveEvent(event)

    def mousePressEvent(self, event):
        if self.picking_mode and event.button() == Qt.LeftButton:
            click_pos = event.pos()

            if self.pixmap() and not self.pixmap().isNull():
                pixmap = self.pixmap()
                label_size = self.size()

                pixmap_size = pixmap.size()

                x_offset = (label_size.width() - pixmap_size.width()) // 2
                y_offset = (label_size.height() - pixmap_size.height()) // 2

                if (x_offset <= click_pos.x() <= x_offset + pixmap_size.width() and
                        y_offset <= click_pos.y() <= y_offset + pixmap_size.height()):
                    img_x = int(click_pos.x() - x_offset)
                    img_y = int(click_pos.y() - y_offset)

                    img_x = max(0, min(img_x, pixmap_size.width() - 1))
                    img_y = max(0, min(img_y, pixmap_size.height() - 1))

                    image = pixmap.toImage()
                    color = image.pixelColor(img_x, img_y)

                    self.color_picked.emit(color, QPoint(img_x, img_y))

        super().mousePressEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)

        if self.parent_widget and hasattr(self.parent_widget, 'preview_lines'):
            painter = QPainter(self)

            pen = QPen()
            pen.setStyle(Qt.DashLine)
            pen.setWidth(2)

            if self.parent_widget.is_dark_mode:
                pen.setColor(QColor(255, 255, 0, 200))
            else:
                pen.setColor(QColor(255, 0, 0, 200))

            painter.setPen(pen)

            if self.pixmap() and not self.pixmap().isNull():
                pixmap = self.pixmap()
                label_size = self.size()

                pixmap_size = pixmap.size()

                x_offset = (label_size.width() - pixmap_size.width()) // 2
                y_offset = (label_size.height() - pixmap_size.height()) // 2

                for line_pos in self.parent_widget.preview_lines:
                    relative_pos = line_pos / 100000.0
                    x = x_offset + relative_pos * pixmap_size.width()

                    painter.drawLine(x, y_offset - 15, x, y_offset + pixmap_size.height() + 15)

                painter.end()

        if self.mouse_pos and self.pixmap() and not self.pixmap().isNull():
            painter = QPainter(self)
            painter.setRenderHint(QPainter.Antialiasing)

            pixmap = self.pixmap()
            label_size = self.size()
            pixmap_size = pixmap.size()
            ps_w = pixmap_size.width()
            ps_h = pixmap_size.height()

            x_offset = (label_size.width() - ps_w) // 2
            y_offset = (label_size.height() - ps_h) // 2

            mouse_x, mouse_y = self.mouse_pos.x(), self.mouse_pos.y()
            if (x_offset <= mouse_x <= x_offset + ps_w and
                    y_offset <= mouse_y <= y_offset + ps_h):

                img_x = mouse_x - x_offset
                img_y = mouse_y - y_offset

                current_color = self.get_color_at_position(img_x, img_y, pixmap)

                magnifier_radius = 40
                magnifier_center = QPoint(mouse_x, mouse_y)
                magnifier_rect = QRect(
                    magnifier_center.x() - magnifier_radius,
                    magnifier_center.y() - magnifier_radius,
                    magnifier_radius * 2,
                    magnifier_radius * 2
                )

                path = QPainterPath()
                path.addEllipse(magnifier_rect)
                painter.setClipPath(path)

                zoom_factor = 2
                zoom_size = magnifier_radius // zoom_factor

                magnifier_pixmap = QPixmap(magnifier_radius * 2, magnifier_radius * 2)
                magnifier_pixmap.fill(self.background_color)
                magnifier_painter = QPainter(magnifier_pixmap)

                src_left = img_x - zoom_size
                src_top = img_y - zoom_size
                src_right = img_x + zoom_size
                src_bottom = img_y + zoom_size

                for y in range(magnifier_radius * 2):
                    for x in range(magnifier_radius * 2):
                        orig_x = src_left + (x * 2 * zoom_size) // (magnifier_radius * 2)
                        orig_y = src_top + (y * 2 * zoom_size) // (magnifier_radius * 2)

                        color = self.get_color_at_position(orig_x, orig_y, pixmap)
                        magnifier_painter.setPen(QPen(color))
                        magnifier_painter.drawPoint(x, y)

                magnifier_painter.end()

                painter.drawPixmap(magnifier_rect, magnifier_pixmap)

                painter.setClipping(False)
                painter.setPen(QPen(Qt.black, 2))
                painter.drawEllipse(magnifier_rect)

                center_x = magnifier_center.x()
                center_y = magnifier_center.y()
                painter.setPen(QPen(Qt.red, 1))
                painter.drawLine(center_x - 5, center_y, center_x + 5, center_y)
                painter.drawLine(center_x, center_y - 5, center_x, center_y + 5)

                color_indicator_size = 20
                color_indicator_rect = QRect(
                    magnifier_center.x() - color_indicator_size // 2,
                    magnifier_center.y() + magnifier_radius + 5,
                    color_indicator_size,
                    color_indicator_size
                )
                painter.fillRect(color_indicator_rect, current_color)
                painter.setPen(QPen(Qt.black, 1))
                painter.drawRect(color_indicator_rect)

                painter.end()


class GradientExtractor(QWidget):

    def __init__(self, parent=None, is_dark_mode=False):
        super().__init__(parent)
        self.current_image = None
        self.is_dark_mode = is_dark_mode
        self.preview_lines = []
        self.color_picking_active = False
        self.current_pick_index = 0
        self.picked_colors = []
        self.active_positions = []
        self.init_ui()
        self.apply_theme(is_dark_mode)

    def init_ui(self):
        self.setWindowTitle("渐变提取")
        self.setMinimumSize(800, 750)

        layout = QVBoxLayout()

        input_group = QGroupBox("图片输入")
        font = input_group.font()
        font.setBold(True)
        input_group.setFont(font)
        input_layout = QHBoxLayout()
        input_layout.setSpacing(20)
        input_layout.setContentsMargins(15, 20, 15, 15)

        self.paste_btn = QPushButton("从剪贴板粘贴")
        self.upload_btn = QPushButton("上传图片")
        self.paste_btn.setMinimumHeight(30)
        self.upload_btn.setMinimumHeight(30)
        self.paste_btn.clicked.connect(self.paste_from_clipboard)
        self.upload_btn.clicked.connect(self.upload_image)

        input_layout.addWidget(self.paste_btn)
        input_layout.addWidget(self.upload_btn)
        input_group.setLayout(input_layout)

        preview_group = QGroupBox("预览")
        font = preview_group.font()
        font.setBold(True)
        preview_group.setFont(font)

        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setContentsMargins(15, 20, 15, 20)
        preview_layout.setSpacing(12)

        self.image_label = ClickableImageLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(200, 150)
        self.image_label.color_picked.connect(self.on_color_picked)
        preview_layout.addWidget(self.image_label)

        coord_group = QGroupBox("位置参数")
        font = coord_group.font()
        font.setBold(True)
        coord_group.setFont(font)

        coord_layout = QVBoxLayout(coord_group)
        coord_layout.setSpacing(12)
        coord_layout.setContentsMargins(15, 20, 15, 15)

        positions_layout = QHBoxLayout()
        positions_layout.setSpacing(12)

        self.coord_inputs = []
        self.position_checkboxes = []
        self.color_displays = []
        coord_labels = ["位置1", "位置2", "位置3", "位置4"]
        default_values = [0, 74000, 83000, 100000]

        for i, (label, default) in enumerate(zip(coord_labels, default_values)):
            item_layout = QVBoxLayout()
            item_layout.setAlignment(Qt.AlignCenter)
            item_layout.setSpacing(5)

            label_widget = QLabel(label)
            label_widget.setStyleSheet("border: none;")
            item_layout.addWidget(label_widget, alignment=Qt.AlignCenter)

            control_layout = QHBoxLayout()
            control_layout.setAlignment(Qt.AlignCenter)
            control_layout.setSpacing(5)

            if i == 0 or i == 3:
                spinbox = QSpinBox()
                spinbox.setRange(0, 100000)
                spinbox.setValue(default)
                spinbox.setSingleStep(1000)
                spinbox.setFixedWidth(100)
                spinbox.setEnabled(False)
                spinbox.setButtonSymbols(QAbstractSpinBox.NoButtons)
                checkbox = None
                control_layout.addWidget(spinbox)
            else:
                checkbox = QCheckBox()
                checkbox.setChecked(False)
                checkbox.setFixedSize(20, 20)

                spinbox = QSpinBox()
                spinbox.setRange(0, 100000)
                spinbox.setSingleStep(1000)
                spinbox.setFixedWidth(100)
                spinbox.setValue(0)
                spinbox.setEnabled(False)
                spinbox.setButtonSymbols(QAbstractSpinBox.NoButtons)
                spinbox.setSpecialValueText("空")

                def make_checkbox_handler(spinbox_ref, default_val):
                    def on_checkbox_changed(checked):
                        if checked:
                            spinbox_ref.setEnabled(True)
                            spinbox_ref.setButtonSymbols(QAbstractSpinBox.UpDownArrows)
                            spinbox_ref.setSpecialValueText("")
                            spinbox_ref.setValue(default_val)
                        else:
                            spinbox_ref.setEnabled(False)
                            spinbox_ref.setButtonSymbols(QAbstractSpinBox.NoButtons)
                            spinbox_ref.setSpecialValueText("空")
                            spinbox_ref.setValue(0)
                        self.update_preview_lines()

                    return on_checkbox_changed

                checkbox.toggled.connect(make_checkbox_handler(spinbox, default))

                spinbox.valueChanged.connect(self.update_preview_lines)

                control_layout.addWidget(checkbox)
                control_layout.addWidget(spinbox)

            self.coord_inputs.append(spinbox)
            self.position_checkboxes.append(checkbox)
            item_layout.addLayout(control_layout)

            color_display = QLabel("颜色: 未设置")
            color_display.setStyleSheet("border: 1px solid #ccc; padding: 5px; margin-top: 5px;")
            color_display.setAlignment(Qt.AlignCenter)
            color_display.setMinimumHeight(30)
            self.color_displays.append(color_display)
            item_layout.addWidget(color_display)

            positions_layout.addLayout(item_layout)

        self.coord_inputs[0].valueChanged.connect(self.update_preview_lines)
        self.coord_inputs[3].valueChanged.connect(self.update_preview_lines)

        coord_layout.addLayout(positions_layout)

        self.extract_btn = QPushButton("提取渐变颜色")
        self.extract_btn.clicked.connect(self.start_color_picking)
        self.extract_btn.setEnabled(False)

        font = self.extract_btn.font()
        font.setBold(True)
        self.extract_btn.setFont(font)

        fm = QFontMetrics(font)
        text_width = fm.horizontalAdvance(self.extract_btn.text())
        self.extract_btn.setFixedWidth(text_width + 20)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.extract_btn, alignment=Qt.AlignHCenter)
        coord_layout.addLayout(btn_layout)

        coord_group.setLayout(coord_layout)

        result_group = QGroupBox("JSON结果")
        font = result_group.font()
        font.setBold(True)
        result_group.setFont(font)

        result_layout = QVBoxLayout(result_group)
        result_layout.setContentsMargins(15, 20, 15, 15)
        result_layout.setSpacing(10)

        self.result_text = QTextEdit()
        self.result_text.setMaximumHeight(150)

        self.copy_btn = QPushButton("复制JSON")
        self.copy_btn.clicked.connect(self.copy_json)

        result_layout.addWidget(self.result_text)
        result_layout.addWidget(self.copy_btn, alignment=Qt.AlignHCenter)
        result_group.setLayout(result_layout)

        layout.addWidget(input_group)
        layout.addWidget(preview_group)
        layout.addWidget(coord_group)
        layout.addWidget(result_group)

        self.setLayout(layout)

    def center_window(self):
        qr = self.frameGeometry()
        screen_geometry = QGuiApplication.primaryScreen().availableGeometry()
        cp = screen_geometry.center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def apply_theme(self, is_dark_mode):
        self.is_dark_mode = is_dark_mode
        if is_dark_mode:
            self.setStyleSheet("""
                QWidget { background-color: #333; color: #fff; }
                QGroupBox { 
                    border: 1px solid #555; 
                    font-weight: bold;
                    padding-top: 10px;
                    margin-top: 5px;
                }
                QGroupBox::title {
                    color: #fff;
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 5px 0 5px;
                }
                QLabel { color: #fff; }
                QPushButton { 
                    background-color: #555; 
                    color: white; 
                    border: 1px solid #777;
                    padding: 5px;
                    border-radius: 3px;
                }
                QPushButton:hover { background-color: #666; }
                QPushButton:disabled { 
                    background-color: #444; 
                    color: #888;
                }
                QTextEdit, QSpinBox { 
                    background-color: #444; 
                    color: white; 
                    border: 1px solid #666;
                }
                QSpinBox:disabled {
                    background-color: #333;
                    color: #888;
                }
                QCheckBox {
                    color: #fff;
                    spacing: 5px;
                }
                QCheckBox::indicator {
                    width: 16px;
                    height: 16px;
                    border-radius: 3px;
                }
                QCheckBox::indicator:unchecked {
                    background-color: #444;
                    border: 2px solid #666;
                }
                QCheckBox::indicator:checked {
                    background-color: #0078d4;
                    border: 2px solid #0078d4;
                    image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMTAiIHZpZXdCb3g9IjAgMCAxMCAxMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTggM0w0IDdMMiA1IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjEuNSIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIi8+Cjwvc3ZnPgo=);
                }
                QCheckBox::indicator:hover {
                    border: 2px solid #0078d4;
                }
            """)
            if self.image_label.pixmap() is None or self.image_label.pixmap().isNull():
                self.image_label.setStyleSheet(
                    "border: 2px dashed #666; padding: 20px; color: #ccc; background-color: #444;")
        else:
            self.setStyleSheet("""
                QCheckBox::indicator {
                    width: 16px;
                    height: 16px;
                    border-radius: 3px;
                }
                QCheckBox::indicator:unchecked {
                    background-color: white;
                    border: 2px solid #999;
                }
                QCheckBox::indicator:checked {
                    background-color: #0078d4;
                    border: 2px solid #0078d4;
                    image: url(data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAiIGhlaWdodD0iMTAiIHZpZXdCb3g9IjAgMCAxMCAxMCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTggM0w0IDdMMiA1IiBzdHJva2U9IndoaXRlIiBzdHJva2Utd2lkdGg9IjEuNSIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIi8+Cjwvc3ZnPgo=);
                }
                QCheckBox::indicator:hover {
                    border: 2px solid #0078d4;
                }
            """)
            if self.image_label.pixmap() is None or self.image_label.pixmap().isNull():
                self.image_label.setStyleSheet("border: 2px dashed #ccc; padding: 20px;")

    def paste_from_clipboard(self):
        clipboard = QApplication.clipboard()
        pixmap = clipboard.pixmap()

        if not pixmap.isNull():
            self.current_image = pixmap
            self.display_image(pixmap)
            self.extract_btn.setEnabled(True)
        else:
            msg_box = QMessageBox(QMessageBox.Warning, "警告", "剪贴板中没有图片", parent=self)
            msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
            msg_box.exec_()

    def upload_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择图片", "", "图片文件 (*.png *.jpg *.jpeg *.bmp)"
        )

        if file_path:
            pixmap = QPixmap(file_path)
            if not pixmap.isNull():
                self.current_image = pixmap
                self.display_image(pixmap)
                self.extract_btn.setEnabled(True)
            else:
                msg_box = QMessageBox(QMessageBox.Critical, "错误", "无法加载图片", parent=self)
                msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
                msg_box.exec_()

    def display_image(self, pixmap):

        self.image_label.background_color = self.image_label.calculate_background_color(pixmap)

        label_size = self.image_label.size()
        max_width = max(label_size.width(), 200)
        max_height = max(label_size.height(), 150)

        scale_x = max_width / pixmap.width()
        scale_y = max_height / pixmap.height()
        scale = min(scale_x, scale_y, 1.0)

        if scale < 1.0:
            new_width = int(pixmap.width() * scale)
            new_height = int(pixmap.height() * scale)
            scaled_pixmap = pixmap.scaled(new_width, new_height, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        else:
            scaled_pixmap = pixmap

        self.image_label.setPixmap(scaled_pixmap)
        self.image_label.setScaledContents(False)
        self.image_label.setText("")
        self.image_label.setStyleSheet("border: none; padding: 5px;")

        self.update_preview_lines()

    def update_preview_lines(self):
        if not self.current_image:
            return

        self.preview_lines = []

        for i, spinbox in enumerate(self.coord_inputs):
            if i == 0 or i == 3:
                self.preview_lines.append(spinbox.value())
            else:
                checkbox = self.position_checkboxes[i]
                if checkbox and checkbox.isChecked():
                    self.preview_lines.append(spinbox.value())

        self.image_label.update()

    def start_color_picking(self):
        if not self.current_image:
            msg_box = QMessageBox(QMessageBox.Warning, "警告", "请先上传图片", parent=self)
            msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
            msg_box.exec_()
            return

        self.active_positions = []
        for i, spinbox in enumerate(self.coord_inputs):
            if i == 0 or i == 3:
                self.active_positions.append((i, spinbox.value()))
            else:
                checkbox = self.position_checkboxes[i]
                if checkbox and checkbox.isChecked():
                    self.active_positions.append((i, spinbox.value()))

        if not self.active_positions:
            msg_box = QMessageBox(QMessageBox.Warning, "警告", "请至少可用一个位置", parent=self)
            msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
            msg_box.exec_()
            return

        self.active_positions.sort(key=lambda x: x[1])

        self.color_picking_active = True
        self.current_pick_index = 0
        self.picked_colors = []
        self.image_label.set_picking_mode(True)

        self.extract_btn.setText(
            f"请点击位置 {self.active_positions[0][1]} ({self.current_pick_index + 1}/{len(self.active_positions)})")
        self.extract_btn.setEnabled(False)

    def on_color_picked(self, color):
        if not self.color_picking_active:
            return

        pos_index, pos_value = self.active_positions[self.current_pick_index]
        hex_color = f"#{color.red():02X}{color.green():02X}{color.blue():02X}"
        self.picked_colors.append((pos_index, pos_value, hex_color, color))

        color_display = self.color_displays[pos_index]
        color_display.setText(f"颜色: {hex_color}")
        color_display.setStyleSheet(f"""
            border: 2px solid {hex_color}; 
            padding: 5px; 
            margin-top: 5px; 
            background-color: {hex_color};
            color: {'white' if color.red() + color.green() + color.blue() < 384 else 'black'};
        """)

        self.current_pick_index += 1

        if self.current_pick_index >= len(self.active_positions):
            self.finish_color_picking()
        else:
            next_pos_value = self.active_positions[self.current_pick_index][1]
            self.extract_btn.setText(
                f"请点击位置 {next_pos_value} ({self.current_pick_index + 1}/{len(self.active_positions)})")

    def finish_color_picking(self):
        self.color_picking_active = False
        self.image_label.set_picking_mode(False)
        self.extract_btn.setText("提取渐变颜色")
        self.extract_btn.setEnabled(True)
        self.generate_json_result()

    def generate_json_result(self):
        try:
            gradient_config = []

            sorted_colors = sorted(self.picked_colors, key=lambda x: x[1])

            for pos_index, pos_value, hex_color, qt_color in sorted_colors:
                gradient_config.append({
                    "position": pos_value,
                    "color": hex_color
                })

            font_size = "14px"
            font_name = "微软雅黑"

            result = {
                font_size: {
                    "gradient_config": gradient_config,
                    "font_name": font_name
                }
            }

            json_str = json.dumps(result, indent=2, ensure_ascii=False)
            self.result_text.setText(json_str)

        except Exception as e:
            error_msg = f"生成JSON结果失败: {str(e)}"
            msg_box = QMessageBox(QMessageBox.Critical, "错误", error_msg, parent=self)
            msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
            msg_box.exec_()

    def copy_json(self):
        text = self.result_text.toPlainText()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            msg_box = QMessageBox(QMessageBox.Information, "成功", "JSON已复制到剪贴板", parent=self)
            msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
            msg_box.exec_()
        else:
            msg_box = QMessageBox(QMessageBox.Warning, "提示", "没有可复制的JSON结果", parent=self)
            msg_box.setStyleSheet("QMessageBox { border: none; } QMessageBox QLabel { border: none; }")
            msg_box.exec_()

    def reset_colors(self):
        for color_display in self.color_displays:
            color_display.setText("颜色: 未设置")
            color_display.setStyleSheet("border: 1px solid #ccc; padding: 5px; margin-top: 5px;")
        self.picked_colors = []
        self.result_text.clear()


def add_gradient_menu_to_app(parent_widget, menubar, is_dark_mode=False):

    def show_gradient_extractor():
        gradient_window = GradientExtractor(parent_widget, is_dark_mode)
        gradient_window.setParent(parent_widget)
        gradient_window.setWindowModality(Qt.ApplicationModal)
        gradient_window.center_window()
        gradient_window.show()
        return gradient_window

    tools_menu = None
    for action in menubar.actions():
        if action.text() == "工具":
            tools_menu = action.menu()
            break

    if not tools_menu:
        tools_menu = menubar.addMenu("工具")

    gradient_action = tools_menu.addAction("渐变提取")
    gradient_action.triggered.connect(show_gradient_extractor)

    return gradient_action
