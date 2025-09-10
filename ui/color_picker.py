import colorsys

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QPixmap, QPainter, QColor
from PySide6.QtWidgets import QApplication
from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel


class ColorPicker(QDialog):
    color_selected = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("颜色查询")
        self.setFixedSize(400, 420)
        self.current_color = "#9A6FDC"
        self.selected_color = "#9A6FDC"
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        self.color_display = QLabel()
        self.color_display.setFixedSize(380, 50)
        self.color_display.setStyleSheet(f"background-color: {self.current_color}; border: 1px solid black;")
        layout.addWidget(self.color_display)

        self.color_palette = ColorPalette()
        self.color_palette.color_changed.connect(self.on_color_hover)
        self.color_palette.color_clicked.connect(self.on_color_clicked)
        layout.addWidget(self.color_palette)

        layout.addSpacing(10)

        self.status_label = QLabel(f"已选择 : {self.selected_color}")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("border: none; font-weight: bold;")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def on_color_hover(self, color_hex):
        self.current_color = color_hex
        self.color_display.setStyleSheet(f"background-color: {color_hex}; border: 1px solid black;")

    def on_color_clicked(self, color_hex):
        self.selected_color = color_hex
        self.current_color = color_hex
        self.color_display.setStyleSheet(f"background-color: {color_hex}; border: 1px solid black;")

        self.status_label.setText(f"已选择: {color_hex}")

        self.copy_selected_color()
        self.color_selected.emit(color_hex)

    def copy_selected_color(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.selected_color)

        self.status_label.setText(f"已复制: {self.selected_color}")

        from PySide6.QtCore import QTimer
        QTimer.singleShot(1500, lambda: self.status_label.setText(f"已选择: {self.selected_color}"))


class ColorPalette(QLabel):
    color_changed = Signal(str)
    color_clicked = Signal(str)

    def __init__(self):
        super().__init__()
        self.setFixedSize(380, 300)
        self.setMouseTracking(True)
        self.create_palette()

    def create_palette(self):
        pixmap = QPixmap(380, 300)
        painter = QPainter(pixmap)

        for x in range(380):
            for y in range(300):
                h = x / 380.0
                s = y / 150.0 if y < 150 else 1.0
                v = 1.0 if y < 150 else 1.0 - ((y - 150) / 150.0)

                r, g, b = colorsys.hsv_to_rgb(h, s, v)
                color = QColor(int(r * 255), int(g * 255), int(b * 255))
                painter.setPen(color)
                painter.drawPoint(x, y)

        painter.end()
        self.setPixmap(pixmap)

    def mouseMoveEvent(self, event):
        x, y = event.position().x(), event.position().y()
        if 0 <= x < 380 and 0 <= y < 300:
            color = self.get_color_at_position(int(x), int(y))
            self.color_changed.emit(color)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            x, y = event.position().x(), event.position().y()
            if 0 <= x < 380 and 0 <= y < 300:
                color = self.get_color_at_position(int(x), int(y))
                self.color_clicked.emit(color)

    def get_color_at_position(self, x, y):
        h = x / 380.0
        s = y / 150.0 if y < 150 else 1.0
        v = 1.0 if y < 150 else 1.0 - ((y - 150) / 150.0)

        r, g, b = colorsys.hsv_to_rgb(h, s, v)
        return f"#{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"
