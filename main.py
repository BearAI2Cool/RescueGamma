import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from ui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)

    # 设置应用程序信息
    app.setApplicationName("Moss_Go PPT工具")
    app.setApplicationVersion("3.0")
    app.setOrganizationName("Moss_Go")

    # 设置应用图标
    icon_path = "ico/RescueGamma_32.ico"
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # 创建主窗口
    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
