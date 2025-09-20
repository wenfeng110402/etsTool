import sys
import os
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget
import qfluentwidgets as qfw


class EtsToolInstaller(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        qfw.setTheme(qfw.Theme.DARK)
        self.setWindowTitle('E听说外挂安装器')
        self.setGeometry(100, 100, 500, 350)

        # Central Widget and Layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Install Button
        install_button = qfw.PrimaryPushButton('安装', self)
        install_button.clicked.connect(self.install)
        layout.addWidget(install_button)

        # Uninstall Button
        uninstall_button = qfw.PrimaryPushButton('卸载', self)
        uninstall_button.clicked.connect(self.uninstall)
        layout.addWidget(uninstall_button)

    def install(self):
        resource_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Resource')
        target_dir = 'C:\TargetDirectory'  # Replace with your target directory
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

        for item in os.listdir(resource_dir):
            source_item = os.path.join(resource_dir, item)
            target_item = os.path.join(target_dir, item)
            if os.path.isdir(source_item):
                shutil.copytree(source_item, target_item, dirs_exist_ok=True)
            else:
                shutil.copy2(source_item, target_item)
        print("安装完成")

    def uninstall(self):
        target_dir = 'C:\TargetDirectory'  # Replace with your target directory
        if os.path.exists(target_dir):
            shutil.rmtree(target_dir)
        print("卸载完成")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EtsToolInstaller()
    ex.show()
    sys.exit(app.exec_())