import sys
import json
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QDialog, QFrame, QLineEdit, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QMessageBox
from PyQt5.QtCore import Qt

class SettingApp(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FileIn 设置")
        self.setFixedSize(400, 320)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet(self.load_qss())
        self.config = self.load_config()
        self.initUI()

    def load_qss(self):
        # 可维护的统一样式
        return '''
        QDialog { background: #23272e; border-radius: 12px; }
        QFrame { background: #2c313a; border-radius: 8px; }
        QLabel { color: #fff; font-size: 15px; }
        QLineEdit { background: #23272e; color: #fff; border: 1px solid #444; border-radius: 4px; padding: 4px; font-size: 15px; }
        QPushButton { background: #3a8ee6; color: #fff; border-radius: 4px; padding: 6px 18px; font-size: 15px; }
        QPushButton:hover { background: #5bb1ff; }
        '''

    def load_config(self):
        default = {
            "RootFolder": str(Path.home()),
            "window": {
                "width": 400,
                "height": 400,
                "inner_margin_width": 12,
                "inner_margin_height": 12,
                "screen_margins": 10
            },
            "window_position": {"x": 100, "y": 100},
            "sidebar_relative_x": 420,
            "sidebar_relative_y": 0
        }
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                user = json.load(f)
            for k in default:
                if k not in user:
                    user[k] = default[k]
            return user
        except Exception:
            return default

    def initUI(self):
        frame = QFrame(self)
        frame.setGeometry(20, 20, 360, 280)
        layout = QVBoxLayout(frame)
        # 根目录
        hlayout1 = QHBoxLayout()
        hlayout1.addWidget(QLabel("根目录:"))
        self.rootEdit = QLineEdit(self.config["RootFolder"])
        hlayout1.addWidget(self.rootEdit)
        layout.addLayout(hlayout1)
        # 窗口宽高
        hlayout2 = QHBoxLayout()
        hlayout2.addWidget(QLabel("宽度:"))
        self.widthEdit = QLineEdit(str(self.config["window"]["width"]))
        hlayout2.addWidget(self.widthEdit)
        hlayout2.addWidget(QLabel("高度:"))
        self.heightEdit = QLineEdit(str(self.config["window"]["height"]))
        hlayout2.addWidget(self.heightEdit)
        layout.addLayout(hlayout2)
        # 内边距
        hlayout3 = QHBoxLayout()
        hlayout3.addWidget(QLabel("内边距宽:"))
        self.inWEdit = QLineEdit(str(self.config["window"]["inner_margin_width"]))
        hlayout3.addWidget(self.inWEdit)
        hlayout3.addWidget(QLabel("内边距高:"))
        self.inHEdit = QLineEdit(str(self.config["window"]["inner_margin_height"]))
        hlayout3.addWidget(self.inHEdit)
        layout.addLayout(hlayout3)
        # 屏幕边距
        hlayout4 = QHBoxLayout()
        hlayout4.addWidget(QLabel("屏幕边距:"))
        self.marginEdit = QLineEdit(str(self.config["window"]["screen_margins"]))
        hlayout4.addWidget(self.marginEdit)
        layout.addLayout(hlayout4)
        # 侧边栏相对位置
        hlayout5 = QHBoxLayout()
        hlayout5.addWidget(QLabel("侧栏相对X:"))
        self.sidebarXEdit = QLineEdit(str(self.config.get("sidebar_relative_x", 420)))
        hlayout5.addWidget(self.sidebarXEdit)
        hlayout5.addWidget(QLabel("Y:"))
        self.sidebarYEdit = QLineEdit(str(self.config.get("sidebar_relative_y", 0)))
        hlayout5.addWidget(self.sidebarYEdit)
        layout.addLayout(hlayout5)
        # 按钮
        btnLayout = QHBoxLayout()
        saveBtn = QPushButton("保存")
        saveBtn.clicked.connect(self.save)
        btnLayout.addWidget(saveBtn)
        exitBtn = QPushButton("退出")
        exitBtn.clicked.connect(self.close)
        btnLayout.addWidget(exitBtn)
        layout.addLayout(btnLayout)

    def save(self):
        try:
            self.config["RootFolder"] = self.rootEdit.text()
            self.config["window"]["width"] = int(self.widthEdit.text())
            self.config["window"]["height"] = int(self.heightEdit.text())
            self.config["window"]["inner_margin_width"] = int(self.inWEdit.text())
            self.config["window"]["inner_margin_height"] = int(self.inHEdit.text())
            self.config["window"]["screen_margins"] = int(self.marginEdit.text())
            self.config["sidebar_relative_x"] = int(self.sidebarXEdit.text())
            self.config["sidebar_relative_y"] = int(self.sidebarYEdit.text())
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            QMessageBox.information(self, "成功", "配置已保存！")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存失败: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = SettingApp()
    win.show()
    sys.exit(app.exec_())
