import sys
from os import listdir, path, makedirs
from shutil import copyfile
import json
from PyQt5.QtWidgets import (QApplication, QWidget, QFrame, QInputDialog, QLineEdit, QPushButton, QDialog, QGridLayout, QFileDialog, QLabel, QVBoxLayout, QMessageBox, QDesktopWidget)
from PyQt5.QtCore import Qt, QSize, QTimer
from PyQt5.QtGui import QIcon

class SidebarWidget(QFrame):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.attachbarFrame = QFrame(self)
        self.setStyleSheet("""
            QFrame {
                background: rgba(255, 50, 50, 100);
                border-radius: 8px;
            }
        """)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = event.globalPos() - self.pos()
            event.accept()
    
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and hasattr(self, 'drag_position'):
            new_pos = event.globalPos() - self.drag_position
            
            # 确保不会移出屏幕
            screen_rect = self.screen().availableGeometry()
            new_pos.setX(max(0, min(new_pos.x(), screen_rect.width() - self.width())))
            new_pos.setY(max(0, min(new_pos.y(), screen_rect.height() - self.height())))
            
            self.move(new_pos)
            event.accept()

# ==================== 自定义样式弹窗 ====================
class StyledMessageBox(QDialog):
    def __init__(self, parent=None, title="提示", message="", buttons=[]):
        super().__init__(parent)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.TipWindow = QFrame(self)
        self.setStyleSheet("""
            QFrame {
                background: rgba(50, 50, 50, 220);
                border-radius: 10px;
            }
            QLabel {
                color: white;
                background: transparent;
                font: 18px '黑体';
            }
            QPushButton {
                background: rgb(70, 70, 70);
                color: white;
                font: 14px '黑体';
                border-radius: 5px;
                padding: 8px 20px;
                min-width: 80px;
            }
            QPushButton:hover {
                background: rgb(90, 90, 90);
            }
        """)
        
        layout = QVBoxLayout()
        title_label = QLabel(title)
        title_label.setStyleSheet("font: bold 18px '微软雅黑';")
        layout.addWidget(title_label)
        
        message_label = QLabel(message)
        layout.addWidget(message_label)
        
        btn_layout = QVBoxLayout()
        for btn_text in buttons:
            btn = QPushButton(btn_text)
            btn.clicked.connect(lambda _, t=btn_text: self.done(buttons.index(t)+1))
            btn_layout.addWidget(btn)
        
        layout.addLayout(btn_layout)
        self.setLayout(layout)
        self.TipWindow.setGeometry(0, 0, title_label.width(), 150)
        self.resize(300, 150)


# ==================== 设置对话框 ====================
class SettingDialog(QDialog):
    def __init__(self, setting_type="general"):
        super().__init__()
        self.dragPosition = None
        self.setting_type = setting_type

        try:
            with open('config.json', 'r', encoding='utf-8') as config:
                self.config = json.load(config)
        except Exception as e:
            self.config = {"RootFolder": ""}

        self.QSettingBackgroundFrame = QFrame(self)
        self.initUI()

    def initUI(self):
        self.setWindowTitle(f"设置 - {self.setting_type}")
        self.resize(600, 500)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("""QLineEdit {border: 1px solid rgb(41, 57, 85);border-radius: 3px;background: white;selection-color: green;font-size: 14px;font-family: "黑体";color: black;} QLineEdit:hover {border: 1px solid blue;} QPushButton {background-color: rgb(41, 57, 85);color: white;border-radius: 4px;font-size: 14px;font-family: "黑体";}""")

        self.QSettingBackgroundFrame.setObjectName('BackgroundFrame')
        self.QSettingBackgroundFrame.setGeometry(0, 0, 600, 500)
        self.QSettingBackgroundFrame.mousePressEvent = self.mousePressEvent
        self.QSettingBackgroundFrame.mouseMoveEvent = self.mouseMoveEvent
        self.QSettingBackgroundFrame.setStyleSheet('''#BackgroundFrame{background-color: rgba(0, 0, 0, 200);border-radius: 10px;}''')

        if self.setting_type == "general":
            self.initGeneralSettings()
        elif self.setting_type == "icons":
            self.initIconSettings()


        self.QExitButton = QPushButton(self)
        self.QExitButton.setGeometry(560, 20, 20, 20)
        self.QExitButton.setStyleSheet("QPushButton{background-color: darkred;border: none;border-radius: 10px;} QPushButton:hover{background-color: red;}")
        self.QExitButton.clicked.connect(self.close)


    def initGeneralSettings(self):
        self.QRootFolderEdit = QLineEdit(self)
        self.QRootFolderEdit.setGeometry(10, 60, 280, 30)
        self.QRootFolderEdit.setText(self.config['RootFolder'])
        self.QRootFolderEdit.setPlaceholderText("请输入初始目录")

        self.QWidthEdit = QLineEdit(self)
        self.QWidthEdit.setGeometry(10, 100, 100, 30)
        self.QWidthEdit.setText(str(self.config.get("window_width", 300)))
        self.QWidthEdit.setPlaceholderText("窗口宽度")

        self.QHeightEdit = QLineEdit(self)
        self.QHeightEdit.setGeometry(110, 100, 100, 30)
        self.QHeightEdit.setText(str(self.config.get("window_height", 400)))
        self.QHeightEdit.setPlaceholderText("窗口高度")

        self.QInnerWidthEdit = QLineEdit(self)
        self.QInnerWidthEdit.setGeometry(210, 100, 100, 30)
        self.QInnerWidthEdit.setText(str(self.config.get("window_inner_margin_width", 12)))
        self.QInnerWidthEdit.setPlaceholderText("窗口内边距宽度")

        self.QInnerHeightEdit = QLineEdit(self)
        self.QInnerHeightEdit.setGeometry(310, 100, 100, 30)
        self.QInnerHeightEdit.setText(str(self.config.get("window_inner_margin_height", 12)))
        self.QInnerHeightEdit.setPlaceholderText("窗口内边距高度")

        self.QMarginsEdit = QLineEdit(self)
        self.QMarginsEdit.setGeometry(410, 100, 100, 30)
        self.QMarginsEdit.setText(str(self.config.get("window_screen_margins", 10)))
        self.QMarginsEdit.setPlaceholderText("窗口边距")

        self.QSidebarMarginEdit = QLineEdit(self)
        self.QSidebarMarginEdit.setGeometry(510, 100, 100, 30)
        self.QSidebarMarginEdit.setText(str(self.config.get("sidebar_margin", 20)))
        self.QSidebarMarginEdit.setPlaceholderText("附加栏边距")

        self.QResetButton = QPushButton(self)
        self.QResetButton.setGeometry(510, 60, 100, 30)
        self.QResetButton.setText("重置")
        self.QResetButton.clicked.connect(self.reset)

        self.QSaveButton = QPushButton(self)
        self.QSaveButton.setGeometry(295, 60, 100, 30)
        self.QSaveButton.setText("保存")
        self.QSaveButton.clicked.connect(self.save)

    def reset(self):
        self.QRootFolderEdit.setText(self.config['RootFolder'])
        self.QWidthEdit.setText(str(self.config.get("window_width", 300)))
        self.QHeightEdit.setText(str(self.config.get("window_height", 400)))
        self.QInnerWidthEdit.setText(str(self.config.get("window_inner_margin_width", 12)))
        self.QInnerHeightEdit.setText(str(self.config.get("window_inner_margin_height", 12)))
        self.QMarginsEdit.setText(str(self.config.get("window_screen_margins", 10)))

    def initIconSettings(self):
        self.QManageFileIconFrame = QFrame(self)
        self.QManageFileIconFrame.setGeometry(10, 60, 580, 420)
        self.QManageFileIconFrame.setStyleSheet('''background-color: rgb(41, 57, 85, 100);border-radius: 10px;''')
        self.loadFileIconList()

    def save(self):
        RootFolder = f"{self.QRootFolderEdit.text()}"
        if not path.exists(RootFolder):
            QMessageBox.warning(self, "错误", "根目录不存在！")
            return
        
        self.config.update({
            "RootFolder": RootFolder,
            "window_width": int(self.QWidthEdit.text()),
            "window_height": int(self.QHeightEdit.text()),
            "window_inner_margin_width": int(self.QInnerWidthEdit.text()),
            "window_inner_margin_height": int(self.QInnerHeightEdit.text()),
            "window_screen_margins": int(self.QMarginsEdit.text()),
            "sidebar_margin": int(self.QSidebarMarginEdit.text()),
            "sidebar_width": int(self.QSidebarWidthEdit.text()) if hasattr(self, 'QSidebarWidthEdit') else 200
        })
        
        try:
            with open('config.json', 'w', encoding='utf-8') as configFile:
                json.dump(self.config, configFile)
        except Exception as e:
            self.showMessage("错误", f"保存失败: {str(e)}")

    def loadFileIconList(self):
        try:
            FileIconList = [file for file in listdir("resource/img") if file.endswith('.png')]
        except Exception as e:
            FileIconList = []

        [child.deleteLater() for child in self.QManageFileIconFrame.children()]

        QFileIconListGridLayout = QGridLayout(self.QManageFileIconFrame)
        QFileIconListGridLayout.setContentsMargins(10, 10, 10, 10)
        QFileIconListGridLayout.setHorizontalSpacing(10)
        QFileIconListGridLayout.setVerticalSpacing(10)

        maxButtonsPerRow = 11
        for index, file in enumerate(FileIconList):
            QFileIconButton = QPushButton(self.QManageFileIconFrame)
            QFileIconButton.setIcon(QIcon(f"resource/img/{file}"))
            QFileIconButton.setIconSize(QSize(32, 32))
            QFileIconButton.setFixedSize(44, 44)
            row = index // maxButtonsPerRow
            col = index % maxButtonsPerRow
            QFileIconListGridLayout.addWidget(QFileIconButton, row, col)
            QFileIconButton.clicked.connect(lambda checked, f=file: self.setFileIcon(f))
        
        QAddFileIconButton = QPushButton(self.QManageFileIconFrame)
        QAddFileIconButton.setIcon(QIcon(f"resource/img/Root/MaterialSymbolsAddBoxOutline.png"))
        QAddFileIconButton.setIconSize(QSize(32, 32))
        QAddFileIconButton.setFixedSize(44, 44)
        QFileIconListGridLayout.addWidget(QAddFileIconButton, row, col+1)
        QAddFileIconButton.clicked.connect(self.addNewFileIcon)

        QFileIconListGridLayout.update()

    def setFileIcon(self, file):
        print(f"Selected file icon: {file}")

    def addNewFileIcon(self):
        fileName, ok = QInputDialog.getText(self, "添加文件图标", "扩展名（如txt）:")
        if ok and fileName:
            filePath, _ = QFileDialog.getOpenFileName(
                self, "选择图标文件", "", "图片文件 (*.png *.jpg)")
            if filePath:
                try:
                    ext = path.splitext(filePath)[1]
                    target_path = f"resource/img/{fileName}{ext}"
                    if path.exists(target_path):
                        raise FileExistsError
                    makedirs("resource/img", exist_ok=True)
                    copyfile(filePath, target_path)
                    self.loadFileIconList()
                except Exception as e:
                    self.showMessage("错误", f"添加失败: {str(e)}")

    def showMessage(self, title, message):
        msg = StyledMessageBox(self, title, message, ["确定"])
        msg.exec_()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            delta = event.globalPos() - self.dragPosition
            self.move(delta.x(), delta.y())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    sidebar = SidebarWidget()
    sidebar.show()
    sys.exit(app.exec_())
