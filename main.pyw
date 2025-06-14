import pypinyin
from pypinyin import Style
import sys
from os import listdir, path, startfile, makedirs, system, rename, remove
from shutil import copyfile, move, rmtree
from win32com.client import Dispatch
import json
from PyQt5.QtCore import Qt, QSize, QTimer, QPoint, QMimeData, QUrl
from PyQt5.QtGui import QIcon, QDrag, QPixmap, QPainter
from PyQt5.QtWidgets import (QApplication, QWidget, QListWidget, QFrame, QListWidgetItem, 
                            QMenu, QInputDialog, QLineEdit, QPushButton, QDialog, 
                            QGridLayout, QFileDialog, QLabel, QVBoxLayout, QMessageBox)


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
        else:
            doc.loadList()

        self.QExitButton = QPushButton(self)
        self.QExitButton.setGeometry(560, 20, 20, 20)
        self.QExitButton.setStyleSheet("QPushButton{background-color: darkred;border: none;border-radius: 10px;} QPushButton:hover{background-color: red;}")
        self.QExitButton.clicked.connect(self.close)


    def initGeneralSettings(self):
        self.QRootFolderEdit = QLineEdit(self)
        self.QRootFolderEdit.setGeometry(10, 60, 280, 30)
        self.QRootFolderEdit.setText(self.config['RootFolder'])
        self.QRootFolderEdit.setPlaceholderText("请输入初始目录")

        self.QSaveButton = QPushButton(self)
        self.QSaveButton.setGeometry(295, 60, 100, 30)
        self.QSaveButton.setText("保存")
        self.QSaveButton.clicked.connect(self.save)

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
        
        self.config["RootFolder"] = RootFolder
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
            newPos = event.globalPos() - self.dragPosition
            screenRect = QApplication.desktop().screenGeometry()
            if newPos.x() < 10:
                newPos.setX(10)
            if newPos.y() < 10:
                newPos.setY(10)
            if newPos.x() + self.width() > screenRect.width() - 10:
                newPos.setX(screenRect.width() - self.width() - 10)
            if newPos.y() + self.height() > screenRect.height() - 10:
                newPos.setY(screenRect.height() - self.height() - 10)
            self.move(newPos)


def getAvailableDrives():
    drives = []
    for drive in range(ord('A'), ord('Z')+1):
        driveName = chr(drive) + ":\\"
        if path.exists(driveName):
            drives.append(driveName)
    return drives


# ====================== 列表项类 ====================
class FileListWidgetItem(QListWidgetItem):
    def __init__(self, file_name, file_path, is_setting=False, setting_type=""):
        super().__init__(file_name)
        self.file_path = file_path
        self.is_setting = is_setting
        self.setting_type = setting_type


# ==================== 主程序 ====================
class DocumentOrganizer(QWidget):
    def __init__(self):
        super().__init__()
        self.dragPosition = None
        self.last_char = ""  # 记录上一次输入的字符
        self.current_match_index = 0  # 记录当前匹配项的索引
        self.pinyin_cache = {}  # 缓存拼音转换结果
        self.input_conversion_cache = {}  # 缓存用户输入的转换结果

        try:
            with open('config.json', 'r', encoding='utf-8') as config_file:
                self.config = json.load(config_file)
        except Exception as e:
            self.config = {"RootFolder": ""}
            
        self.DestinationFolder = self.config['RootFolder']
        self.SupportedFileFormats = ['txt', 'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'jpg', 'png', 'bmp', 'gif', 'jpeg', 'zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'iso', 'fba', 'GHO', 'exe', 'chm', 'dll', 'opa', 'cab', 'cat', 'bat', 'py', 'java', 'c', 'cpp', 'php', 'html', 'htm', 'css', 'js', 'json', 'xml', 'yaml', 'yml', 'csv', 'ini', 'log', 'sys', 'bin', 'tmp', 'db', 'mdb', 'md', 'lnk', 'url', 'crdownload', 'prt', 'igs', 'iges', 'jt', 'obj', 'sql', 'key', 'mp4', 'mp3', 'avi', 'ass', 'srt', 'dat', 'dat_old', 'lock', 'cookie', 'mkv', 'mbr']
            
        self.QBackgroundFrame = QFrame(self)
        self.QFileList = QListWidget(self)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("FileIn")
        self.resize(400, 300)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAcceptDrops(True)

        self.QBackgroundFrame.setObjectName('BackgroundFrame')
        self.QBackgroundFrame.setGeometry(0, 0, 400, 300)
        self.QBackgroundFrame.mousePressEvent = self.mousePressEvent
        self.QBackgroundFrame.mouseMoveEvent = self.mouseMoveEvent

        self.QFileList.setGeometry(10, 10, 380, 280)
        self.QFileList.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.QFileList.itemDoubleClicked.connect(self.onDoubleClick)  
        self.QFileList.keyPressEvent = self.keyPressEvent
        self.QFileList.customContextMenuRequested.connect(self.setting)
        self.QFileList.setContextMenuPolicy(Qt.CustomContextMenu)
        self.QFileList.setDragDropMode(QListWidget.InternalMove)
        self.QFileList.setDragDropMode(QListWidget.DragDrop)
        self.QFileList.setDefaultDropAction(Qt.CopyAction)
        self.QFileList.setDragEnabled(True)
        
        self.QFileList.startDrag = self.startDrag
        self.loadList()

        try:
            with open(".qss", "r", encoding="utf-8") as f:
                self.setStyleSheet(f.read())
        except Exception as e:
            pass

    def onDoubleClick(self, item):
        try:
            selection = item.text()
            
            if hasattr(item, 'is_setting') and item.is_setting:
                self.openSetting(item.setting_type)
                return
            if selection == "Back":
                self.navigateUp()
                return
                
            itemPath = path.join(self.DestinationFolder, selection)
                
            if path.isdir(itemPath):
                self.DestinationFolder = path.abspath(itemPath)
                self.loadList()
            else:
                self.safeOpenFile(itemPath)
        except Exception as e:
            self.showError("错误", f"打开失败: {str(e)}")

    def openSetting(self, setting_type):
        dialog = SettingDialog(setting_type)
        dialog.exec_()
        try:
            with open('config.json', 'r', encoding='utf-8') as f:
                self.config = json.load(f)
                self.DestinationFolder = self.config["RootFolder"]
        except Exception as e:
            pass

    def loadList(self, drives=None):
        # 重置匹配状态
        self.last_char = ""
        self.current_match_index = 0
        self.QFileList.clear()
        if drives is not None:
            for drive in drives:
                self.QFileList.addItem(drive)
                self.QFileList.item(self.QFileList.count()-1).setIcon(QIcon("resource/img/Root/hdd.png"))
        else:
            try:
                FileList = listdir(self.DestinationFolder)
                
                if FileList:
                    FileTypeList = [path.splitext(File)[1][1:] for File in FileList]
                else:
                    FileList = ["Back"]
                    FileTypeList = ["folder"]
                
                for fileName, fileType in zip(FileList, FileTypeList):
                    filePath = path.join(self.DestinationFolder, fileName)
                    item = FileListWidgetItem(fileName, filePath)
                    icon_path = f"resource/img/{fileType}.png"
                    if path.exists(icon_path):
                        item.setIcon(QIcon(icon_path))
                    else:
                        item.setIcon(QIcon("resource/img/file.png"))
                    self.QFileList.addItem(item)
            except Exception as e:
                self.showError("错误", f"无法读取目录: {str(e)}")

    def setting(self, pos):
        selectedItems = self.QFileList.selectedItems()
        if not selectedItems:
            return
            
        item = selectedItems[0]
        FileName = item.text()

        if FileName == "Back":
            return
            
        if hasattr(item, 'is_setting') and item.is_setting:
            return
            
        itemPath = path.join(self.DestinationFolder, FileName)
        menu = QMenu(self)
        menu.setWindowOpacity(0.7843137)
        
        backAction = menu.addAction("返回上一目录")
        UpdateAction = menu.addAction("更新")
        OpenFilePathAction = menu.addAction("打开所在位置")
        renameAction = menu.addAction("重命名")
        deleteAction = menu.addAction("删除")
        removeAction = menu.addAction("移出")
        settingAction = menu.addAction("设置")
        exitAction = menu.addAction("退出")
        
        action = menu.exec_(self.QFileList.mapToGlobal(pos))
        
        if action == renameAction:
            newFileName, okFlag = QInputDialog.getText(self, "重命名", "请输入新的名称:")
            if okFlag and newFileName:
                newPath = path.join(self.DestinationFolder, newFileName)
                try:
                    rename(itemPath, newPath)
                    item.setText(newFileName)
                except Exception as e:
                    self.showError("错误", f"重命名失败: {str(e)}")
        elif action == deleteAction:
            answer = QMessageBox.question(self, "确认", "您确定要删除该文件吗？", QMessageBox.Yes | QMessageBox.No)
            if answer == QMessageBox.Yes:
                try:
                    if path.isdir(itemPath):
                        rmtree(itemPath)
                    else:
                        remove(itemPath)
                    self.QFileList.takeItem(self.QFileList.row(item))
                except Exception as e:
                    self.showError("错误", f"删除失败: {str(e)}")
        elif action == removeAction:
            try:
                move(itemPath, "E:\我的文档\Desktop")
                self.loadList()
            except Exception as e:
                self.showError("Move Error:", str(e))
        elif action == settingAction:
            self.QFileList.clear()
            general_setting = FileListWidgetItem("常规设置", "", True, "general")
            general_setting.setIcon(QIcon("resource/img/settings.png"))
            self.QFileList.addItem(general_setting)

            icon_setting = FileListWidgetItem("图标设置", "", True, "icons")
            icon_setting.setIcon(QIcon("resource/img/icons.png"))
            self.QFileList.addItem(icon_setting)

            back_setting = FileListWidgetItem("退出设置", "", True, "back")
            back_setting.setIcon(QIcon("resource/img/icons.png"))
            self.QFileList.addItem(back_setting)
        elif action == exitAction:
            self.close()
            sys.exit()

        elif action == backAction:
            self.navigateUp()

        elif action == UpdateAction:
            if self.QFileList.item(0).text() == "C:\\":
                self.loadList(getAvailableDrives())
            else:
                self.loadList()

    def navigateUp(self):  # 导航到上一级目录
        ParentFolder = path.dirname(self.DestinationFolder)
        if ParentFolder == self.DestinationFolder:
            self.loadList(getAvailableDrives())
        else:
            self.DestinationFolder = ParentFolder
            self.loadList()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.navigateUp()
        elif event.key() == Qt.Key_Down:
            current_row = self.QFileList.currentRow()
            if self.QFileList.count() == 0:
                return
            if current_row == self.QFileList.count() - 1:
                self.QFileList.setCurrentRow(0)
            else:
                self.QFileList.setCurrentRow(current_row + 1)
        elif event.key() == Qt.Key_Up:
            current_row = self.QFileList.currentRow()
            if self.QFileList.count() == 0:
                return
            if current_row == 0:
                self.QFileList.setCurrentRow(self.QFileList.count() - 1)
            else:
                self.QFileList.setCurrentRow(current_row - 1)
        elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            current_item = self.QFileList.currentItem()
            if current_item:
                self.onDoubleClick(current_item)
        elif event.text():  # 支持所有可打印字符
            input_char = event.text()[0]  # 只取第一个字符
            
            # 将用户输入转换为拼音首字母
            converted_char = self.convert_char_to_pinyin(input_char)
            
            # 如果输入了新字符，重置匹配索引
            if converted_char != self.last_char:
                self.last_char = converted_char
                self.current_match_index = 0
            else:
                # 相同字符，增加索引
                self.current_match_index += 1
            
            # 查找匹配项
            self.find_matching_item(converted_char)
        else:
            super().keyPressEvent(event)

    def convert_char_to_pinyin(self, char):
        """将单个字符转换为拼音首字母"""
        if not char:
            return ''
        
        # 检查缓存
        if char in self.input_conversion_cache:
            return self.input_conversion_cache[char]
        
        # 如果是汉字，转换为拼音首字母
        if '\u4e00' <= char <= '\u9fff':
            try:
                # 获取拼音首字母
                pinyin_list = pypinyin.pinyin(char, style=Style.FIRST_LETTER)
                if pinyin_list and pinyin_list[0]:
                    letter = pinyin_list[0][0].lower()
                else:
                    letter = char.lower()
            except:
                letter = char.lower()
        else:
            letter = char.lower()
        
        # 存入缓存
        self.input_conversion_cache[char] = letter
        return letter

    def get_first_letter(self, text):
        """获取文本的首字母（拼音或原字符）"""
        if not text:
            return ''
        
        # 检查缓存
        if text in self.pinyin_cache:
            return self.pinyin_cache[text]
        
        first_char = text[0]
        
        # 如果是汉字，转换为拼音首字母
        if '\u4e00' <= first_char <= '\u9fff':
            try:
                # 获取拼音首字母
                pinyin_list = pypinyin.pinyin(first_char, style=Style.FIRST_LETTER)
                if pinyin_list and pinyin_list[0]:
                    letter = pinyin_list[0][0].lower()
                else:
                    letter = first_char.lower()
            except:
                letter = first_char.lower()
        else:
            letter = first_char.lower()
        
        # 存入缓存
        self.pinyin_cache[text] = letter
        return letter

    def find_matching_item(self, char):
        """查找首字符匹配的项目（支持汉字转拼音）"""
        items = []
        for i in range(self.QFileList.count()):
            item = self.QFileList.item(i)
            item_text = item.text()
            
            # 获取首字母（可能是拼音首字母）
            first_letter = self.get_first_letter(item_text)
            
            if first_letter == char:
                items.append(item)
        
        # 如果没有匹配项，直接返回
        if not items:
            return
        
        # 确保索引在有效范围内
        if self.current_match_index >= len(items):
            self.current_match_index = 0
        
        # 选中匹配项
        item_to_select = items[self.current_match_index]
        self.QFileList.setCurrentItem(item_to_select)
        # 滚动到选中项
        self.QFileList.scrollToItem(item_to_select)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def startDrag(self, supportedActions):
        selectedItems = self.QFileList.selectedItems()
        if not selectedItems:
            return

        drag = QDrag(self)
        mimeData = QMimeData()

        file_paths = [item.file_path for item in selectedItems if hasattr(item, 'file_path') and item.file_path]

        urls = [QUrl.fromLocalFile(path) for path in file_paths]
        mimeData.setUrls(urls)

        modifiers = QApplication.keyboardModifiers()
        if modifiers == Qt.ShiftModifier:
            drag_action = Qt.MoveAction
            icon_name = "cut.png"
        else:
            drag_action = Qt.CopyAction
            icon_name = "copy.png"

        drag.setMimeData(mimeData)

        pixmap = QPixmap(64, 64)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(Qt.white)
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(0, 0, 64, 64)
        painter.drawPixmap(16, 16, QIcon(f"resource/img/{icon_name}").pixmap(32, 32))
        painter.end()

        drag.setPixmap(pixmap)
        drag.setHotSpot(QPoint(32, 32))

        drag.exec_(drag_action)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            files = [url.toLocalFile() for url in event.mimeData().urls()]
            
            if event.dropAction() == Qt.CopyAction:
                self.handleCopiedFiles(files)
            else:
                self.handleDroppedFiles(files)
            event.accept()
        else:
            event.ignore()
        self.loadList()

    def handleCopiedFiles(self, files):
        try:
            for file in files:
                if path.isfile(file):
                    makedirs(self.DestinationFolder, exist_ok=True)
                    dest_path = path.join(self.DestinationFolder, path.basename(file))
                    copyfile(file, dest_path)
                elif path.isdir(file):
                    dest = path.join(self.DestinationFolder, path.basename(file))
                    system(f'xcopy "{file}" "{dest}" /E /I /Y')
            self.loadList()
        except Exception as e:
            self.showError("复制失败", str(e))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            newPos = event.globalPos() - self.dragPosition
            screenRect = QApplication.desktop().screenGeometry()
            if newPos.x() < 10:
                newPos.setX(10)
            if newPos.y() < 10:
                newPos.setY(10)
            if newPos.x() + self.width() > screenRect.width() - 10:
                newPos.setX(screenRect.width() - self.width() - 10)
            if newPos.y() + self.height() > screenRect.height() - 10:
                newPos.setY(screenRect.height() - self.height() - 10)
            self.move(newPos)
    
    def safeOpenFile(self, path):
        try:
            if path.lower().endswith('.lnk'):
                shell = Dispatch('WScript.Shell')
                shortcut = shell.CreateShortCut(path)
                actual_path = shortcut.TargetPath
                startfile(actual_path)
            else:
                startfile(path)
        except Exception as e:
            self.showError("打开失败", f"错误原因：{str(e)}")

    def showError(self, title, message):
        msg = StyledMessageBox(self, title, message, ["确定"])
        msg.exec_()

    def handleDroppedFiles(self, files):
        try:
            for file in files:
                if path.isfile(file):
                    makedirs(self.DestinationFolder, exist_ok=True)
                    move(file, self.DestinationFolder)
                elif path.isdir(file):
                    dest = path.join(self.DestinationFolder, path.basename(file))
                    system(f'xcopy "{file}" "{dest}" /E /I')
                    rmtree(file)
            self.loadList()
        except Exception as e:
            self.showError("移动失败", str(e))


if __name__ == '__main__':
    try:
        app = QApplication([])
        with open(".qss", "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
        doc = DocumentOrganizer()
        doc.show()
        app.exec_()
    except Exception as e:
        sys.exit(1)
