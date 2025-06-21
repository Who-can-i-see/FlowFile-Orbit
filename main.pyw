from datetime import datetime
import pypinyin
from pypinyin import Style
import sys
from os import listdir, path, startfile, makedirs, system, rename, remove, system
from pathlib import Path
from shutil import copyfile, move, rmtree
from win32com.client import Dispatch
import json
from PyQt5.QtCore import Qt, QPoint, QMimeData, QUrl
from PyQt5.QtGui import QIcon, QDrag, QPixmap, QPainter
from PyQt5.QtWidgets import (QApplication, QWidget, QListWidget, QFrame, QListWidgetItem, QMenu, QInputDialog, QMessageBox)
from CustomizeForm import SidebarWidget, StyledMessageBox, SettingDialog


# ====================== 列表项类 ====================
class FileListWidgetItem(QListWidgetItem):
    def __init__(self, file_name, file_path, is_setting=False, setting_type=""):
        super().__init__(file_name)
        self.file_path = file_path
        self.is_setting = is_setting
        self.setting_type = setting_type

def getAvailableDrives():
    drives = []
    for drive in range(ord('A'), ord('Z')+1):
        driveName = chr(drive) + ":\\"
        if path.exists(driveName):
            drives.append(driveName)
    return drives
# ==================== 主程序 ====================
class DocumentOrganizer(QWidget):
    def __init__(self):
        super().__init__()
        try:
            with open('config.json', 'r', encoding='utf-8') as config_file:
                config = json.load(config_file)
                self.width = config.get('window_width')
                self.height = config.get('window_height') 
                self.WI = config.get('window_inner_margin_width')
                self.HI = config.get('window_inner_margin_height')
                self.Margins = config.get('window_screen_margins')
        except:
            self.width = 400
            self.height = 350
            self.WI = 12
            self.HI = 12
            self.Margins = 10
        self.dragPosition = None
        self.last_char = ""  # 记录上一次输入的字符
        self.current_match_index = 0  # 记录当前匹配项的索引
        self.pinyin_cache = {}  # 缓存拼音转换结果
        self.input_conversion_cache = {}  # 缓存用户输入的转换结果

        self.msgBox = QMessageBox()
        self.msgBox.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.msgBox.setAttribute(Qt.WA_TranslucentBackground)

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
        self.siderbar = SidebarWidget(self) # TODO 只要绑定self就显示不了
        print(f"""siderbar: \t {self.siderbar.geometry().x()} \t {self.siderbar.geometry().y()} \t {self.siderbar.geometry().width()} \t {self.siderbar.geometry().height()}\nmain: \t {self.geometry().x()} \t {self.geometry().y()} \t {self.geometry().width()} \t {self.geometry().height()}""")
        self.siderbar.show()


    def initUI(self):
        self.setWindowTitle("FileIn")
        self.resize(self.width, self.height)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAcceptDrops(True)

        self.QBackgroundFrame.setObjectName('BackgroundFrame')
        self.QBackgroundFrame.setGeometry(0, 0, self.width, self.height)
        self.QBackgroundFrame.mousePressEvent = self.mousePressEvent
        self.QBackgroundFrame.mouseMoveEvent = self.mouseMoveEvent

        self.QFileList.setGeometry(self.WI, self.WI, self.width-self.WI*2, self.height-self.HI*2)
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
                if item.setting_type == "back":
                    self.loadList()
                    return
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
                
                if FileList:  # 非空目录
                    FileTypeList = [path.splitext(File)[1][1:].lower() for File in FileList]
                else:  # 空目录
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

            
        if (hasattr(item, 'is_setting') and item.is_setting) or FileName == "Back":
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

        if action == backAction:
            self.navigateUp()

        elif action == UpdateAction:  # 更新
            if self.QFileList.item(0).text() == "C:\\":
                self.loadList(getAvailableDrives())
            else:
                self.loadList()

        elif action == OpenFilePathAction:  # 打开所在位置
            try:
                if path.isdir(itemPath):  # 目录
                    startfile(itemPath)
                else:  # 文件
                    print(itemPath)
                    system(f'explorer /select,"{itemPath}"')  # 开启资源管理器并且选中文件
            except Exception as e:
                self.showError("错误", f"打开失败: {str(e)}")

        elif action == renameAction:  # 重命名
            newFileName, okFlag = QInputDialog.getText(self, "重命名", "请输入新的名称:")
            if okFlag and newFileName:
                newPath = path.join(self.DestinationFolder, newFileName)
                try:
                    rename(itemPath, newPath)
                    item.setText(newFileName)
                except Exception as e:
                    self.showError("错误", f"重命名失败: {str(e)}")

        elif action == deleteAction:  # 删除
            answer = self.msgBox.question(self, "确认", "您确定要删除该文件吗？", QMessageBox.Yes | QMessageBox.No)
            if answer == QMessageBox.Yes:
                try:
                    if path.isdir(itemPath):
                        rmtree(itemPath)
                    else:
                        remove(itemPath)
                    self.QFileList.takeItem(self.QFileList.row(item))
                except Exception as e:
                    self.showError("错误", f"删除失败: {str(e)}")

        elif action == removeAction:  # 移出
            try:
                move(itemPath, Path.home()) # 将文件移入桌面
                self.loadList()
            except Exception as e:
                self.showError("Move Error:", str(e))

        elif action == settingAction:  # 打开设置
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

        elif action == exitAction:  # 退出程序
            self.close()
            sys.exit()

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
            self.loadList()
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

    def moveEvent(self, event):
        """主窗口移动时更新侧边栏位置"""
        super().moveEvent(event)
        if self.siderbar.attached_side:  # 如果处于吸附状态
            self.siderbar.update_position()

    def resizeEvent(self, event):
        """主窗口调整大小时更新侧边栏位置"""
        super().resizeEvent(event)
        if self.siderbar.attached_side:  # 如果处于吸附状态
            self.siderbar.update_position()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            delta = event.globalPos() - self.dragPosition
            screen = QApplication.primaryScreen().availableGeometry()
            new_x = max(self.Margins, min(delta.x(), screen.width() - self.width - self.Margins))
            new_y = max(self.Margins, min(delta.y(), screen.height() - self.height - self.Margins))
            self.move(new_x, new_y)
    
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
        with open("error.log", "a", encoding="utf-8") as log:
            log.write(f"{datetime.now()} - {title}: {message}\n")
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
        
        with open("style.qss", "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
        doc = DocumentOrganizer()
        doc.show()
        
        app.exec_()
    except Exception as e:
        sys.exit(1)
