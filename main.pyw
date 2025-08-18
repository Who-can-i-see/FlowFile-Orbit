from datetime import datetime
import subprocess
import pypinyin
from pypinyin import Style
import sys
from os import listdir, path, startfile, makedirs, system, rename, remove, walk
from pathlib import Path
from shutil import copyfile, move, rmtree
from win32com.client import Dispatch
import json
from PyQt5.QtCore import Qt, QPoint, QMimeData, QUrl
from PyQt5.QtGui import QIcon, QDrag, QPixmap, QPainter
from PyQt5.QtWidgets import (QApplication, QWidget, QListWidget, QFrame, QListWidgetItem, 
                             QMenu, QInputDialog, QMessageBox, QLineEdit, QDialog, 
                             QVBoxLayout)
from CustomizeForm import StyledMessageBox, SidebarWidget

# ====================== 列表项类 ====================
class FileListWidgetItem(QListWidgetItem):
    def __init__(self, file_name, file_path):
        super().__init__(file_name)
        self.file_path = file_path

def getAvailableDrives():
    drives = []
    for drive in range(ord('A'), ord('Z')+1):
        driveName = chr(drive) + ":\\"
        if path.exists(driveName):
            drives.append(driveName)
    return drives

# ====================== 主程序类 ====================
class DocumentOrganizer(QWidget):
    def __init__(self):
        super().__init__()
        self.config = self.load_config()
        self.width = self.config['window']['width']
        self.height = self.config['window']['height']
        self.WI = self.config['window']['inner_margin_width']
        self.HI = self.config['window']['inner_margin_height']
        self.Margins = self.config['window']['screen_margins']
        print(f'初始化配置中...')
        print(f'主窗口宽度: {self.width} \n高度: {self.height} \n内边距宽度: {self.WI} \n内边距高度: {self.HI} \n屏幕边距: {self.Margins}')
        self.oI = 5
        self.dragPosition = None
        self.screen = QApplication.primaryScreen().availableGeometry()
        self.last_char = ""
        self.current_match_index = 0
        self.pinyin_cache = {}
        self.input_conversion_cache = {}
        self.search_result_paths = []
        self.is_search_mode = False
        self.msgBox = QMessageBox()
        self.msgBox.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.msgBox.setAttribute(Qt.WA_TranslucentBackground)
        self.DestinationFolder = self.config['RootFolder']
        self.SupportedFileFormats = ['txt', 'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'jpg', 'png', 'bmp', 'gif', 'jpeg', 'zip', 'rar', '7z', 'tar', 'gz', 'bz2', 'iso', 'fba', 'GHO', 'exe', 'chm', 'dll', 'opa', 'cab', 'cat', 'bat', 'py', 'java', 'c', 'cpp', 'php', 'html', 'htm', 'css', 'js', 'json', 'xml', 'yaml', 'yml', 'csv', 'ini', 'log', 'sys', 'bin', 'tmp', 'db', 'mdb', 'md', 'lnk', 'url', 'crdownload', 'prt', 'igs', 'iges', 'jt', 'obj', 'sql', 'key', 'mp4', 'mp3', 'avi', 'ass', 'srt', 'dat', 'dat_old', 'lock', 'cookie', 'mkv', 'mbr']
        self.QBackgroundFrame= QFrame(self)
        self.QFileList = QListWidget(self)
        self.sidebar = SidebarWidget(self)
        self.sidebar.pin_state_changed.connect(self.handle_pin_state)
        self.initUI()
        self.restore_window_position()
        self.sidebar.show()
        self.resizeEvent = self.on_resize

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
        
        self.QFileList.setGeometry(self.WI, self.HI, self.width-self.WI*2, self.height-self.HI*2)
        self.QFileList.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.QFileList.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
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

    def save_window_position(self):
        """保存主窗体绝对位置和侧边栏相对位置到配置文件"""
        self.config['window_position'] = {
            'x': self.pos().x(),
            'y': self.pos().y()
        }
        if hasattr(self, 'sidebar'):
            self.config['sidebar_relative_x'] = self.sidebar.relative_offset.x()
            self.config['sidebar_relative_y'] = self.sidebar.relative_offset.y()
        self.config['RootFolder'] = self.DestinationFolder
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4)
    
    def restore_window_position(self):
        """恢复主窗体绝对位置和侧边栏相对位置"""
        try:
            # 恢复主窗体位置
            if 'window_position' in self.config:
                pos = self.config['window_position']
                self.move(pos['x'], pos['y'])
                print(f"恢复主窗体位置: ({pos['x']}, {pos['y']})")
            # 恢复侧边栏相对位置
            if hasattr(self, 'sidebar'):
                rel_x = self.config.get('sidebar_relative_x', None)
                rel_y = self.config.get('sidebar_relative_y', None)
                if rel_x is not None and rel_y is not None:
                    self.sidebar.relative_offset = QPoint(rel_x, rel_y)
                else:
                    # 默认放在主窗体右侧，间距20
                    self.sidebar.relative_offset = QPoint(self.width + 20, 0)
                new_pos = self.pos() + self.sidebar.relative_offset
                self.sidebar.move(new_pos)
        except Exception as e:
            print(f"恢复窗口和侧边栏位置失败: {e}")

    def handle_pin_state(self, is_pinned):
        """处理侧边栏固定/取消固定状态"""
        if is_pinned:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        self.show()
        if self.sidebar.isVisible():
            self.sidebar.show()

    def onDoubleClick(self, item):
        # 双击打开文件或目录
        try:
            selection = item.text()
            if selection == "Back":  # 返回上一级目录
                self.navigateUp()
                return
            
            # 在搜索模式下使用完整路径
            if self.is_search_mode:
                itemPath = item.file_path
            else:
                itemPath = path.join(self.DestinationFolder, selection)
            
            if path.isdir(itemPath):
                self.DestinationFolder = path.abspath(itemPath)
                self.loadList()
            else:
                self.safeOpenFile(itemPath)
        except Exception as e:
            self.showError("错误", f"打开失败: {str(e)}")
    
    def loadList(self, drives=None):
        self.last_char = ""  # 重置上一次输入字符
        self.current_match_index = 0  # 重置当前匹配项索引
        self.QFileList.clear()  # 清空列表
        
        if drives is not None:  # 显示磁盘列表
            for drive in drives:
                self.QFileList.addItem(drive)
                self.QFileList.item(self.QFileList.count()-1).setIcon(QIcon("resource/img/Root/hdd.png"))
        else:  # 显示当前目录文件列表
            try:
                if self.is_search_mode:  # 搜索模式下显示搜索结果
                    FileList = self.search_result_paths
                else:  # 正常模式下显示当前目录文件列表
                    FileList = listdir(self.DestinationFolder)
                
                if FileList:  # 非空目录
                    FileTypeList = [path.splitext(File)[1][1:].lower() for File in FileList]
                else:  # 空目录
                    FileList = ["Back"]
                    FileTypeList = ["folder"]
                
                for i, (fileName, fileType) in enumerate(zip(FileList, FileTypeList)):
                    if self.is_search_mode:  # 搜索模式下直接使用完整路径
                        filePath = fileName
                        fileName = path.basename(filePath)
                    else:  # 正常模式下使用相对路径
                        filePath = path.join(self.DestinationFolder, fileName)
                    
                    item = FileListWidgetItem(fileName, filePath)
                    
                    # 设置图标
                    icon_path = f"resource/img/{fileType}.png"
                    if not path.exists(icon_path):
                        icon_path = "resource/img/txt.png"
                    
                    item.setIcon(QIcon(icon_path))
                    self.QFileList.addItem(item)
                
                self.search_result_paths = []  # 搜索模式下清空搜索结果
            except Exception as e:
                self.showError("错误", f"无法读取目录: {str(e)}")
    
    def load_config(self):
        """加载配置文件，确保所有字段完整"""
        default_config = {
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
            with open('config.json', 'r', encoding='utf-8') as f:
                user_config = json.load(f)
            def merge_dict(d1, d2):
                for k, v in d2.items():
                    if k in d1 and isinstance(v, dict):
                        merge_dict(d1[k], v)
                    else:
                        d1[k] = v
                return d1
            return merge_dict(default_config, user_config)
        except Exception as e:
            print(f"加载配置失败，使用默认配置: {str(e)}")
            return default_config

    def _merge_config(self, default, new):
        """递归合并配置字典"""
        if isinstance(new, dict) and isinstance(default, dict):
            for key, value in new.items():
                if key in default:
                    if isinstance(value, dict) and isinstance(default[key], dict):
                        default[key] = self._merge_config(default[key], value)
                    else:
                        default[key] = value
            return default
        return new
    
    def navigateUp(self):
        # 导航到上一级目录
        ParentFolder = path.dirname(self.DestinationFolder)
        if ParentFolder == self.DestinationFolder:
            self.loadList(getAvailableDrives())
        else:
            self.DestinationFolder = ParentFolder
            self.loadList()
    
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Slash:  # "/" 键
            self.showSearchDialog()
            return
            
        if event.key() == Qt.Key_Escape:
            if self.is_search_mode:
                self.is_search_mode = False
                self.loadList()
            else:
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
            current_item = self.QFileList.currentItem() if self.QFileList.currentItem() else None
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
                # 如果输入相同字符，增加索引以循环匹配
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

    def on_resize(self, event):
        """处理窗口大小变化事件"""
        # 获取新的窗口尺寸
        new_width = self.geometry().width()
        new_height = self.geometry().height()

        # 更新背景框架大小
        self.QBackgroundFrame.setGeometry(0, 0, new_width, new_height)

        # 更新文件列表大小（减去内边距）
        self.QFileList.setGeometry(
            self.WI, 
            self.HI, 
            new_width - self.WI * 2, 
            new_height - self.HI * 2
        )

        # 调用父类的resizeEvent
        super().resizeEvent(event)

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
    
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            delta = event.globalPos() - self.dragPosition
            new_x = max(self.Margins, min(delta.x(), self.screen.width() - self.width - self.Margins))
            new_y = max(self.Margins, min(delta.y(), self.screen.height() -self.height - self.Margins))
            self.move(new_x, new_y)
            # 侧边栏跟随主窗体移动
            if hasattr(self, 'sidebar'):
                sidebar_new_pos = self.pos() + self.sidebar.relative_offset
                self.sidebar.move(sidebar_new_pos)
            event.accept()
    
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
    
    def setting(self, pos):
        selectedItems = self.QFileList.selectedItems()
        if not selectedItems:
            return
        item = selectedItems[0]
        FileName = item.text()
        if (hasattr(item, 'is_setting') and item.is_setting) or FileName == "Back":
            return
        itemPath = path.join(self.DestinationFolder, FileName) if not self.is_search_mode else item.file_path
        menu = QMenu(self)
        menu.setWindowOpacity(0.7843137)
        backAction = menu.addAction("返回上一目录")
        UpdateAction = menu.addAction("更新")
        OpenFilePathAction = menu.addAction("打开所在位置")
        renameAction = menu.addAction("重命名")
        deleteAction = menu.addAction("删除")
        removeAction = menu.addAction("移出")
        exitAction = menu.addAction("退出")
        action = menu.exec_(self.QFileList.mapToGlobal(pos))
        if action == backAction:
            if self.is_search_mode:
                self.is_search_mode = False
                self.loadList()
            else:
                self.navigateUp()
        elif action == UpdateAction:
            self.loadList()
        elif action == OpenFilePathAction:
            system(f'explorer /select,"{itemPath}"')
        elif action == renameAction:
            new_name, ok = QInputDialog.getText(self, "重命名", "新文件名:", QLineEdit.Normal, FileName)
            if ok and new_name:
                try:
                    new_path = path.join(self.DestinationFolder, new_name)
                    rename(itemPath, new_path)
                    self.loadList()
                except Exception as e:
                    self.showError("错误", f"重命名失败: {str(e)}")
        elif action == deleteAction:
            try:
                if path.isdir(itemPath):
                    rmtree(itemPath)
                else:
                    remove(itemPath)
                self.loadList()
            except Exception as e:
                self.showError("错误", f"删除失败: {str(e)}")
        elif action == removeAction:
            # 这里可自定义移出逻辑
            pass
        elif action == exitAction:
            self.close()

    def execute_command(self, command):
        try:
            process = subprocess.Popen(
                f"chcp 65001>null && del null && cd {self.DestinationFolder} && {command}",
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )  # 在当前目录执行命令
            stdout, stderr = process.communicate()
            # 如果有输出，弹出对话框显示
            if stdout or stderr:
                self.showError("命令执行结果", f"执行命令成功，退出代码：{process.returncode}\n\n{stdout.decode()}\n\n{stderr.decode()}")
        except Exception as e:
            self.showError("错误", f"执行命令失败: {str(e)}")
        except Exception as e:
            self.showError("错误", f"执行命令失败: {str(e)}")

    def showSearchResult(self, keyword):
        """显示搜索结果"""
        self.QFileList.clear()
        self.search_result_paths = []
        
        # 空关键词处理
        if not keyword:
            self.is_search_mode = False
            self.loadList()
            return
            
        # 遍历目录查找匹配文件
        for root, dirs, files in walk(self.DestinationFolder):
            for file in files:
                if keyword.lower() in file.lower():
                    full_path = path.join(root, file)
                    self.search_result_paths.append(full_path)
        
        self.is_search_mode = True
        self.loadList()

    def showSearchDialog(self):
        """显示搜索对话框"""
        dialog = SearchDialog(self)
        
        # 右下角显示
        geo = self.geometry()
        dialog.move(geo.x() + geo.width() - dialog.width() - 20, geo.y() + geo.height() - dialog.height() - 20)
        
        keyword = dialog.getText()
        if keyword is not None:
            if keyword[0] == "/":
                keyword = keyword[1:]
                self.execute_command(keyword)
            else:
                self.showSearchResult(keyword)

    def showError(self, title, message):
        """显示错误消息"""
        msg = StyledMessageBox(self, title, message)
        msg.setWindowTitle(title)
        msg.exec_()

class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedSize(320, 60)
        
        layout = QVBoxLayout(self)
        self.edit = QLineEdit(self)
        layout.addWidget(self.edit)
        
        self.edit.returnPressed.connect(self.accept)
        self.edit.installEventFilter(self)  # 捕获回车键
        
        self.result = None
        
        self.edit.setStyleSheet(""" 
            QLineEdit { 
                background-color: rgba(255, 255, 255, 200); 
                border-radius: 5px;
                border: 0px solid rgba(1,110,255,100);
                color: #000; 
                font-size: 16px; 
                font-family: '黑体'; 
            } 
        """)
    
    def eventFilter(self, obj, event):
        if obj == self.edit and event.type() == event.KeyPress:
            if event.key() == Qt.Key_Escape:
                self.reject()
                return True
        return super().eventFilter(obj, event)
    
    def getText(self):
        if self.exec_() == QDialog.Accepted: # 点击确定按钮
            return self.edit.text()
        return None

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