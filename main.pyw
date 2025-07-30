from datetime import datetime
import pypinyin
from pypinyin import Style
import sys
from os import listdir, path, startfile, makedirs, system, rename, remove, walk

from pathlib import Path
from shutil import copyfile, move, rmtree
from win32com.client import Dispatch
import win32gui
import win32con
import win32ui
from PIL import Image
import json
from PyQt5.QtCore import QRect, Qt, QPoint, QMimeData, QUrl, QSize, pyqtSignal
from PyQt5.QtGui import QIcon, QDrag, QPixmap, QPainter, QPen
from PyQt5.QtWidgets import (QApplication, QWidget, QListWidget, QFrame, QListWidgetItem, 
                             QMenu, QInputDialog, QMessageBox, QLineEdit, QDialog, 
                             QVBoxLayout, QPushButton)
from CustomizeForm import StyledMessageBox, SettingDialog

# 解决DeprecationWarning的兼容性设置
import sip
if hasattr(sip, 'setapi'):
    sip.setapi('QString', 2)
    sip.setapi('QVariant', 2)

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

# ====================== 侧边栏类 ====================
class SidebarWidget(QFrame):
    # 定义固定状态变更信号
    pin_state_changed = pyqtSignal(bool)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent
        self.pinned = False
        self.attachbarFrame = QFrame(self)
        
        # 拖动相关变量
        self.is_dragging = False
        self.drag_start_pos = QPoint()
        self.original_pos = QPoint()
        self.adsorb_threshold = 15
        self.adsorbed_edge = None
        self.relative_offset = QPoint(0, 0)  # 存储与主窗口的相对位置偏移
        
        # 设置布局管理器
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        
        self.layout = QVBoxLayout(self.attachbarFrame)
        self.layout.setSpacing(10)
        self.layout.setContentsMargins(5, 5, 5, 5)
        self.layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        
        self.main_layout.addWidget(self.attachbarFrame)
        
        self.btns = { 
            "pinnedButton": [self.pinning, ["resource/img/UI/MynauiPinSolid-grey.png", "resource/img/UI/MynauiPinSolid-blue.png"], 0] 
        }
        
        self.initUI()
        self.addButton()
    
    def initUI(self):
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # 设置固定宽度和自适应高度
        self.setFixedWidth(200)  # 固定宽度200px
        self.setMinimumHeight(50)  # 最小高度50px
        
        self.setObjectName("SidebarWidget")
        self.attachbarFrame.setObjectName("attachbarFrame")
        
        self.setStyleSheet(""" 
            #SidebarWidget { 
                background: rgba(0, 0, 0, 100); 
                border-radius: 8px; 
                padding: 5px;
            } 
            #attachbarFrame {
                background: transparent;
                background: rgba(0, 0, 0, 100); 
                border-radius: 8px;
            } 
        """)
        
        self.setCursor(Qt.ArrowCursor)
    
    def addButton(self):
        for btnName, btnInfo in self.btns.items():
            btn = QPushButton()
            btn.resize(30, 30)
            
            icon_path = btnInfo[1][0]
            if not path.exists(icon_path):
                icon_path = "resource/img/default_icon.png"
                print(f"警告: 图标文件 {btnInfo[1][0]} 不存在，使用默认图标")
            
            btn.setIcon(QIcon(icon_path))
            btn.setIconSize(QSize(30, 30))
            
            btn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border-radius: 0px;
                    padding: 5px;
                }
                QPushButton:hover {
                    background: rgba(255, 255, 255, 50);
                    border-radius: 4px;
                }
                QPushButton:pressed {
                    background: rgba(255, 255, 255, 100);
                }
            """)
            
            btn.clicked.connect(btnInfo[0])
            self.layout.addWidget(btn)
            self.btns[btnName].append(btn)
    
    def update_custom_buttons(self):
        """更新侧边栏自定义按钮"""
        # 先移除现有自定义按钮
        for btn in self.custom_buttons:
            self.layout.removeWidget(btn)
            btn.deleteLater()
        self.custom_buttons = []

        # 从配置加载并添加自定义按钮
        self.load_sidebar_config()
        if 'sidebar_items' in self.config:
            for item in self.config['sidebar_items']:
                btn = self.create_sidebar_button(item)
                self.custom_buttons.append(btn)
                self.layout.addWidget(btn)
    
    def pinning(self):
        self.pinned = not self.pinned
        self.toggleIcon("pinnedButton")
        self.pin_state_changed.emit(self.pinned)
    
    def toggleIcon(self, btnName):
        if btnName in self.btns:
            state = self.btns[btnName][2]
            new_state = 1 - state
            new_icon_path = self.btns[btnName][1][new_state]
            
            if not path.exists(new_icon_path):
                new_icon_path = "resource/img/default_icon.png"
            
            self.btns[btnName][3].setIcon(QIcon(new_icon_path))
            self.btns[btnName][2] = new_state
    
    def mousePressEvent(self, event):
        if (event.button() == Qt.LeftButton and 
            event.pos().x() < self.width() - 20 and 
            0 < event.pos().y() < self.height()):
            
            self.is_dragging = True
            self.drag_start_pos = event.globalPos()
            self.original_pos = self.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
        else:
            super().mousePressEvent(event)
    
    def mouseMoveEvent(self, event):
        if self.is_dragging and event.buttons() == Qt.LeftButton:
            delta = event.globalPos() - self.drag_start_pos
            new_pos = self.original_pos + delta
            
            adsorbed_edge = None
            if self.parent_widget:
                new_pos, adsorbed_edge = self.calculate_adsorb_position(new_pos)
            
            self.move(new_pos)
            self.adsorbed_edge = adsorbed_edge if adsorbed_edge else self.adsorbed_edge
            event.accept()
        else:
            self.setCursor(Qt.ArrowCursor)
            super().mouseMoveEvent(event)
    
    def mouseReleaseEvent(self, event):
        if self.is_dragging and event.button() == Qt.LeftButton:
            self.is_dragging = False
            self.setCursor(Qt.ArrowCursor)
            
            # 拖动结束后，记录sidebar和主窗口的相对位置偏移
            if self.parent_widget:
                parent_pos = self.parent_widget.pos()
                self_pos = self.pos()
                self.relative_offset = QPoint(
                    self_pos.x() - parent_pos.x(),
                    self_pos.y() - parent_pos.y()
                )
                
            event.accept()
        else:
            super().mouseReleaseEvent(event)
    
    def calculate_adsorb_position(self, new_pos):
        if not self.parent_widget:
            return new_pos, None
            
        parent_global_pos = self.parent_widget.mapToGlobal(QPoint(0, 0))
        parent_geo = QRect(parent_global_pos, self.parent_widget.size())
        self_geo = self.geometry()
        
        distances = {
            'left': abs(new_pos.x() + self_geo.width() - parent_geo.x()),
            'right': abs(parent_geo.x() + parent_geo.width() - new_pos.x()),
            'top': abs(new_pos.y() + self_geo.height() - parent_geo.y()),
            'bottom': abs(parent_geo.y() + parent_geo.height() - new_pos.y())
        }
        
        min_edge = min(distances, key=distances.get)
        min_distance = distances[min_edge]
        
        if min_distance < self.adsorb_threshold:
            self.relative_offset = QPoint(0, 0)  # 吸附时重置偏移量
            if min_edge == 'left':
                return QPoint(parent_geo.x() - self_geo.width(), new_pos.y()), 'left'
            elif min_edge == 'right':
                return QPoint(parent_geo.x() + parent_geo.width(), new_pos.y()), 'right'
            elif min_edge == 'top':
                return QPoint(new_pos.x(), parent_geo.y() - self_geo.height()), 'top'
            elif min_edge == 'bottom':
                return QPoint(new_pos.x(), parent_geo.y() + parent_geo.height()), 'bottom'
        
        return new_pos, None
    
    def paintEvent(self, event):
        super().paintEvent(event)
        
        if self.is_dragging and self.parent_widget and self.adsorbed_edge:
            painter = QPainter(self)
            painter.setPen(QPen(Qt.white, 2, Qt.DashLine))
            
            if self.adsorbed_edge == 'left':
                painter.drawLine(0, 5, 0, self.height() - 5)
            elif self.adsorbed_edge == 'right':
                painter.drawLine(self.width() - 1, 5, self.width() - 1, self.height() - 5)
            elif self.adsorbed_edge == 'top':
                painter.drawLine(5, 0, self.width() - 5, 0)
            elif self.adsorbed_edge == 'bottom':
                painter.drawLine(5, self.height() - 1, self.width() - 5, self.height() - 1)
    
    def show_sidebar_button_menu(self, btn, pos, file_path, file_name):
        """显示侧边栏按钮右键菜单"""
        menu = QMenu(self)
        openAction = menu.addAction("打开文件位置")
        removeAction = menu.addAction("取消侧边栏按钮")

        action = menu.exec_(btn.mapToGlobal(pos))
        if action == openAction:
            self.open_file_location(file_path)
        elif action == removeAction:
            self.remove_sidebar_item(file_name)

    def remove_sidebar_item(self, file_name):
        """从侧边栏移除项目"""
        self.load_sidebar_config()
        if 'sidebar_items' in self.config:
            self.config['sidebar_items'] = [
                item for item in self.config['sidebar_items'] 
                if item['name'] != file_name
            ]
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            self.update_custom_buttons()

# ====================== 主程序类 ====================
class DocumentOrganizer(QWidget):
    def __init__(self):
        super().__init__()
        try:
            with open('config.json', 'r', encoding='utf-8') as config_file:
                self.config = json.load(config_file)
                self.width = self.config['window']['width']
                self.height = self.config['window']['height']
                self.WI = self.config['window']['inner_margin_width']
                self.HI = self.config['window']['inner_margin_height']
                self.Margins = self.config['window']['screen_margins']
        except:
            self.width = 400
            self.height = 350
            self.WI = 12
            self.HI = 12
            self.Margins = 10
        
        self.restore_window_position()
        QApplication.instance().aboutToQuit.connect(self.save_window_position)
        
        self.oI = 5
        self.dragPosition = None
        self.screen = QApplication.primaryScreen().availableGeometry()
        self.last_char = ""
        self.current_match_index = 0
        self.pinyin_cache = {}
        self.input_conversion_cache = {}
        self.search_result_paths = []
        self.is_search_mode = False
        self.custom_buttons = []
        
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
        
        self.QBackgroundFrame= QFrame(self)
        self.QFileList = QListWidget(self)
        
        # 初始化侧边栏
        self.sidebar = SidebarWidget(self)
        self.sidebar.pin_state_changed.connect(self.handle_pin_state)
        
        self.initUI()
        self.update_sidebar_position()
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
        """保存窗口位置到配置文件"""
        self.config['window_position'] = {
            'x': self.pos().x(),
            'y': self.pos().y()
        }
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4)

    def restore_window_position(self):
        """从配置文件恢复窗口位置"""
        if 'window_position' in self.config:
            pos = self.config['window_position']
            self.move(pos['x'], pos['y'])

    def update_sidebar_position(self):
        """初始化默认位置并设置初始相对偏移"""
        main_geometry = self.geometry()
        sidebar_geometry = self.sidebar.geometry()
        
        x = main_geometry.x() + main_geometry.width() + 5
        y = main_geometry.y() + (main_geometry.height() - sidebar_geometry.height()) // 2
        
        self.sidebar.move(x, y)
        
        # 记录初始相对偏移
        self.sidebar.relative_offset = QPoint(
            x - main_geometry.x(),
            y - main_geometry.y()
        )
    
    def on_resize(self, event):
        """主窗口大小变化时更新吸附位置"""
        if hasattr(self, 'sidebar') and self.sidebar.adsorbed_edge:
            self.update_adsorbed_position()
        super().resizeEvent(event)
    
    def moveEvent(self, event):
        """主窗口移动时保持侧边栏相对位置"""
        if hasattr(self, 'sidebar'):
            # 情况1：侧边栏处于吸附状态
            if self.sidebar.adsorbed_edge:
                self.update_adsorbed_position()
                
            # 情况2：侧边栏处于自由状态，使用相对偏移保持位置
            else:
                new_x = self.pos().x() + self.sidebar.relative_offset.x()
                new_y = self.pos().y() + self.sidebar.relative_offset.y()
                self.sidebar.move(new_x, new_y)
                
        super().moveEvent(event)
    
    def update_adsorbed_position(self):
        """吸附状态下更新位置"""
        if not hasattr(self, 'sidebar') or not self.sidebar.adsorbed_edge:
            return
            
        edge = self.sidebar.adsorbed_edge
        sidebar_geo = self.sidebar.geometry()
        parent_global_pos = self.mapToGlobal(QPoint(0, 0))
        parent_width = self.geometry().width()
        parent_height = self.geometry().height()

        # 根据吸附边缘计算位置
        if edge == 'left':
            x = parent_global_pos.x() - sidebar_geo.width()
            y = parent_global_pos.y() + self.sidebar.relative_offset.y()
        elif edge == 'right':
            x = parent_global_pos.x() + parent_width
            y = parent_global_pos.y() + self.sidebar.relative_offset.y()
        elif edge == 'top':
            x = parent_global_pos.x() + self.sidebar.relative_offset.x()
            y = parent_global_pos.y() - sidebar_geo.height()
        elif edge == 'bottom':
            x = parent_global_pos.x() + self.sidebar.relative_offset.x()
            y = parent_global_pos.y() + parent_height
        else:
            return
            
        self.sidebar.move(x, y)
    
    def handle_pin_state(self, is_pinned):
        if is_pinned:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        
        self.show()
        if self.sidebar.isVisible():
            self.sidebar.show()
    
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
        self.last_char = ""
        self.current_match_index = 0
        self.QFileList.clear()
        
        if drives is not None:
            for drive in drives:
                self.QFileList.addItem(drive)
                self.QFileList.item(self.QFileList.count()-1).setIcon(QIcon("resource/img/Root/hdd.png"))
        else:
            try:
                if self.is_search_mode:
                    FileList = self.search_result_paths
                else:
                    FileList = listdir(self.DestinationFolder)
                
                if FileList:  # 非空目录
                    FileTypeList = [path.splitext(File)[1][1:].lower() for File in FileList]
                else:  # 空目录
                    FileList = ["Back"]
                    FileTypeList = ["folder"]
                
                for fileName, fileType in zip(FileList, FileTypeList):
                    if self.is_search_mode is False:  # 搜索模式下不用再次拼接路径
                        filePath = path.join(self.DestinationFolder, fileName)
                    else:  # 搜索模式下直接使用文件名
                        filePath = fileName
                        fileName = path.basename(filePath)
                    
                    item = FileListWidgetItem(fileName, filePath)
                    icon_path = f"resource/img/{fileType}.png"
                    if path.exists(icon_path):
                        item.setIcon(QIcon(icon_path))
                    else:
                        item.setIcon(QIcon("resource/img/txt.png"))
                    self.QFileList.addItem(item)
                
                self.search_result_paths = []  # 搜索模式下清空搜索结果
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
        addSidebarAction = menu.addAction("添加到侧边栏")
        settingAction = menu.addAction("设置")
        exitAction = menu.addAction("退出")
        
        action = menu.exec_(self.QFileList.mapToGlobal(pos))
        
        if action == backAction:
            if self.is_search_mode:
                self.is_search_mode = False
                self.loadList()
            else:
                self.navigateUp()
        elif action == UpdateAction:  # 更新
            if self.QFileList.item(0).text() == "C:\\":
                self.loadList(getAvailableDrives())
            else:
                self.loadList()
            self.is_search_mode = False
        elif action == OpenFilePathAction:  # 打开所在位置
            try:
                if path.isdir(itemPath):  # 目录
                    startfile(itemPath)
                else:  # 文件
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
                move(itemPath, Path.home())  # 将文件移入桌面
                self.loadList()
            except Exception as e:
                self.showError("Move Error:", str(e))
        if action == addSidebarAction:
            self.add_to_sidebar(item.file_path, FileName)
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

    def add_to_sidebar(self, file_path, file_name):
        """添加文件到侧边栏"""
        # 获取文件图标
        icon_path = self.get_file_icon(file_path)

        # 保存到配置
        if 'sidebar_items' not in self.config:
            self.config['sidebar_items'] = []

        # 检查是否已存在
        for item in self.config['sidebar_items']:
            if item['path'] == file_path:
                QMessageBox.information(self, "提示", "该文件已在侧边栏中")
                return

        self.config['sidebar_items'].append({
            'name': file_name,
            'path': file_path,
            'icon': icon_path,
            'use_builtin_icon': False
        })

        # 更新配置文件
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4)

        # 更新侧边栏显示
        self.sidebar.update_custom_buttons()

    def get_file_icon(self, file_path, size=(32, 32)):
        """获取文件系统原生图标"""
        # 创建临时图标保存目录
        if not path.exists("temp_icons"):
            makedirs("temp_icons")

        # 生成唯一的临时文件名
        icon_hash = hash(file_path)
        temp_icon_path = f"temp_icons/icon_{icon_hash}_{size[0]}x{size[1]}.png"

        # 如果临时图标已存在，直接返回
        if path.exists(temp_icon_path):
            return temp_icon_path

        try:
            # 获取文件属性
            attrs = win32gui.SHGetFileInfo(
                file_path,
                win32con.FILE_ATTRIBUTE_NORMAL,
                (0, 0, 0, 0, win32con.SHGFI_ICON | win32con.SHGFI_SMALLICON)
            )

            # 提取图标句柄
            hicon = attrs[0]

            # 将图标转换为位图
            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, size[0], size[1])
            hdc = hdc.CreateCompatibleDC()
            hdc.SelectObject(hbmp)
            hdc.DrawIcon((0, 0), hicon)

            # 保存位图到内存
            bmp_info = hbmp.GetInfo()
            bmp_bits = hbmp.GetBitmapBits(True)

            # 使用PIL转换为PNG
            img = Image.frombuffer(
                'RGB',
                (bmp_info['bmWidth'], bmp_info['bmHeight']),
                bmp_bits, 'raw', 'BGRX', 0, 1
            )

            # 保存到临时文件
            img.save(temp_icon_path, "PNG")

            # 释放资源
            win32gui.DestroyIcon(hicon)

            return temp_icon_path

        except Exception as e:
            print(f"获取文件图标失败: {e}")
            return "resource/img/file.png"  # 返回默认图标

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
    
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            delta = event.globalPos() - self.dragPosition
            new_x = max(self.Margins, min(delta.x(), self.screen.width() - self.width - self.Margins))
            new_y = max(self.Margins, min(delta.y(), self.screen.height() -self.height - self.Margins))
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
    
    def showSearchResult(self, keyword):
        self.QFileList.clear()
        self.search_result_paths = []
        for root, dirs, files in walk(self.DestinationFolder):
            for file in files:
                if keyword.lower() in file.lower():
                    full_path = path.join(root, file)
                    self.search_result_paths.append(full_path)
        self.is_search_mode = True
        self.loadList()

    def showSearchDialog(self):
        dialog = SearchDialog(self)
        
        # 右下角显示
        geo = self.geometry()
        dialog.move(geo.x() + geo.width() - dialog.width() - 20, geo.y() + geo.height() - dialog.height() - 20)
        
        keyword = dialog.getText()
        if keyword:
            self.showSearchResult(keyword)

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
                border-radius: 4px;
                border: 0px solid;
                color: #000; 
                font-size: 16px; 
                font-family: '思源黑体'; 
            } 
        """)
    
    def eventFilter(self, obj, event):
        if obj == self.edit and event.type() == event.KeyPress:
            if event.key() == Qt.Key_Escape:
                self.reject()
                return True
        return super().eventFilter(obj, event)
    
    def getText(self):
        if self.exec_() == QDialog.Accepted:
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
