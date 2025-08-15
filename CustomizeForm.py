from os import path
from PyQt5.QtWidgets import QDialog, QFrame, QLabel, QPushButton, QVBoxLayout
from PyQt5.QtCore import Qt, QPoint, QRect, pyqtSignal, QSize
from PyQt5.QtGui import QIcon

class StyledMessageBox(QDialog):
    def __init__(self, parent=None, title="提示", message="", buttons=[]):
        super().__init__(parent)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.TipWindow = QFrame(self)
        self.setStyleSheet('''
            QFrame { background: rgba(50, 50, 50, 220); border-radius: 10px; }
            QLabel { color: white; background: transparent; font: 18px "黑体"; }
            QPushButton { background: rgb(70, 70, 70); color: white; font: 14px "黑体"; border-radius: 5px; padding: 8px 20px; min-width: 80px; }
            QPushButton:hover { background: rgb(90, 90, 90); }
        ''')
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

# ====================== 侧边栏类 ====================
class SidebarWidget(QFrame):
    # 定义固定状态变更信号
    pin_state_changed = pyqtSignal(bool)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_widget = parent
        self.pinned = False
        self.attachbarFrame = QFrame(self)
        self.config = {}
        self.custom_buttons = []
        self.is_dragging = False
        self.drag_start_pos = QPoint()
        self.original_pos = QPoint()
        self.adsorb_threshold = 15
        self.adsorbed_edge = None
        self.relative_offset = QPoint(0, 0)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.layout = QVBoxLayout(self.attachbarFrame)
        self.layout.setSpacing(10)
        self.layout.setContentsMargins(5, 5, 5, 5)
        self.layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.main_layout.addWidget(self.attachbarFrame)
        self.btns = { "pinnedButton": [self.pinning, ["resource/img/UI/MynauiPinSolid-grey.png", "resource/img/UI/MynauiPinSolid-blue.png"], 0] }
        self.initUI()
        self.addButton()
    def initUI(self):
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setFixedWidth(200)
        self.setMinimumHeight(50)
        self.setObjectName("SidebarWidget")
        self.attachbarFrame.setObjectName("attachbarFrame")
        self.setStyleSheet(""" #SidebarWidget { background: rgba(0, 0, 0, 100); border-radius: 8px; padding: 5px;} #attachbarFrame {background: transparent;background: rgba(0, 0, 0, 100); border-radius: 8px;} """)
        self.setCursor(Qt.ArrowCursor)
    def addButton(self):
        for btnName, btnInfo in self.btns.items():
            btn = QPushButton()
            btn.resize(30, 30)
            icon_path = btnInfo[1][0]
            btn.setIcon(QIcon(icon_path))
            btn.setIconSize(QSize(30, 30))
            btn.setStyleSheet("""QPushButton {background: transparent;border-radius: 0px;padding: 5px;}QPushButton:hover {background: rgba(255, 255, 255, 50);border-radius: 4px;}QPushButton:pressed {background: rgba(255, 255, 255, 100);} """)
            btn.clicked.connect(btnInfo[0])
            self.layout.addWidget(btn)
            self.btns[btnName].append(btn)
    def pinning(self):
        self.pinned = not self.pinned
        self.toggleIcon("pinnedButton")
        self.pin_state_changed.emit(self.pinned)
    def toggleIcon(self, btnName):
        if btnName in self.btns:
            state = self.btns[btnName][2]
            new_state = 1 - state
            new_icon_path = self.btns[btnName][1][new_state]
            self.btns[btnName][3].setIcon(QIcon(new_icon_path))
            self.btns[btnName][2] = new_state
    def mousePressEvent(self, event):
        if (event.button() == Qt.LeftButton and event.pos().x() < self.width() - 20 and 0 < event.pos().y() < self.height()):
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
        """计算新的位置是否需要吸附到父窗口的边缘"""
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
        min_edge = min(distances, key=distances.get)  # 最小距离的边
        min_distance = distances[min_edge]  # 最小距离
        if min_distance < self.adsorb_threshold:  # 吸附阈值
            self.relative_offset = QPoint(0, 0)
            if min_edge == 'left':
                return QPoint(parent_geo.x() - self_geo.width(), new_pos.y()), 'left'
            elif min_edge == 'right':
                return QPoint(parent_geo.x() + parent_geo.width(), new_pos.y()), 'right'
            elif min_edge == 'top':
                return QPoint(new_pos.x(), parent_geo.y() - self_geo.height()), 'top'
            elif min_edge == 'bottom':
                return QPoint(new_pos.x(), parent_geo.y() + parent_geo.height()), 'bottom'
        return new_pos, None