# pyinstaller.exe -F -w --add-data="logo.png;." -i .\logo.ico --exclude-module PyQt5 .\one_key_d5_v3.py --name 天天D5
import sys
import os
import pyautogui
import pygetwindow as gw
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QGridLayout,
    QLabel, QPushButton, QDialog, QScrollArea, QFrame, QSystemTrayIcon, QMenu
)
# 改为
from PyQt6.QtCore import Qt, pyqtSignal, QPropertyAnimation, QEasingCurve
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
from PyQt6.QtGui import (
    QFont, QIcon, QColor, QPalette, QCursor,
    QPixmap, QLinearGradient, QBrush, QPainter,
    QPen, QPainterPath, QAction
)
from PyQt6.QtWidgets import QGraphicsDropShadowEffect
import ctypes
from ctypes import wintypes

# 唯一标识，用于单实例检查和通信
APP_UNIQUE_NAME = "KeyBorrowSystem_D5_9A3B7C5D"

# 资源路径处理函数
def resource_path(relative_path):
    """获取绝对路径以处理 PyInstaller 单文件打包的资源"""
    try:
        # PyInstaller 创建临时文件夹存储资源
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class SingleInstanceChecker:
    """单实例检查器，确保只有一个程序实例在运行"""
    def __init__(self, unique_name):
        self.unique_name = unique_name
        self.mutex_handle = None
        self.local_server = None

    def is_already_running(self):
        """检查是否已有实例运行"""
        # 创建互斥锁
        mutex_name = f"Global\\{self.unique_name}"

        self.mutex_handle = ctypes.windll.kernel32.CreateMutexW(
            None,  # 默认安全属性
            True,  # 初始拥有者
            mutex_name  # 互斥锁名称
        )

        if self.mutex_handle == 0:
            return True  # 创建失败，可能已存在

        # 检查错误代码
        last_error = ctypes.windll.kernel32.GetLastError()
        if last_error == 183:  # ERROR_ALREADY_EXISTS
            ctypes.windll.kernel32.CloseHandle(self.mutex_handle)
            return True  # 已存在实例

        return False

    def setup_local_server(self, show_window_callback):
        """设置本地服务器用于接收显示窗口命令"""
        self.local_server = QLocalServer()
        # 如果服务器已存在，移除它
        if QLocalServer.removeServer(self.unique_name):
            print("移除已存在的本地服务器")

        if not self.local_server.listen(self.unique_name):
            print(f"无法启动本地服务器: {self.local_server.errorString()}")
            return False

        # 连接新连接信号
        self.local_server.newConnection.connect(self._on_new_connection)
        self.show_window_callback = show_window_callback
        return True

    def _on_new_connection(self):
        """处理新的连接，接收显示窗口命令"""
        client_socket = self.local_server.nextPendingConnection()
        if client_socket:
            client_socket.waitForReadyRead()
            # 读取消息内容
            data = client_socket.readAll().data().decode()
            if data == "show_window":
                # 调用显示窗口的回调函数
                self.show_window_callback()
            client_socket.disconnectFromServer()

    def send_show_command(self):
        """向已运行的实例发送显示窗口命令"""
        socket = QLocalSocket()
        socket.connectToServer(self.unique_name)

        if socket.waitForConnected(1000):
            # 发送显示窗口命令
            socket.write("show_window".encode())
            socket.waitForBytesWritten()
            socket.disconnectFromServer()
            return True
        return False


class AnimatedButton(QPushButton):
    """带缩放动画效果的自定义按钮"""

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFixedSize(100, 40)

        # 设置初始样式
        self.setStyleSheet("""
            QPushButton {
                background-color: #64b4f0;
                color: white;
                border: none;
                border-radius: 10px;
                font: bold 11pt 'Microsoft YaHei';
                transition: all 0.2s ease;
            }
            QPushButton:hover {
                background-color: #50a0dc;
                transform: translateY(-2px);
            }
            QPushButton:pressed {
                transform: translateY(1px);
            }
        """)

        # 添加阴影效果
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setXOffset(0)
        shadow.setYOffset(5)
        shadow.setColor(QColor(100, 180, 240, 150))
        self.setGraphicsEffect(shadow)


class GradientFrame(QFrame):
    """带渐变背景的框架"""

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # 创建渐变背景
        gradient = QLinearGradient(0, 0, self.width(), self.height())
        gradient.setColorAt(0.0, QColor(100, 180, 240))
        gradient.setColorAt(1.0, QColor(50, 80, 120))
        painter.setBrush(gradient)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(self.rect(), 10, 10)


class NameTable(QMainWindow):
    selected = pyqtSignal(str, str)  # 修改：添加姓名参数
    show_requested = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("钥匙借还D5")

        # 使用资源路径函数设置图标
        self.setWindowIcon(QIcon(resource_path("logo.png")))

        self.resize(580, 500)
        self.center()

        # 设置主窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8fbff;
                font-family: 'Microsoft YaHei';
            }
            QLabel {
                font-size: 12pt;
            }
        """)

        # 创建主部件
        main_widget = QWidget()
        main_widget.setObjectName("mainWidget")
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(0)

        # 头部渐变标题
        header = GradientFrame()
        header.setFixedHeight(80)
        header_layout = QVBoxLayout(header)

        title = QLabel("请选择D5")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                color: white;
                font: bold 20pt 'Microsoft YaHei';
                letter-spacing: 2px;
            }
        """)
        header_layout.addWidget(title, 1, Qt.AlignmentFlag.AlignCenter)

        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea { 
                border: none; 
                background: transparent; 
            }
            QScrollBar:vertical { 
                width: 12px; 
                background: transparent; 
            }
            QScrollBar::handle:vertical { 
                background: #a0c0e0; 
                border-radius: 6px; 
                min-height: 30px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)

        # 添加内容容器
        scroll_widget = QWidget()
        scroll_widget.setStyleSheet("background-color: transparent;")
        scroll_layout = QGridLayout(scroll_widget)
        scroll_layout.setSpacing(15)
        scroll_layout.setContentsMargins(25, 25, 25, 25)

        # 人员数据
        names = {
            "徐国威": "107940",
            "姚建华": "108584",
            "关梓杰": "134052",
            "罗伟洪": "127199",
            "韩振宇": "107143",
            "吴东清": "150341",
            "王美懿": "149930",
            "黎文祥": "127166",
            "黎敏": "153858",
            "魏健豪": "153000",
            "廖思海": "153863",
            "吴文超": "120557",
            "钟轶华": "103907",
            "陈孟熙": "116873",
            "郭艳梅": "113303",
        }

        # 需要显示红色的人名列表
        red_names = ["黎敏", "魏健豪", "廖思海"]

        # 创建网格项
        row, col = 0, 0
        btn_font = QFont("Microsoft YaHei", 10)

        for name in names:
            # 创建卡片容器
            card = QFrame()
            card.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border-radius: 8px;
                    border: 1px solid #e0e0e0;
                    padding: 10px;
                }
            """)
            card_layout = QVBoxLayout(card)
            card_layout.setSpacing(12)

            # 姓名标签
            name_label = QLabel(name)
            name_label.setStyleSheet(f"""
                QLabel {{
                    font: bold 14pt 'Microsoft YaHei';
                    color: {'#e53333' if name in red_names else '#464646'};
                }}
            """)
            name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            card_layout.addWidget(name_label)

            # 操作按钮
            button = AnimatedButton("选择")
            button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            button.clicked.connect(lambda _, n=name, i=names[name]: self.on_select(n, i))
            card_layout.addWidget(button, 0, Qt.AlignmentFlag.AlignCenter)

            # 添加到网格布局
            scroll_layout.addWidget(card, row, col)

            # 更新网格位置
            col += 1
            if col > 1:
                col = 0
                row += 1

        scroll_area.setWidget(scroll_widget)

        # 添加到主布局
        main_layout.addWidget(header, 0)
        main_layout.addWidget(scroll_area, 1)

        # 设置主部件
        self.setCentralWidget(main_widget)

        # 创建淡入动画
        self.opacity_anim = QPropertyAnimation(self, b"windowOpacity")
        self.opacity_anim.setDuration(500)
        self.opacity_anim.setStartValue(0.0)
        self.opacity_anim.setEndValue(1.0)
        self.opacity_anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        # 初始化系统托盘
        self.init_tray_icon()

    def center(self):
        """将窗口居中显示"""
        screen = QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        self.move(
            int((screen.width() - size.width()) / 2),
            int((screen.height() - size.height()) / 2)
        )

    def showEvent(self, event):
        """窗口显示时启动淡入动画"""
        super().showEvent(event)
        self.opacity_anim.start()

    def closeEvent(self, event):
        """重写关闭事件，隐藏窗口而不是关闭"""
        event.ignore()
        self.hide_to_tray()

    def init_tray_icon(self):
        """初始化系统托盘图标"""
        # 创建托盘图标
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(resource_path("logo.png")))
        self.tray_icon.setToolTip("钥匙借还D5系统")

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = QAction("显示界面", self)
        show_action.triggered.connect(self.show_from_tray)

        quit_action = QAction("退出", self)
        quit_action.triggered.connect(QApplication.quit)

        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)

        # 连接双击事件
        self.tray_icon.activated.connect(self.tray_icon_activated)

        # 显示托盘图标
        self.tray_icon.show()

    def tray_icon_activated(self, reason):
        """托盘图标激活事件处理"""
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_from_tray()

    def show_from_tray(self):
        """从托盘显示窗口"""
        self.show_window()

    def show_window(self):
        """显示并恢复窗口"""
        # 清除最小化状态并显示窗口
        self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.show()
        self.activateWindow()
        self.raise_()
        self.opacity_anim.start()

    def hide_to_tray(self):
        """隐藏窗口到系统托盘"""
        self.hide()
        self.tray_icon.showMessage(
            "钥匙借还D5",
            "程序已最小化到系统托盘",
            QSystemTrayIcon.MessageIcon.Information,
            2000
        )

    def on_select(self, name, id_val):
        self.selected.emit(name, id_val)  # 修改：发送姓名和ID
        self.hide_to_tray()


class EnhancedReminderDialog(QDialog):
    def __init__(self, message, parent=None):
        super().__init__(parent)
        self.setWindowTitle("系统提醒")

        # 使用资源路径函数设置图标
        self.setWindowIcon(QIcon(resource_path("logo.png")))

        self.setFixedSize(350, 220)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.center()

        # 添加动画效果
        self.setWindowOpacity(0.0)
        self.anim = QPropertyAnimation(self, b"windowOpacity")
        self.anim.setDuration(500)
        self.anim.setStartValue(0.0)
        self.anim.setEndValue(1.0)
        self.anim.setEasingCurve(QEasingCurve.Type.OutBack)
        self.anim.start()

        # 主容器
        container = QFrame(self)
        container.setFixedSize(350, 220)
        container.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 15px;
                border: 1px solid #e0e0e0;
            }
        """)

        # 添加阴影效果
        shadow = QGraphicsDropShadowEffect(container)
        shadow.setBlurRadius(20)
        shadow.setXOffset(0)
        shadow.setYOffset(0)
        shadow.setColor(QColor(0, 0, 0, 60))
        container.setGraphicsEffect(shadow)

        # 主布局
        main_layout = QVBoxLayout(container)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(25, 25, 25, 25)

        # 图标 - 使用资源路径函数
        icon_label = QLabel()
        icon_label.setPixmap(QIcon(resource_path("logo.png")).pixmap(48, 48))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # 消息
        msg_label = QLabel(message)
        msg_label.setFont(QFont("Microsoft YaHei", 11))
        msg_label.setStyleSheet("color: #e53333;")
        msg_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        msg_label.setWordWrap(True)

        # 确定按钮
        ok_button = AnimatedButton("确定")
        ok_button.clicked.connect(self.accept)

        # 组合布局
        main_layout.addWidget(icon_label)
        main_layout.addWidget(msg_label)
        main_layout.addWidget(ok_button, 0, Qt.AlignmentFlag.AlignCenter)

        # 设置布局
        container.setLayout(main_layout)
        layout = QVBoxLayout(self)
        layout.addWidget(container)

    def center(self):
        """将窗口居中显示"""
        screen = QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        self.move(
            int((screen.width() - size.width()) / 2),
            int((screen.height() - size.height()) / 2)
        )

    def mousePressEvent(self, event):
        """支持拖动窗口"""
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_position = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        """支持拖动窗口"""
        if event.buttons() == Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self.drag_position)
            event.accept()


def one_key(name, id_val):  # 修改：添加姓名参数
    # 根据姓名选择不同的钥匙列表
    if name in ["黎敏", "魏健豪", "廖思海", "吴文超", "钟轶华", "陈孟熙", "郭艳梅"]:
        key_list = ['508317']
    else:
        key_list = ['508318', '508319', '508321', '508326']

    for key in key_list:
        pyautogui.write(key)
        pyautogui.press('tab')

    window_title = "车站钥匙系统 扫码签借"
    windows = gw.getWindowsWithTitle(window_title)

    if windows:
        win = windows[0]
        if win.isMinimized:
            win.restore()
        win.activate()

        pyautogui.click(x=1160, y=400)
        pyautogui.write(id_val)  # 修改：使用ID值
        pyautogui.click(x=1160, y=465)
        pyautogui.moveTo(x=1160, y=780)
    else:
        window_title = "扫码签还"
        windows = gw.getWindowsWithTitle(window_title)

        if windows:
            win = windows[0]
            if win.isMinimized:
                win.restore()
            win.activate()

            pyautogui.click(x=1160, y=400)
            pyautogui.write(id_val)  # 修改：使用ID值
            pyautogui.click(x=1160, y=465)
            pyautogui.moveTo(x=1160, y=780)
        else:
            dialog = EnhancedReminderDialog("请检查钥匙系统是否已打开!")
            dialog.exec()


if __name__ == "__main__":
    # 创建单实例检查器
    instance_checker = SingleInstanceChecker(APP_UNIQUE_NAME)

    # 检查是否已有实例运行
    if instance_checker.is_already_running():
        # 如果已有实例，发送显示窗口命令并退出当前实例
        instance_checker.send_show_command()
        sys.exit(0)

    app = QApplication(sys.argv)

    # 设置应用程序图标（影响任务栏图标）
    app.setWindowIcon(QIcon(resource_path("logo.png")))

    # 确保应用在关闭最后一个窗口时不退出
    app.setQuitOnLastWindowClosed(False)

    # 创建主窗口
    window = NameTable()

    # 设置本地服务器，用于接收显示窗口命令
    instance_checker.setup_local_server(window.show_window)

    window.show()
    window.selected.connect(one_key)  # 现在传递姓名和ID
    sys.exit(app.exec())