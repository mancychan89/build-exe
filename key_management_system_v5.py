# pyinstaller.exe -F -w --add-data "msyh.ttc;." --add-data "app_icon.ico;." --exclude PyQt5 --icon=app_icon.ico .\key_management_system_v5.py --name 还你门匙

import sys
import os
import pyautogui
import pygetwindow as gw
import pyperclip
import xlrd
from pathlib import Path
import barcode
from barcode.writer import ImageWriter
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QDialog, QWidget, QLabel,
    QPushButton, QListWidget, QLineEdit, QVBoxLayout, QHBoxLayout,
    QMessageBox, QGroupBox, QStyleFactory, QFrame, QSystemTrayIcon, QMenu,
    QListWidgetItem, QAbstractItemView
)
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor, QPixmap, QAction
from PyQt6.QtCore import Qt, QSize, QTimer
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
import pkg_resources

# 用于单实例检查和通信的模块
import ctypes
from ctypes import wintypes

# 数据库和Excel相关导入
import sqlite3
import pandas as pd
from collections import defaultdict

# 唯一标识，用于互斥锁和本地通信
APP_UNIQUE_NAME = "KeyManagementSystem_v4_8F7E1D3C"

# 数据库文件路径
DB_PATH = r"D:\wdconfig\db\database.sqlite"
ROOM_EXCEL_PATH = r'D:\wdconfig\room.xls'
USE_EXCEL_PATH = r'D:\wdconfig\USE.xls'


class SingleInstanceChecker:
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


# 获取资源文件路径
def resource_path(relative_path):
    """获取打包后资源的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# 设置应用全局样式
def set_app_style(app):
    # 使用Fusion主题作为基础
    app.setStyle(QStyleFactory.create("Fusion"))

    # 创建自定义调色板
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(240, 248, 255))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 139))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 250, 250))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 248, 255))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(220, 220, 255))
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.black)
    palette.setColor(QPalette.ColorRole.Text, QColor(25, 25, 112))
    palette.setColor(QPalette.ColorRole.Button, QColor(176, 196, 222))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 128))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(70, 130, 180))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)

    app.setPalette(palette)

    # 设置字体
    font_path = resource_path("msyh.ttc")
    if os.path.exists(font_path):
        font = QFont(font_path, 10)
    else:
        font = QFont("Arial", 10)
    font.setBold(False)
    app.setFont(font)


class CustomWriter(ImageWriter):
    def __init__(self):
        super().__init__()
        self.font_path = resource_path("msyh.ttc")
        self.font_size = 10


def generate_product_barcode(data, output_dir=r'C:\Users\taojinzhan\Desktop\钥匙条形码'):
    os.makedirs(output_dir, exist_ok=True)

    try:
        writer = CustomWriter()
        code_class = barcode.get_barcode_class('code128')
        barcode_obj = code_class(data, writer=writer)

        save_options = {
            'module_width': 0.4,
            'module_height': 15,
            'font_size': 6,
            'text_distance': 3,
        }

        filename = os.path.join(output_dir, data)
        barcode_obj.save(filename, save_options)
        return filename
    except Exception as e:
        return str(e)


def one_key(key_list, main_window):
    if key_list and key_list[-1] == '手自动':
        key_list.pop()
        warning = True
    else:
        warning = False

    for key in key_list:
        pyautogui.write(key)
        pyautogui.press('tab')
        QTimer.singleShot(100, lambda: None)

    window_title = "车站钥匙系统 扫码签借"
    windows = gw.getWindowsWithTitle(window_title)

    if not windows:
        window_title = "扫码签还"
        windows = gw.getWindowsWithTitle(window_title)

    if windows:
        win = windows[0]
        if win.isMinimized:
            win.restore()
        win.activate()

        pyautogui.click(x=980, y=550)
        pyautogui.click(x=1160, y=400)
        pyautogui.click(x=1160, y=780)

        if warning:
            show_alert_dialog(main_window)
            return True
    else:
        QMessageBox.warning(main_window, "系统错误", "请检查钥匙系统是否已打开!")

    return warning


def show_alert_dialog(parent):
    dialog = AlertDialog(parent)
    screen_geometry = QApplication.primaryScreen().availableGeometry()
    dialog.move(
        (screen_geometry.width() - dialog.width()) // 2,
        (screen_geometry.height() - dialog.height()) // 2
    )
    dialog.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dialog.activateWindow()
    dialog.raise_()
    dialog.exec()


class AlertDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("提醒")
        self.setFixedSize(300, 180)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setWindowFlags(
            self.windowFlags() |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint
        )

        layout = QVBoxLayout()

        icon_label = QLabel()
        icon_label.setPixmap(
            QApplication.style().standardIcon(QApplication.style().StandardPixmap.SP_MessageBoxWarning).pixmap(48, 48))
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(icon_label)

        label = QLabel('"手/自动" 钥匙需扫码!!!')
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("font-size: 14pt; font-weight: bold; color: #8B0000;")
        layout.addWidget(label)

        button = QPushButton("确定")
        button.setFixedSize(100, 35)
        button.setStyleSheet(
            "QPushButton {"
            "background-color: #4682B4;"
            "color: white;"
            "border-radius: 5px;"
            "font-weight: bold;"
            "}"
            "QPushButton:hover {"
            "background-color: #5F9EA0;"
            "}"
            "QPushButton:pressed {"
            "background-color: #3A5FCD;"
            "}"
        )
        button.clicked.connect(self.accept)
        button.setDefault(True)
        button.setAutoDefault(True)

        button_layout = QHBoxLayout()
        button_layout.addWidget(button)
        button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.setStyleSheet("background-color: #F0F8FF;")
        button.setFocus()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            self.accept()
        else:
            super().keyPressEvent(event)

    def showEvent(self, event):
        super().showEvent(event)
        self.activateWindow()
        self.raise_()


class UserSelectionDialog(QDialog):
    """用户选择对话框"""

    def __init__(self, users, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择使用人员")
        self.setFixedSize(300, 200)  # 缩小窗口尺寸
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        self.selected_user = None

        # 创建布局
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(5)

        # 创建列表框
        self.user_list = QListWidget()
        self.user_list.addItems(users)
        self.user_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.user_list.setStyleSheet(
            "QListWidget {"
            "background-color: #FFFFFF;"
            "border: 1px solid #B0C4DE;"
            "border-radius: 3px;"
            "}"
            "QListWidget::item {"
            "padding: 10px;"
            "font-size: 14pt;"  # 增大字体大小
            "}"
            "QListWidget::item:selected {"
            "background-color: #4682B4;"
            "color: white;"
            "}"
        )

        # 默认选择第一个用户
        if self.user_list.count() > 0:
            self.user_list.setCurrentRow(0)
            self.selected_user = self.user_list.item(0).text()

        layout.addWidget(self.user_list)

        self.setLayout(layout)

        # 连接双击事件
        self.user_list.itemDoubleClicked.connect(self.on_item_double_clicked)
        # 连接项目选择变化事件
        self.user_list.itemSelectionChanged.connect(self.on_selection_changed)

    def on_selection_changed(self):
        """当选择变化时更新选中的用户"""
        selected_items = self.user_list.selectedItems()
        if selected_items:
            self.selected_user = selected_items[0].text()

    def on_item_double_clicked(self, item):
        """双击列表项时触发"""
        self.selected_user = item.text()
        self.accept()

    def keyPressEvent(self, event):
        """处理键盘事件"""
        if event.key() == Qt.Key.Key_Return or event.key() == Qt.Key.Key_Enter:
            # 回车键确认选择
            self.accept()
        elif event.key() == Qt.Key.Key_Escape:
            # ESC键取消
            self.reject()
        else:
            super().keyPressEvent(event)

    def accept(self):
        """确认选择"""
        # 如果没有选中任何用户，尝试使用默认选中的第一个
        if not self.selected_user and self.user_list.count() > 0:
            self.selected_user = self.user_list.item(0).text()
        super().accept()


# 数据库和Excel操作函数
def load_and_merge_data():
    """加载并合并数据"""
    try:
        with sqlite3.connect(DB_PATH) as conn:
            log_df = pd.read_sql_query("SELECT * FROM loging", conn)

        room_df = pd.read_excel(ROOM_EXCEL_PATH)

        return pd.merge(
            log_df,
            room_df,
            left_on="房间",
            right_on="钥匙名称",
            how="left"
        )[['使用人员', '存放地址']]
    except Exception as e:
        print(f"加载数据时出错: {str(e)}")
        return None


def group_addresses_by_user(merged_data):
    """按用户分组存放地址"""
    user_address_map = defaultdict(list)
    for _, row in merged_data.iterrows():
        user_address_map[row['使用人员']].append(row['存放地址'])
    return dict(user_address_map)


def get_user_id(username):
    """获取用户胸卡号"""
    try:
        user_df = pd.read_excel(USE_EXCEL_PATH)
        user_df = user_df[['姓名', '胸卡号']]
        result = user_df[user_df['姓名'] == username]['胸卡号']
        return result.values[0] if not result.empty else None
    except Exception as e:
        print(f"获取用户ID时出错: {str(e)}")
        return None


def return_keys_by_user(username):
    """按用户归还钥匙"""
    try:
        # 加载并合并数据
        merged_data = load_and_merge_data()
        if merged_data is None or merged_data.empty:
            return False, "没有找到钥匙数据"

        # 按用户分组地址
        user_address_map = group_addresses_by_user(merged_data)
        print(merged_data)
        if username not in user_address_map:
            return False, f"没有找到用户 {username} 的钥匙信息"

        # 获取用户胸卡号
        user_id = get_user_id(username)
        if user_id is None:
            return False, f"没有找到用户 {username} 的胸卡号"

        # 获取用户的钥匙列表
        keys = user_address_map[username]
        print(keys)
        if not keys:
            return False, f"用户 {username} 没有借用任何钥匙"

        # 自动输入钥匙编号
        for key in keys:
            pyautogui.write(str(key))
            pyautogui.press('tab')
            QTimer.singleShot(100, lambda: None)

        # 查找并激活扫码签还窗口
        window_title = "扫码签还"
        windows = gw.getWindowsWithTitle(window_title)

        if not windows:
            return False, "未找到'扫码签还'窗口"

        win = windows[0]
        if win.isMinimized:
            win.restore()
        win.activate()

        # 模拟点击操作
        pyautogui.click(x=980, y=550)
        pyautogui.click(x=1160, y=400)
        # pyautogui.write(str(user_id))
        pyperclip.copy(str(user_id))
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.click(x=1160, y=440)
        pyautogui.moveTo(x=1160, y=780)

        return True, f"已成功处理用户 {username} 的钥匙归还"

    except Exception as e:
        return False, f"处理钥匙归还时出错: {str(e)}"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("还你门匙系统")
        self.setMinimumSize(400, 300)

        # 设置窗口图标
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # 初始化单位列表，添加"0. 使用人归还钥匙"选项
        self.units = ['使用人归还钥匙', '房开（巡检）【环调卡】', '西屋', '维修三', '微波炉', 'D5',
                      'G', '信号（7条电缆井）', '房开（环控一串 & 环调卡）', '设计院']
        self.key_list = [
            [],  # 索引0对应"使用人归还钥匙"，不需要钥匙列表
            ['508111', '508249', '508020', '508030', '508040', '508008', '508068', '508069', '508071', '手自动'], # 房开（巡检）【环调卡】
            ['508306', '508069', '手自动'], # 西屋
            ['508030', '508071'], # 维修三
            ['508296'], # 微波炉
            ['508318', '508319', '508321', '508326'], # D5
            ['508317'], # G
            ['508029', '508009', '508101', '508108', '508052', '508054', '508089'], # 信号（7条电缆井）
            ['508111', '508249'],  # 房开（环控一串 & 环调卡）
            ['508249', '508049', '508055', '508058', '508075', '508079', '508087', '508093', '手自动'] # 设计院
        ]

        self.key_code_list = self.load_key_codes()
        self.init_ui()

        # 初始化系统托盘
        self.tray_icon = None
        self.init_tray_icon()

    def init_tray_icon(self):
        """初始化系统托盘图标"""
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray_icon = QSystemTrayIcon(self)
            # 设置托盘图标
            icon_path = resource_path("app_icon.ico")
            if os.path.exists(icon_path):
                self.tray_icon.setIcon(QIcon(icon_path))
            else:
                # 使用默认图标
                self.tray_icon.setIcon(QApplication.style().standardIcon(
                    QApplication.style().StandardPixmap.SP_ComputerIcon))

            # 创建托盘菜单
            tray_menu = QMenu()

            # 显示窗口动作
            show_action = QAction("显示主窗口", self)
            show_action.triggered.connect(self.show_window)
            tray_menu.addAction(show_action)

            # 隐藏窗口动作
            hide_action = QAction("隐藏窗口", self)
            hide_action.triggered.connect(self.hide_window)
            tray_menu.addAction(hide_action)

            tray_menu.addSeparator()  # 分隔线

            # 退出动作
            exit_action = QAction("退出程序", self)
            exit_action.triggered.connect(self.quit_application)
            tray_menu.addAction(exit_action)

            # 设置托盘菜单
            self.tray_icon.setContextMenu(tray_menu)

            # 连接托盘图标激活信号（双击恢复窗口）
            self.tray_icon.activated.connect(self.on_tray_activated)

            # 显示托盘图标
            self.tray_icon.show()

    def on_tray_activated(self, reason):
        """托盘图标激活事件处理"""
        # 双击托盘图标恢复窗口
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_window()

    def hide_window(self):
        """隐藏窗口到系统托盘"""
        self.hide()
        self.status_label.setText("程序已最小化到系统托盘")

        # 显示提示消息
        if self.tray_icon:
            self.tray_icon.showMessage(
                "还你门匙系统",
                "程序正在后台运行，双击托盘图标可恢复窗口。",
                QSystemTrayIcon.MessageIcon.Information,
                2000  # 显示2秒
            )

    def show_window(self):
        """显示并恢复窗口，同时清空输入框"""
        # 关键修改：恢复窗口前先清空输入框
        self.input_field.clear()

        # 恢复窗口状态
        self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.show()
        self.raise_()  # 窗口置顶
        self.activateWindow()  # 激活窗口
        self.status_label.setText("窗口已恢复")

    def quit_application(self):
        """退出应用程序"""
        if self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        """重写关闭事件，使其最小化到托盘而不是退出"""
        event.ignore()  # 忽略关闭事件
        self.hide_window()  # 隐藏到托盘

    def changeEvent(self, event):
        """处理窗口状态变化事件，特别是最小化事件"""
        if event.type() == event.Type.WindowStateChange:
            # 如果窗口被最小化
            if self.windowState() & Qt.WindowState.WindowMinimized:
                # 延迟隐藏窗口，给用户视觉反馈
                QTimer.singleShot(100, self.hide_window)
        super().changeEvent(event)

    def load_key_codes(self):
        """从room.xls加载钥匙代码列表"""
        file_path = Path(r"D:\wdconfig\room.xls")
        if file_path.exists():
            try:
                workbook = xlrd.open_workbook(file_path)
                sheet = workbook.sheet_by_index(0)
                a_column = sheet.col_values(0)

                # 过滤掉空值、空字符串，并转换为整数再转字符串
                result = []
                for a in a_column[1:]:  # 跳过标题行
                    # 检查是否为空值或空字符串
                    if a is not None and str(a).strip() != '':
                        try:
                            # 尝试转换为整数（处理浮点数如 508111.0 → 508111）
                            result.append(str(int(float(a))))
                        except (ValueError, TypeError):
                            # 转换失败则跳过
                            continue
                return result
            except Exception as e:
                print(f"加载钥匙代码失败: {e}")
                return []
        return []

    def init_ui(self):
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        title_label = QLabel("还你门匙系统")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet(
            "font-size: 24pt;"
            "font-weight: bold;"
            "color: #2F4F4F;"
            "margin: 10px 0 20px 0;"
        )
        main_layout.addWidget(title_label)

        hbox = QHBoxLayout()
        main_layout.addLayout(hbox)

        left_group = QGroupBox("单位列表")
        left_group.setStyleSheet(
            "QGroupBox {"
            "font-weight: bold;"
            "border: 1px solid #87CEEB;"
            "border-radius: 5px;"
            "margin-top: 10px;"
            "}"
            "QGroupBox::title {"
            "subcontrol-origin: margin;"
            "subcontrol-position: top center;"
            "padding: 0 5px;"
            "background-color: transparent;"
            "color: #2F4F4F;"
            "}"
        )
        left_layout = QVBoxLayout()
        left_group.setLayout(left_layout)
        hbox.addWidget(left_group, 2)

        self.unit_list = QListWidget()
        # 添加单位列表项，从0开始编号
        self.unit_list.addItems([f"{i}. {unit}" for i, unit in enumerate(self.units)])
        self.unit_list.setStyleSheet(
            "QListWidget {"
            "background-color: #FFFFFF;"
            "border: 1px solid #B0C4DE;"
            "border-radius: 3px;"
            "}"
            "QListWidget::item {"
            "padding: 8px;"
            "}"
            "QListWidget::item:selected {"
            "background-color: #4682B4;"
            "color: white;"
            "}"
        )
        self.unit_list.setSpacing(3)
        self.unit_list.itemDoubleClicked.connect(self.on_unit_selected)
        left_layout.addWidget(self.unit_list)

        right_group = QGroupBox("操作区")
        right_group.setStyleSheet(left_group.styleSheet())
        right_layout = QVBoxLayout()
        right_group.setLayout(right_layout)
        hbox.addWidget(right_group, 1)

        input_group = QGroupBox("钥匙编号")
        input_group_layout = QVBoxLayout()
        input_group.setLayout(input_group_layout)
        right_layout.addWidget(input_group)

        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("输入钥匙编号或单位序号...")
        self.input_field.setStyleSheet(
            "QLineEdit {"
            "background-color: white;"
            "border: 1px solid #B0C4DE;"
            "border-radius: 3px;"
            "padding: 8px;"
            "font-size: 12pt;"
            "}"
        )
        self.input_field.returnPressed.connect(self.on_action)
        input_group_layout.addWidget(self.input_field)

        button_layout = QHBoxLayout()
        right_layout.addLayout(button_layout)

        self.action_button = QPushButton("借还钥匙")
        self.action_button.setStyleSheet(
            "QPushButton {"
            "background-color: #4682B4;"
            "color: white;"
            "border: none;"
            "border-radius: 5px;"
            "padding: 10px;"
            "font-weight: bold;"
            "}"
            "QPushButton:hover {"
            "background-color: #5F9EA0;"
            "}"
            "QPushButton:pressed {"
            "background-color: #3A5FCD;"
            "}"
        )
        self.action_button.clicked.connect(self.on_action)
        button_layout.addWidget(self.action_button)

        self.barcode_button = QPushButton("生成条码")
        self.barcode_button.setStyleSheet(
            "QPushButton {"
            "background-color: #32CD32;"
            "color: white;"
            "border: none;"
            "border-radius: 5px;"
            "padding: 10px;"
            "font-weight: bold;"
            "}"
            "QPushButton:hover {"
            "background-color: #3CB371;"
            "}"
            "QPushButton:pressed {"
            "background-color: #228B22;"
            "}"
        )
        self.barcode_button.clicked.connect(self.on_barcode_action)
        button_layout.addWidget(self.barcode_button)

        status_frame = QFrame()
        status_frame.setFrameShape(QFrame.Shape.StyledPanel)
        status_frame.setStyleSheet("background-color: #E6E6FA; border-top: 1px solid #87CEEB;")
        status_layout = QHBoxLayout()
        status_frame.setLayout(status_layout)
        main_layout.addWidget(status_frame)

        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #2F4F4F; padding: 5px;")
        status_layout.addWidget(self.status_label)

        help_label = QLabel(
            "其它输入方式：\n1. 输入钥匙编号直接借用\n2. 点击生成条码按钮生成钥匙条形码\n3. 输入0并回车按使用人归还钥匙")
        help_label.setStyleSheet("font-size: 10pt; color: #696969; margin-top: 10px;")
        help_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(help_label)

    def showEvent(self, event):
        """窗口显示事件，确保输入框获得焦点"""
        super().showEvent(event)
        self.input_field.setFocus()

    def on_action(self):
        input_text = self.input_field.text().strip()
        if not input_text:
            return

        # 检查是否输入0，触发按使用人归还钥匙功能
        if input_text == '0':
            self.on_return_keys_by_user()
            return

        if input_text in self.key_code_list:
            self.input_field.clear()
            pyautogui.write(input_text)
            pyautogui.press('tab')
            self.status_label.setText(f"已输入钥匙编号: {input_text}")
            return

        # 检查是否输入的是单位序号（1-7）
        if input_text.isdigit():
            index = int(input_text)
            if 1 <= index < len(self.units):
                self.on_unit_selected(self.unit_list.item(index))
                return

        self.status_label.setText("输入有误，请重新输入...")

    def on_barcode_action(self):
        input_text = self.input_field.text().strip()
        if not input_text:
            self.status_label.setText("请输入有效的钥匙编码")
            return

        filename = generate_product_barcode(input_text)
        if os.path.exists(filename + ".png"):
            self.status_label.setText(f"条码已生成: {filename}.png")
            msg_box = QMessageBox(
                QMessageBox.Icon.Information,
                "条码生成成功",
                f"钥匙条形码已生成:\n{filename}.png",
                QMessageBox.StandardButton.Ok,
                self
            )
            icon_path = resource_path("app_icon.ico")
            if os.path.exists(icon_path):
                msg_box.setWindowIcon(QIcon(icon_path))
            msg_box.exec()
        else:
            self.status_label.setText(f"条码生成失败: {filename}")

    def on_unit_selected(self, item):
        if item is None:
            return

        index = self.unit_list.row(item)
        if index == 0:  # 0. 使用人归还钥匙
            self.on_return_keys_by_user()
            return

        if index >= 1 and index < len(self.key_list):
            self.status_label.setText(f"正在处理单位: {self.units[index]}")
            QTimer.singleShot(500, lambda: self.execute_unit_action(index))

    def execute_unit_action(self, index):
        original_key_list = self.key_list[index].copy()
        warning = one_key(original_key_list, self)
        if warning:
            self.status_label.setText(f"处理单位: {self.units[index]} - 需扫码!")
            # 添加关闭窗口的逻辑
            QTimer.singleShot(1000, self.close)
        else:
            self.status_label.setText(f"处理单位: {self.units[index]} - 完成!")
            # 添加关闭窗口的逻辑
            QTimer.singleShot(1000, self.close)

    def on_return_keys_by_user(self):
        """按使用人归还钥匙功能"""
        try:
            self.status_label.setText("正在加载用户数据...")
            QApplication.processEvents()  # 刷新界面

            # 加载并合并数据
            merged_data = load_and_merge_data()
            if merged_data is None or merged_data.empty:
                QMessageBox.warning(self, "数据加载失败", "未能加载钥匙数据，请检查数据库连接")
                self.status_label.setText("就绪")
                return

            # 获取所有使用人员
            users = merged_data['使用人员'].dropna().unique().tolist()
            if not users:
                QMessageBox.warning(self, "无数据", "没有找到任何使用人员的钥匙借用记录")
                self.status_label.setText("就绪")
                return

            # 创建用户选择对话框
            dialog = UserSelectionDialog(users, self)
            dialog.setWindowTitle("选择使用人员")

            # 显示对话框
            if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected_user:
                selected_user = dialog.selected_user
                self.status_label.setText(f"正在处理用户 {selected_user} 的钥匙归还...")
                QApplication.processEvents()  # 刷新界面

                # 执行钥匙归还操作
                success, message = return_keys_by_user(selected_user)

                # 显示结果（不使用弹窗）
                if success:
                    self.status_label.setText(f"已完成用户 {selected_user} 的钥匙归还")
                else:
                    self.status_label.setText(f"处理失败: {message}")

                # 关闭窗口
                QTimer.singleShot(1000, self.close)
            else:
                self.status_label.setText("已取消操作")

        except Exception as e:
            QMessageBox.warning(self, "操作失败", f"处理钥匙归还时出错: {str(e)}")
            self.status_label.setText("就绪")


if __name__ == "__main__":
    # 创建单实例检查器
    instance_checker = SingleInstanceChecker(APP_UNIQUE_NAME)

    # 检查是否已有实例运行
    if instance_checker.is_already_running():
        # 如果已有实例，发送显示窗口命令并退出当前实例
        instance_checker.send_show_command()
        sys.exit(0)

    app = QApplication(sys.argv)

    # 重要：设置应用程序不在最后一个窗口关闭时退出
    app.setQuitOnLastWindowClosed(False)

    # 创建主窗口
    window = MainWindow()

    # 设置本地服务器，用于接收显示窗口命令
    instance_checker.setup_local_server(window.show_window)

    set_app_style(app)
    window.show()
    sys.exit(app.exec())