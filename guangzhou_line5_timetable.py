# pyinstaller.exe -F -w --add-data "time.png;." -i .\time.ico .\guangzhou_line5_timetable.py --name 时刻表助手

import sys
import os
import pandas as pd
import re
import time
import json
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QHBoxLayout, QWidget, QToolTip, QDialog, QListWidget,
                             QListWidgetItem, QLabel, QFrame, QMenu, QFileDialog,
                             QMessageBox, QTextEdit, QLineEdit, QSplitter, QProgressDialog,
                             QInputDialog, QComboBox, QGridLayout, QHeaderView, QTableWidget,
                             QTableWidgetItem)
from PyQt6.QtGui import QFont, QCursor, QAction, QIcon, QColor
from PyQt6.QtCore import Qt, QPoint, QRect, QThread, pyqtSignal, QTimer
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
import ctypes
from ctypes import wintypes
import requests
from bs4 import BeautifulSoup
ItemRole_HasLiveSpecialReminder = Qt.ItemDataRole.UserRole + 1   # 自定义角色

# Glink模块代码 - 合并到主程序中
import uuid
import random
from Crypto.PublicKey import RSA
from Crypto.Cipher import PKCS1_v1_5
from base64 import b64encode
from urllib.parse import quote

# -------------------- 工具函数 --------------------
def fmt_train_id(key: str) -> str:
    """
    统一车次显示格式：去掉多余前导 0，但至少保留 1 个。
    00504 -> 0504
    0123  -> 0123
    1504  -> 1504
    """
    if len(key) <= 1:
        return key
    return key.lstrip('0') or '0'      # 防止全 0 被剥光
# --------------------------------------------------


class Glink:
    def __init__(self):
        self.username = 'taojinzhan'
        self.password = 'Money$100'
        # 个人与目标用户信息
        self.MY_USER_ID = "9d58a9d9-2390-488e-bc7f-87f060f9bbe1"
        self.MY_USER_NAME = "淘金站"

        self.session = requests.Session()
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
            "Content-Type": "application/x-www-form-urlencoded",
        }

        # 目标用户列表
        self.Target_User_List = [
            {'滘口': '6ffd8743-a75e-4acd-8d5a-5d89b9364821'},
            {'坦尾': '96b3192a-b343-47a3-9649-b20007286f89'},
            {'中山八': 'ce9476cd-170b-4a23-835e-b593ee7f22b6'},
            {'西场': '036f25d7-27aa-475f-a21f-b53a868a8383'},
            {'西村': 'e6f7ec93-8cea-48c8-9868-af7b7b637ebc'},
            {'广州火车站': 'fe0605ca-171b-45b9-b60f-369993534bbf'},
            {'小北': '9c57dbaa-a06a-4ed9-a64b-95969fc3dc53'},
            {'淘金': '9d58a9d9-2390-488e-bc7f-87f060f9bbe1'},
            {'区庄': '791713d7-945d-4761-9c51-62ec0cb8dddd'},
            {'动物园': '80cb0e95-d022-44e2-9972-7ecddbf0df99'},
            {'杨箕': '0056bbbf-0a93-4bc9-8f4d-23875de71c30'},
            {'五羊邨': '0da9f4f6-44d8-493d-8ca0-fce7ca6dc0ac'},
            {'珠江新城': 'e0d58a07-3244-46e7-bb95-7a92de4719c5'},
            {'猎德': 'acd796f2-b1ab-4b13-97fa-a311e1efaba8'},
            {'潭村': 'a2b3459e-01bb-49d4-ab67-bc7e92e6265f'},
            {'员村': '4b8030ee-6ce0-4f88-9843-345e41edfa70'},
            {'科韵路': '215388f6-9671-4c5c-b753-1c00b555c34d'},
            {'车陂南': '99df64b8-fd37-4a85-bf6f-83c02bb2a4c5'},
            {'东圃': 'd0ea3b8f-a6f0-4357-bbb1-f0c1a3b126e5'},
            {'三溪': '83438d6a-5e00-4851-b644-0e63169267c1'},
            {'鱼珠': '5b889d49-69c4-4aad-a8f7-c9579d067749'},
            {'大沙地': '3f726a23-9341-432f-b95f-a9b8ba75f285'},
            {'大沙地': 'a5598846-d009-4f34-80d7-df2d72919790'},
            {'文冲': 'f03e4d3b-491a-4c78-bb03-21690d93203f'},
            {'双沙': 'd9c050ef-359a-4526-b276-5cafe81ec0fc'},
            {'庙头': 'c0bb0aed-8776-42a4-996c-8b924861cf48'},
            {'夏园': '196dc98e-974f-48a1-9999-96b8cddbf8e2'},
            {'保盈大道': 'db92b3e6-be9b-48a6-86a3-5fef4d0691dd'},
            {'夏港': '9aaf5f35-3d16-471d-8438-81912d0fe3d8'},
            {'黄埔新港': '592235a0-18ff-4239-a9a1-5d4f013ec1df'},

        ]

        self.TARGET_USER_ID = ""
        self.TARGET_USER_NAME = ""

    def rsa_encrypt(self, plaintext):
        # RSA 公钥
        RSA_PUBLIC_KEY = f"""-----BEGIN PUBLIC KEY-----
        {self.sso_public_key}
        -----END PUBLIC KEY-----"""

        public_key_pem = RSA_PUBLIC_KEY.strip().encode("utf-8")
        public_key = RSA.importKey(public_key_pem)
        cipher = PKCS1_v1_5.new(public_key)
        encrypted_bytes = cipher.encrypt(plaintext.encode("utf-8"))
        return b64encode(encrypted_bytes).decode("utf-8")

    def get_csrf_token(self):
        url = 'https://www.bingolink.biz/sso/login'
        r = self.session.get(url, headers=self.headers)

        # bs4查找 csrf_token 的值
        soup = BeautifulSoup(r.text, 'lxml')
        meta_tag = soup.find('meta', attrs={'name': 'csrf'})
        if meta_tag:
            self.csrf_token = meta_tag.get('content')

        # 使用正则表达式提取 ssoPublicKey 的值
        pattern = r"Global\.ssoPublicKey='([^']+)'"
        match = re.search(pattern, r.text)
        self.sso_public_key = match.group(1)

    def login(self):
        rsa_encrypt_username = self.rsa_encrypt(self.username)
        rsa_encrypt_password = self.rsa_encrypt(self.password)
        data = {
            'csrf_token': self.csrf_token,
            'return_url': '/sso/oauth2/authorize?client_id=clientId&response_type=code&redirect_uri=http://192.168.80.61:8015/ssoclient?return_url=http://192.168.80.61:8015/&v2_compatible=true',
            'username': rsa_encrypt_username,
            'encrypt_type': 'rsa',
            'encrypt_type_username': 'rsa',
            'password': rsa_encrypt_password,
        }

        r = self.session.post(
            "https://www.bingolink.biz/sso/login",
            data=data,
            headers=self.headers,
        )

    def getCurrentUserInfo(self):
        # 检测是否登录成功
        r = self.session.get("https://www.bingolink.biz/webos/imUam/getCurrentUserInfo", headers=self.headers)

        if '淘金站' in r.text:
            return True, '登陆成功, 淘金站!'
        else:
            return False, '登陆失败...'

    def set_target_user(self, user_name):
        """设置目标用户"""
        for user in self.Target_User_List:
            if user_name in user:
                self.TARGET_USER_NAME = user_name
                self.TARGET_USER_ID = user[user_name]
                return True
        return False

    def send_message(self, message):
        msg_id = str(uuid.uuid4()).upper()
        encoded_content = quote(quote(message, encoding="utf-8"), encoding="utf-8")
        send_time = datetime.now().strftime("%H:%M:%S")

        message_params = {
            "id": msg_id,
            "fromId": self.MY_USER_ID,
            "fromName": self.MY_USER_NAME,
            "fromType": 1,
            "toId": self.TARGET_USER_ID,
            "toName": self.TARGET_USER_NAME,
            "toType": 1,
            "content": encoded_content,
            "msgType": 1,
            "noticePic": "progs.gif",
            "isLocal": "true",
            "noticeTitle": "正在发送中...",
            "sendTime": send_time,
            "fromCompany": self.MY_USER_NAME
        }

        # 发送消息
        random_t = random.random()
        send_api = f"https://www.bingolink.biz/webos/imMsg/send?t={random_t}"

        try:
            response = self.session.post(
                send_api,
                data=message_params,
                headers=self.headers,
                timeout=30,
                allow_redirects=False
            )

            if response.text:
                return True, '消息发送成功！'
            else:
                return False, '消息发送失败！'
        except Exception as e:
            return False, f'发送异常: {str(e)}'


class GlinkThread(QThread):
    """Glink操作线程"""
    login_result = pyqtSignal(bool, str)
    send_result = pyqtSignal(bool, str)

    def __init__(self, glink_instance):
        super().__init__()
        self.glink = glink_instance
        self.message = ""
        self.target_user = ""

    def set_task(self, task_type, **kwargs):
        """设置任务类型和参数"""
        self.task_type = task_type
        self.kwargs = kwargs

    def run(self):
        """线程执行函数"""
        try:
            if self.task_type == "login":
                self.glink.get_csrf_token()
                self.glink.login()
                success, message = self.glink.getCurrentUserInfo()
                self.login_result.emit(success, message)

            elif self.task_type == "send_message":
                self.message = self.kwargs.get("message", "")
                self.target_user = self.kwargs.get("target_user", "")

                if not self.glink.set_target_user(self.target_user):
                    self.send_result.emit(False, f"未找到用户: {self.target_user}")
                    return

                success, message = self.glink.send_message(self.message)
                self.send_result.emit(success, message)

        except Exception as e:
            if self.task_type == "login":
                self.login_result.emit(False, f"登录异常: {str(e)}")
            elif self.task_type == "send_message":
                self.send_result.emit(False, f"发送异常: {str(e)}")


# 唯一标识，用于单实例检查和通信
APP_UNIQUE_NAME = "MetroSchedule_Line5_8B3C7D2E"


class MetroReportCrawler:
    """广州地铁报表数据爬取工具"""

    BASE_URL = "http://frpt.gzmetro.com/webroot/decision/view/report"

    def __init__(self):
        self.session = requests.Session()
        self.headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/138.0.0.0 Safari/537.36",
        }
        self.timestamp = str(self._get_millis_timestamp())

    def _get_millis_timestamp(self) -> int:
        """获取当前毫秒级时间戳

        Returns:
            int: 13位毫秒级时间戳
        """
        return time.time_ns() // 1_000_000  # 纳秒转毫秒

    def _fetch_session_id(self) -> None:
        """获取会话ID并更新请求头"""
        response = self.session.post(
            f"{self.BASE_URL}?viewlet=02_LMIS/[751f][4ea7][6548][7387][5206][6790]/"
            "[8fd0][8425][524d][68c0][67e5][62a5][8868].cpt",
            headers=self.headers,
        )

        # 使用正则表达式从JavaScript代码中提取sessionID
        session_id_pattern = (
            r"this\.currentSessionID\s*=\s*'([a-f0-9]{8}-[a-f0-9]{4}-"
            r"[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})'"
        )
        match = re.search(session_id_pattern, response.text)
        self.headers['sessionID'] = match.group(1)

    def get_previous_dates(self):
        # 获取当前日期
        today = datetime.now().date()

        # 计算昨天和前天
        yesterday = today - timedelta(days=1)
        day_before_yesterday = today - timedelta(days=2)

        return yesterday.strftime('%Y-%m-%d'), day_before_yesterday.strftime('%Y-%m-%d')

    def _submit_query_parameters(self, date) -> None:
        """提交查询参数"""
        query_data = {
            "__parameters__": '{"LABEL1":"(6:00-[6b21][65e5]5:59,[6309][4f5c][4e1a][65e5][671f])",'
                              '"LABEL0":"[8fd0][8425][524d][68c0][67e5][60c5][51b5][6c47][603b][8868]",'
                              '"BEGINDATE":"' + date + '",'
                                                       '"LABELBEGINDATE":"[5386][53f2][6570][636e][67e5][8be2][ff1a]",'
                                                       '"LINE":"[4e94][53f7][7ebf]",'
                                                       '"LABELLINE":"[7ebf][8def][ff1a]"}',
            "_": self.timestamp
        }

        self.session.post(
            f"{self.BASE_URL}?op=fr_dialog&cmd=parameters_d",
            data=query_data,
            headers=self.headers
        )

    def fetch_report_data(self) -> str:
        """获取报表数据

        Returns:
            str: 报表HTML内容
        """
        response = self.session.post(
            f"{self.BASE_URL}?_={self.timestamp}"
            "&__boxModel__=true&op=page_content&pn=1"
            "&__webpage__=true&_paperWidth=477"
            "&_paperHeight=780&__fit__=false",
            headers=self.headers
        )
        return response.text

    def execute(self) -> tuple:
        """执行完整的爬取流程"""
        datas = self.get_previous_dates()
        for day in datas:
            self._fetch_session_id()
            self._submit_query_parameters(day)
            report_data = self.fetch_report_data()
            result = json.loads(report_data)
            soup = BeautifulSoup(result['html'], "lxml")
            tr = soup.find('tr', id='r-51-0')
            if tr:
                tds = tr.find_all('td')
                if len(tds) >= 8:
                    timetable = tds[8].text.strip()
                    if timetable:
                        # 解析日期
                        date_obj = datetime.strptime(day, "%Y-%m-%d")
                        # 加一天
                        next_day = date_obj + timedelta(days=1)
                        today = next_day.strftime("%Y-%m-%d")
                        return today, timetable
        return None, None


class CrawlerThread(QThread):
    """爬取线程，用于在后台执行爬取任务"""
    finished = pyqtSignal(tuple)  # 信号：爬取完成，返回结果
    error = pyqtSignal(str)  # 信号：爬取错误，返回错误信息

    def run(self):
        """线程执行函数"""
        try:
            crawler = MetroReportCrawler()
            result = crawler.execute()
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# 新增：文件搜索线程
class FileSearchThread(QThread):
    """文件搜索线程，用于在后台搜索匹配的Excel文件"""
    found = pyqtSignal(str)  # 找到文件
    not_found = pyqtSignal()  # 未找到文件
    error = pyqtSignal(str)  # 错误信息

    def __init__(self, search_text, search_path, search_subdirs=True):
        super().__init__()
        self.search_text = search_text
        self.search_path = search_path
        self.search_subdirs = search_subdirs
        self.running = True

    def run(self):
        """搜索匹配的Excel文件"""
        try:
            # 检查路径是否存在
            if not os.path.exists(self.search_path):
                self.error.emit(f"路径不存在: {self.search_path}")
                return

            # 检查是否有-或_
            has_separator = '-' in self.search_text or '_' in self.search_text

            # 搜索文件
            for root, dirs, files in os.walk(self.search_path):
                if not self.running:  # 检查是否需要停止
                    return

                for file in files:
                    if not self.running:  # 检查是否需要停止
                        return

                    # 只处理Excel文件
                    if file.lower().endswith(('.xlsx', '.xls')):
                        filename = os.path.splitext(file)[0]

                        if has_separator:
                            # 分割搜索文本
                            separators = []
                            if '-' in self.search_text:
                                separators.append('-')
                            if '_' in self.search_text:
                                separators.append('_')

                            # 尝试所有分隔符
                            matched = False
                            for sep in separators:
                                parts = self.search_text.split(sep)
                                if len(parts) == 2:
                                    left, right = parts
                                    # 检查文件名是否同时包含左右两部分
                                    if left in filename and right in filename:
                                        matched = True
                                        break

                            if matched:
                                self.found.emit(os.path.join(root, file))
                                return
                        else:
                            # 直接匹配文件名
                            if self.search_text in filename:
                                self.found.emit(os.path.join(root, file))
                                return

                # 如果不搜索子目录，只处理当前目录
                if not self.search_subdirs:
                    break

            # 如果没有找到文件
            self.not_found.emit()

        except Exception as e:
            self.error.emit(f"搜索错误: {str(e)}")

    def stop(self):
        """停止搜索线程"""
        self.running = False
        self.wait()


class SingleInstanceChecker:
    """单实例检查器，确保只有一个程序实例在运行"""

    def __init__(self, unique_name):
        self.unique_name = unique_name
        self.mutex_handle = None
        self.local_server = None
        self.show_window_callback = None

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


class TrainScheduleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # 判断是否在打包环境中运行
        if getattr(sys, 'frozen', False):
            # 如果是打包后的环境，使用临时目录中的图标
            base_path = sys._MEIPASS
        else:
            # 如果是普通 Python 环境，使用当前目录
            base_path = os.path.dirname(__file__)

        # 1. 初始化单实例检查器
        self.instance_checker = SingleInstanceChecker(APP_UNIQUE_NAME)

        # 检查是否已有实例运行
        if self.instance_checker.is_already_running():
            # 如果已有实例，发送显示窗口命令并退出当前实例
            self.instance_checker.send_show_command()
            sys.exit(0)

        icon_path = os.path.join(base_path, 'time.png')
        self.setWindowIcon(QIcon(icon_path))

        # 设置窗口标志：无边框，但在任务栏显示
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            # Qt.WindowType.WindowStaysOnTopHint |  # 窗口置顶
            Qt.WindowType.Window |  # 允许窗口在任务栏显示
            Qt.WindowType.X11BypassWindowManagerHint  # 完全绕过窗口管理器
        )

        # 车站列表
        self.STATION_LIST = [
            '滘口', '坦尾', '中山八', '西场', '西村', '广州火车站', '小北', '淘金', '区庄', '动物园',
            '杨箕', '五羊邨', '珠江新城', '猎德', '潭村', '员村', '科韵路', '车陂南', '东圃', '三溪',
            '鱼珠', '大沙地', '大沙东', '文冲', '双沙', '庙头', '夏园', '保盈大道', '夏港', '黄埔新港'
        ]
        self.selected_station = None  # 保存当前选择的车站
        self.processed_schedule = None  # 保存处理后的时刻表数据
        self.original_schedule = None  # 保存原始时刻表数据，用于重新加载

        # 记录修改历史
        self.modification_history = {
            'added': [],
            'modified': [],
            'deleted': []
        }

        # 统一字体设置
        self.BUTTON_FONT = QFont("Microsoft YaHei UI", 10)
        self.BUTTON_FONT.setBold(True)

        # 菜单字体设置
        self.MENU_FONT = QFont("Microsoft YaHei UI", 9)

        # 初始化固定状态：默认固定（显示完整窗口）
        self.pinned = True
        self.MAIN_WINDOW_WIDTH = 110
        self.MAIN_WINDOW_FIXED_HEIGHT = 440
        self.PIN_BUTTON_SIZE = (35, 35)  # 图钉按钮尺寸
        self.EDGE_SPACING = 10  # 窗口与屏幕边缘的间距
        self.HIDDEN_EDGE_SPACING = 0  # 隐藏状态下与边缘的间距
        self.PIN_ICON_FONT_SIZE = 16  # 图标字体大小
        self.HIDDEN_Y_OFFSET = -150  # 隐藏状态下向上偏移的距离

        # 添加微小的额外空间，确保图标完整显示
        self.HIDDEN_WINDOW_PADDING = 2

        # 配置全局提示框字体
        self.setup_tooltip_style()

        # 初始化提醒功能相关变量
        self.reminders = []  # 存储所有提醒
        self.reminder_timer = QTimer()
        self.reminder_timer.setInterval(1000)  # 每秒检查一次提醒
        self.reminder_timer.timeout.connect(self.check_reminders)
        self.reminder_timer.start()

        # 新增：共享文件夹路径
        self.SHARED_PATH = r"\\10.106.115.200\淘金中心站新共享\01 安全模块\【1】中心站安全文件盒（A类）\A1 行车组织（仅电子版）\02 运营时刻表\01 五号线"

        # 新增：当前爬取的时刻表文本
        self.current_timetable_text = ""

        self.init_ui()

        # 设置默认选中"淘金"车站
        default_station = "淘金"
        if default_station in self.STATION_LIST:
            self._on_station_selected(default_station)

        # 2. 启动本地服务器
        self.instance_checker.setup_local_server(self.show_normal)

    def show(self):
        """重写show方法，仅保留窗口显示逻辑"""
        super().show()

    def show_normal(self):
        """优化窗口显示逻辑"""
        # 恢复窗口状态：如果处于隐藏状态，先切换为固定状态
        if not self.pinned:
            self.toggle_pin()
        # 确保窗口正常显示（非最小化）
        self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.show()
        self.activateWindow()
        self.raise_()
        # 重置窗口位置到右侧居中
        self.move(self.get_fixed_window_pos())

    def setup_tooltip_style(self):
        """设置全局提示框的样式"""
        QToolTip.setFont(QFont("Microsoft YaHei UI", 9))
        self.setStyleSheet("""
            QToolTip {
                background-color: #f0f0f0;
                color: #333333;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 4px 8px;
                font-weight: normal;
            }
            QMainWindow {
                background-color: #f8f9fa;
            }
        """)

    def init_ui(self):
        # 设置窗口初始属性
        self.setWindowTitle("地铁时刻表查询")
        # 进一步调整窗口高度为300，使布局更紧凑
        self.setFixedSize(self.MAIN_WINDOW_WIDTH, 440)  # 增加高度，解决退出按钮显示不全问题

        # 将窗口初始位置设置在右侧居中
        self.move(self.get_fixed_window_pos())

        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局 - 调整边距使内容居中
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setContentsMargins(5, 5, 5, 5)
        self.main_layout.setSpacing(3)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        # 创建控制按钮容器
        self.control_buttons_layout = QHBoxLayout()
        self.control_buttons_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.main_layout.addLayout(self.control_buttons_layout)

        # 图钉按钮
        self.pin_button = QPushButton()
        self.pin_button.setFixedSize(*self.PIN_BUTTON_SIZE)

        # 设置图钉按钮样式
        font = self.pin_button.font()
        font.setPointSize(self.PIN_ICON_FONT_SIZE)
        self.pin_button.setFont(font)
        self.pin_button.setStyleSheet("""
            QPushButton { 
                border: none; 
                padding: 0px; 
                background-color: transparent;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
                border-radius: 3px;
            }
        """)

        # 初始化按钮状态
        self.pin_button.setText("🦒")  # 长颈鹿图标表示已固定
        self.pin_button.setToolTip("点击隐藏窗口")

        # 绑定图钉按钮事件
        self.pin_button.clicked.connect(self.toggle_pin)

        # 添加控制按钮到布局
        self.control_buttons_layout.addWidget(self.pin_button)

        # 创建按钮布局
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(3)
        buttons_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        # 车站选择按钮
        self.btn_station = QPushButton("选择车站")
        self.setup_station_button()

        # 创建车站选择菜单
        self.station_menu = QMenu(self)
        self.btn_station.setMenu(self.station_menu)
        self.station_menu.aboutToShow.connect(self._refresh_station_menu)
        self.station_menu.aboutToShow.connect(self._position_station_menu)

        buttons_layout.addWidget(self.btn_station)

        # 功能按钮2：时刻表查询（绑定点击事件）
        self.btn_func2 = QPushButton("时刻表查询")
        self.setup_button(self.btn_func2)
        self.btn_func2.clicked.connect(self.schedule_query)
        buttons_layout.addWidget(self.btn_func2)

        # 功能按钮3：时刻表导入
        self.btn_func3 = QPushButton("时刻表导入")
        self.setup_button(self.btn_func3)
        self.btn_func3.clicked.connect(self.import_timetable)
        buttons_layout.addWidget(self.btn_func3)

        # 功能按钮4：车次查询
        self.btn_func4 = QPushButton("车次查询")
        self.setup_button(self.btn_func4)
        self.btn_func4.clicked.connect(self.query_train_schedule)
        buttons_layout.addWidget(self.btn_func4)

        # 功能按钮5：时刻表获取
        self.btn_func6 = QPushButton("时刻表获取")
        self.setup_button(self.btn_func6)
        self.btn_func6.clicked.connect(self.get_timetable_from_web)
        buttons_layout.addWidget(self.btn_func6)

        # 功能按钮6：已设置提醒（新增）
        self.btn_func5 = QPushButton("已设置提醒")
        self.setup_button(self.btn_func5)
        self.btn_func5.clicked.connect(self.show_reminders_panel)
        buttons_layout.addWidget(self.btn_func5)



        # 放在“退出”按钮之前，看起来整齐
        self.btn_about = QPushButton("关于")
        self.setup_button(self.btn_about)
        self.btn_about.clicked.connect(self.show_about)
        buttons_layout.addWidget(self.btn_about)

        # 退出按钮
        self.btn_exit = QPushButton("退出")
        self.setup_button(self.btn_exit)
        self.btn_exit.clicked.connect(self.quit_application)  # 修改为调用quit_application方法
        buttons_layout.addWidget(self.btn_exit)

        # 存储所有功能按钮
        self.functional_buttons = [
            self.btn_station, self.btn_func2, self.btn_func3,
            self.btn_func4, self.btn_func5, self.btn_func6,
            self.btn_about,  # <── 新增
            self.btn_exit
        ]

        # 将按钮布局添加到主布局
        self.main_layout.addLayout(buttons_layout)
    # -------------- 2. 新增槽函数 --------------
    def show_about(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("时刻表助手")
        msg.setText(
            "Ver: 202604\n"
            "Dev: 肥鹏"
        )
        msg.setIcon(QMessageBox.Icon.Information)

        # 确保大小已计算
        msg.adjustSize()

        # 屏幕正中央
        screen = QApplication.primaryScreen().availableGeometry().center()
        msg.move(screen.x() - msg.width() // 2,
                 screen.y() - msg.height() // 2)

        msg.exec()

    def setup_button(self, button):
        """设置普通按钮的样式和属性"""
        button.setFont(self.BUTTON_FONT)
        button.setFixedSize(90, 45)
        button.setContentsMargins(0, 0, 0, 0)
        button.setToolTip("")  # 初始不显示提示

        button.setStyleSheet("""
            QPushButton {
                color: #333333;
                background-color: transparent;
                border: 1px solid #dddddd;
                border-radius: 3px;
                padding: 2px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
                border-color: #bbbbbb;
            }
            QPushButton:pressed {
                background-color: #e0e0e0;
            }
        """)

    def setup_station_button(self):
        """专门设置车站按钮的样式，隐藏下拉箭头"""
        self.btn_station.setFont(self.BUTTON_FONT)
        self.btn_station.setFixedSize(90, 45)
        self.btn_station.setContentsMargins(0, 0, 0, 0)

        self.btn_station.setStyleSheet("""
            QPushButton {
                color: #2c3e50;
                background-color: #ecf0f1;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 4px 8px;
                text-align: center;
            }
            QPushButton::menu-indicator {
                image: none; /* 隐藏下拉箭头 */
            }
            QPushButton:hover {
                background-color: #d5dbdb;
                border-color: #95a5a6;
            }
            QPushButton:pressed {
                background-color: #b3b6b7;
            }
        """)

    def _position_station_menu(self):
        """精确计算菜单位置"""
        if not self.station_menu:
            return

        button_rect = self.btn_station.geometry()
        global_pos = self.mapToGlobal(button_rect.topLeft())

        menu_width = 150
        self.station_menu.setFixedWidth(menu_width)

        self.station_menu.adjustSize()
        menu_height = self.station_menu.height()

        x = global_pos.x() - menu_width - 1
        y = global_pos.y()

        screen_geometry = QApplication.primaryScreen().availableGeometry()

        if x < screen_geometry.left():
            x = screen_geometry.left()
        if y < screen_geometry.top():
            y = screen_geometry.top()
        if y + menu_height > screen_geometry.bottom():
            y = screen_geometry.bottom() - menu_height

        self.station_menu.move(x, y)

    def _refresh_station_menu(self):
        """刷新车站下拉菜单"""
        self.station_menu.clear()
        self.station_menu.setFont(self.MENU_FONT)

        if not self.STATION_LIST:
            no_station_action = QAction("暂无车站数据", self)
            no_station_action.setEnabled(False)
            self.station_menu.addAction(no_station_action)
            return

        for station in self.STATION_LIST:
            station_action = QAction(station, self)
            station_action.setFont(self.MENU_FONT)
            station_action.triggered.connect(
                lambda checked, s=station: self._on_station_selected(s)
            )
            self.station_menu.addAction(station_action)

    def _on_station_selected(self, station):
        """处理车站选择事件"""
        self.selected_station = station
        self.btn_station.setText(station)
        self.btn_station.setStyleSheet(f"""
            QPushButton {{
                color: #ff0000;
                background-color: #ecf0f1;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 4px 8px;
                text-align: center;
                font-family: "{self.BUTTON_FONT.family()}";
                font-size: {self.BUTTON_FONT.pointSize()}pt;
                font-weight: {'bold' if self.BUTTON_FONT.bold() else 'normal'};
            }}
            QPushButton::menu-indicator {{
                image: none;
            }}
            QPushButton:hover {{
                background-color: #d5dbdb;
                border-color: #95a5a6;
            }}
            QPushButton:pressed {{
                background-color: #b3b6b7;
            }}
        """)

    def get_fixed_window_pos(self, window_width=None, window_height=None, is_hidden=False):
        if window_width is None:
            window_width = self.MAIN_WINDOW_WIDTH
        if window_height is None:
            window_height = self.MAIN_WINDOW_FIXED_HEIGHT

        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_width = screen_geometry.width()
        screen_height = screen_geometry.height()

        spacing = self.HIDDEN_EDGE_SPACING if is_hidden else self.EDGE_SPACING
        x = screen_width - window_width - spacing

        if is_hidden:
            y = (screen_height - window_height) // 2 + self.HIDDEN_Y_OFFSET
        else:
            y = (screen_height - window_height) // 2
        return QPoint(x, y)

    def toggle_pin(self):
        self.pinned = not self.pinned
        if self.pinned:
            # 显示窗口和任务栏图标
            self.setWindowFlags(
                Qt.WindowType.FramelessWindowHint |
                Qt.WindowType.Window |
                Qt.WindowType.X11BypassWindowManagerHint
            )
            self.setFixedSize(self.MAIN_WINDOW_WIDTH, 440)
            self.show()
            self.pin_button.setText("🦒")
            self.pin_button.setToolTip("点击隐藏窗口")
            for btn in self.functional_buttons:
                btn.show()
            self.move(self.get_fixed_window_pos())
        else:
            # 隐藏任务栏图标，但保持窗口可见
            self.setWindowFlags(
                Qt.WindowType.FramelessWindowHint |
                Qt.WindowType.Tool |  # 防止窗口出现在任务栏
                Qt.WindowType.X11BypassWindowManagerHint
            )
            self.show()  # 重新显示以应用新的窗口标志

            window_width = self.PIN_BUTTON_SIZE[0] + self.HIDDEN_WINDOW_PADDING
            window_height = self.PIN_BUTTON_SIZE[1] + self.HIDDEN_WINDOW_PADDING
            self.setFixedSize(window_width, window_height)

            self.pin_button.setText("🐼")
            self.pin_button.setToolTip("点击显示窗口")
            for btn in self.functional_buttons:
                btn.hide()
            self.move(self.get_fixed_window_pos(
                window_width=window_width,
                window_height=window_height,
                is_hidden=True
            ))

    def show_center_message(self, title, message, icon=QMessageBox.Icon.Information):
        """显示一个在屏幕中心的提示框"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(icon)

        msg_box.adjustSize()

        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_center = screen_geometry.center()

        msg_size = msg_box.geometry()

        center_x = screen_center.x() - msg_size.width() // 2
        center_y = screen_center.y() - msg_size.height() // 2

        msg_box.move(center_x, center_y)
        msg_box.exec()

    def show_center_question(self, title, message):
        """显示一个在屏幕中心的确认对话框"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(QMessageBox.Icon.Question)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.No)

        msg_box.adjustSize()

        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_center = screen_geometry.center()

        msg_size = msg_box.geometry()

        center_x = screen_center.x() - msg_size.width() // 2
        center_y = screen_center.y() - msg_size.height() // 2

        msg_box.move(center_x, center_y)
        return msg_box.exec()

    # 显示时刻表信息的对话框
    def show_schedule_dialog(self, station, schedule_data):
        """显示指定车站的时刻表信息，分左右两栏显示上下行信息"""
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{station}站 时刻表")
        dialog.setFixedSize(600, 300)  # 进一步调整窗口高度，减少留白

        # 设置对话框字体
        font = QFont("Microsoft YaHei UI", 10)
        dialog.setFont(font)

        # 主布局
        main_layout = QVBoxLayout(dialog)
        main_layout.setContentsMargins(5, 5, 5, 0)  # 设置边距，底部不留间隔
        main_layout.setSpacing(2)  # 减小上下行方向与按钮之间的间距

        # 创建分割器用于左右布局
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # 上行方向列表
        upbound_frame = QFrame()
        upbound_frame.setFrameShape(QFrame.Shape.StyledPanel)
        upbound_layout = QVBoxLayout(upbound_frame)
        upbound_layout.setContentsMargins(2, 2, 2, 2)  # 减少内边距
        upbound_layout.setSpacing(2)  # 减少间距

        upbound_label = QLabel("【上行方向】")
        upbound_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        upbound_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #2c3e50; margin: 0px;")
        upbound_layout.addWidget(upbound_label)

        self.upbound_list = QListWidget()
        # 修改上行方向列表的样式表，将选中项背景色改为柔和的黄色
        self.upbound_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 0px;
            }
            QListWidget::item {
                padding: 4px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background-color: #ffe066;  /* 较深的黄色，不刺眼 */
                color: #333333;             /* 深灰色文字，提高对比度 */
            }
        """)
        upbound_layout.addWidget(self.upbound_list)
        splitter.addWidget(upbound_frame)

        # 下行方向列表
        downbound_frame = QFrame()
        downbound_frame.setFrameShape(QFrame.Shape.StyledPanel)
        downbound_layout = QVBoxLayout(downbound_frame)
        downbound_layout.setContentsMargins(2, 2, 2, 2)  # 减少内边距
        downbound_layout.setSpacing(2)  # 减少间距

        downbound_label = QLabel("【下行方向】")
        downbound_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        downbound_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #2c3e50; margin: 0px;")
        downbound_layout.addWidget(downbound_label)

        self.downbound_list = QListWidget()
        self.downbound_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 0px;
            }
            QListWidget::item {
                padding: 4px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
        """)
        downbound_layout.addWidget(self.downbound_list)
        splitter.addWidget(downbound_frame)

        # 添加更新按钮布局
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 0, 0, 0)  # 减少边距，底部不留间隔
        button_layout.setSpacing(5)  # 减少间距，为新按钮留出空间

        # 创建实时更新按钮
        self.update_button = QPushButton("开始自动更新")
        self.update_button.setFixedSize(120, 36)  # 稍微减小宽度，为新按钮留出空间

        # 创建一键添加所有不停站提醒按钮
        self.add_all_reminders_button = QPushButton("批量加不停站")
        self.add_all_reminders_button.setFixedSize(120, 36)  # 设置合适的尺寸

        # 添加当前时间显示标签
        self.time_label = QLabel()
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.time_label.setStyleSheet("font-size: 16px; color: #666666; margin: 0px;")

        # 添加到布局
        button_layout.addWidget(self.update_button)
        button_layout.addWidget(self.add_all_reminders_button)  # 添加新按钮
        button_layout.addStretch()
        button_layout.addWidget(self.time_label)

        main_layout.addLayout(button_layout)

        # 初始化自动更新状态
        self.auto_update_enabled = False
        self.update_timer = QTimer()
        self.update_timer.setInterval(1000)  # 每秒更新一次

        # 填充数据
        self._populate_schedule_lists(station, schedule_data)

        # 更新当前时间显示
        self._update_current_time()

        # 绑定按钮点击事件
        def toggle_auto_update():
            self.auto_update_enabled = not self.auto_update_enabled

            if self.auto_update_enabled:
                self.update_button.setText("停止自动更新")
                self.update_button.setProperty("autoUpdate", "true")
                self.update_timer.timeout.connect(
                    lambda: self._populate_schedule_lists(self.selected_station, self.processed_schedule))
                self.update_timer.start()
                # 自动更新开启时禁用一键添加按钮
                self.add_all_reminders_button.setEnabled(False)
            else:
                self.update_button.setText("开始自动更新")
                self.update_button.setProperty("autoUpdate", "false")
                self.update_timer.timeout.disconnect()
                self.update_timer.stop()
                # 自动更新关闭时启用一键添加按钮
                self.add_all_reminders_button.setEnabled(True)

            # 更新按钮样式
            self.update_button.style().unpolish(self.update_button)
            self.update_button.style().polish(self.update_button)

        self.update_button.clicked.connect(toggle_auto_update)

        # 绑定一键添加所有不停站提醒按钮点击事件
        def add_all_non_stop_reminders():
            # 检查是否有自动更新在运行
            if hasattr(self, 'auto_update_enabled') and self.auto_update_enabled:
                self.show_center_message(
                    "操作提示",
                    "请先停止自动更新功能，再执行此操作",
                    QMessageBox.Icon.Warning
                )
                return

            # 调用添加所有不停站提醒的方法
            self._add_all_non_stop_reminders(station, schedule_data)

        self.add_all_reminders_button.clicked.connect(add_all_non_stop_reminders)

        # 对话框关闭时停止定时器
        def on_dialog_closed():
            if self.update_timer.isActive():
                self.update_timer.stop()

        # 初始化时启用一键添加按钮
        self.add_all_reminders_button.setEnabled(True)

        dialog.finished.connect(on_dialog_closed)

        # 为列表添加右键菜单
        self.upbound_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.upbound_list.customContextMenuRequested.connect(
            lambda pos: self._show_context_menu(self.upbound_list, pos, '上行'))

        self.downbound_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.downbound_list.customContextMenuRequested.connect(
            lambda pos: self._show_context_menu(self.downbound_list, pos, '下行'))

        # 居中显示对话框
        dialog.adjustSize()
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_center = screen_geometry.center()
        dialog.move(
            screen_center.x() - dialog.width() // 2,
            screen_center.y() - dialog.height() // 2
        )

        # 保存对话框引用，以便在设置提醒时使用
        self.schedule_dialog = dialog

        dialog.exec()

    def _show_context_menu(self, list_widget, pos, direction):
        """显示右键菜单"""

        # 获取当前选中的项
        selected_items = list_widget.selectedItems()
        if not selected_items:
            return

        selected_item = selected_items[0]
        index = list_widget.row(selected_item)

        # 获取车次时间
        current_text = selected_item.text()

        # 修复点：使用更宽松的正则表达式，允许时间部分为空
        match = re.match(r'车次: (\d+)(?:  时间: (.+))?', current_text)
        if not match:
            return

        current_train_id = match.group(1)  # 这已经是 fmt_train_id 格式（3-5位）
        current_time = match.group(2) if match.group(2) else ""

        # 创建右键菜单
        context_menu = QMenu()
        context_menu.setFont(self.MENU_FONT)

        # 添加菜单项
        add_above_action = QAction("插入车次(上方)", self)
        add_below_action = QAction("插入车次(下方)", self)
        modify_action = QAction("修改车次", self)
        delete_action = QAction("删除车次", self)
        reload_action = QAction("重新加载", self)

        # ========== 修复：检查是否已设置任何类型的提醒（包括不停站和特殊乘客） ==========
        now = datetime.now()

        # 检查是否已设置特殊乘客提醒
        has_special_passenger = any(
            r['active'] and
            r['type'] == 'special_passenger' and
            r['train_id'] == current_train_id and
            r['reminder_time'] >= now
            for r in self.reminders
        )

        # 检查是否已设置不停站提醒（新增）
        has_non_stop = any(
            r['active'] and
            r['type'] == 'non_stop' and
            r['train_id'] == current_train_id and
            r['reminder_time'] >= now
            for r in self.reminders
        )

        # 合并判断：只要有任一类型的有效提醒，就视为"已设置提醒"
        has_any_reminder = has_special_passenger or has_non_stop
        # ============================================================================

        # 构造菜单项
        modify_action = QAction("修改车次", self)
        delete_action = QAction("删除车次", self)

        reminder_action = None
        if current_time and current_time.strip():
            if has_any_reminder:  # ← 修改：使用合并后的判断
                # 灰体不可选
                reminder_action = QAction("已设置提醒", self)
                reminder_action.setEnabled(False)
                modify_action.setEnabled(False)
            else:
                # 原来逻辑不变
                try:
                    datetime.strptime(current_time.replace("\t文", "").strip(), "%H:%M:%S")
                    reminder_action = QAction("设置提醒", self)
                except ValueError:
                    if "文" in current_time:
                        reminder_action = QAction("设置提醒", self)
                    else:
                        reminder_action = QAction("设置不停站提醒", self)

        # 绑定事件
        add_above_action.triggered.connect(lambda: self._add_train(list_widget, index, direction, above=True))
        add_below_action.triggered.connect(lambda: self._add_train(list_widget, index, direction, above=False))
        modify_action.triggered.connect(lambda: self._modify_train(list_widget, selected_item, index, direction))
        delete_action.triggered.connect(lambda: self._delete_train(list_widget, selected_item, index, direction))
        reload_action.triggered.connect(self._reload_schedule)

        # 添加到菜单
        context_menu.addAction(add_above_action)
        context_menu.addAction(add_below_action)
        context_menu.addSeparator()
        context_menu.addAction(modify_action)
        context_menu.addAction(delete_action)

        # 添加查询全站点停站时间菜单项
        if current_time and current_time.strip():
            context_menu.addSeparator()
            query_all_stations_action = QAction("全线各站时间", self)
            query_all_stations_action.triggered.connect(
                lambda: self._query_all_stations_stop_time(current_train_id, direction))
            context_menu.addAction(query_all_stations_action)

        # 只有当时间非空时才添加提醒菜单项（修复点1）
        if reminder_action:
            context_menu.addSeparator()
            reminder_action.triggered.connect(
                lambda: self._set_reminder(list_widget, selected_item, index, direction))  # 绑定提醒功能
            context_menu.addAction(reminder_action)  # 添加提醒菜单项

        # 添加Glink发送消息菜单项
        context_menu.addSeparator()
        glink_action = QAction("Glink发送消息", self)
        glink_action.triggered.connect(
            lambda: self._send_message_via_glink(list_widget, selected_item, index, direction))
        context_menu.addAction(glink_action)

        context_menu.addSeparator()
        context_menu.addAction(reload_action)

                # 在鼠标位置显示菜单
        global_pos = list_widget.mapToGlobal(pos)
        context_menu.exec(global_pos)

    def _check_auto_update_status(self):
        """检查自动更新状态，如果正在运行则显示提示并返回True"""
        if hasattr(self, 'auto_update_enabled') and self.auto_update_enabled:
            self.show_center_message(
                "操作提示",
                "请先停止自动更新功能，再执行此操作",
                QMessageBox.Icon.Warning
            )
            return True
        return False

    def _add_train(self, list_widget, index, direction, above=True):
        """新增车次"""
        # 检查自动更新状态
        if self._check_auto_update_status():
            return

        # 获取当前选中的车次信息作为参考
        current_item = list_widget.item(index)
        if not current_item:
            return

        current_text = current_item.text()
        match = re.match(r'车次: (\d+)  时间: (.+)', current_text)
        if not match:
            self.show_center_message("错误", "无法解析当前车次信息", QMessageBox.Icon.Critical)
            return

        current_train_id = match.group(1)
        current_time = match.group(2)

        # 创建输入对话框
        dialog = QDialog()
        dialog.setWindowTitle("新增车次")
        dialog.setFixedSize(250, 150)

        layout = QVBoxLayout(dialog)

        # 车次号输入
        train_id_layout = QHBoxLayout()
        train_id_label = QLabel("车次号:")
        self.train_id_input = QLineEdit()
        self.train_id_input.setText(current_train_id)
        train_id_layout.addWidget(train_id_label)
        train_id_layout.addWidget(self.train_id_input)

        # 时间输入
        time_layout = QHBoxLayout()
        time_label = QLabel("时间:")
        self.time_input = QLineEdit()
        self.time_input.setText(current_time)
        time_layout.addWidget(time_label)
        time_layout.addWidget(self.time_input)

        # 确认按钮
        confirm_btn = QPushButton("确认")
        confirm_btn.clicked.connect(dialog.accept)

        # 添加到布局
        layout.addLayout(train_id_layout)
        layout.addLayout(time_layout)
        layout.addWidget(confirm_btn)

        # 显示对话框
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_train_id = self.train_id_input.text().strip()
            new_time = self.time_input.text().strip()

            # 验证输入
            if not new_train_id.isdigit() or len(new_train_id) < 3 or len(new_train_id) > 5:
                self.show_center_message("错误", "请输入3-5位数字车次号", QMessageBox.Icon.Critical)
                return

            # 验证时间不能为空
            if not new_time.strip():
                self.show_center_message("错误", "时间不能为空", QMessageBox.Icon.Critical)
                return

            # 创建新的车次数据 - 统一格式
            full_id = f"0{new_train_id}"
            new_train_data = {
                full_id: new_time,
                '__display__': new_train_id,  # 保存用户输入的原始格式
                '__source__': 'manual'
            }

            # 插入到数据结构中
            insert_index = index if above else index + 1
            self.processed_schedule[self.selected_station][direction].insert(insert_index, new_train_data)

            # 记录修改历史
            self.modification_history['added'].append({
                'station': self.selected_station,
                'direction': direction,
                'index': insert_index,
                'train_data': new_train_data
            })

            # 重新填充列表
            self._populate_schedule_lists(self.selected_station, self.processed_schedule)

            # 显示成功信息
            self.show_center_message("成功", "车次添加成功", QMessageBox.Icon.Information)

    def _modify_train(self, list_widget, item, index, direction):
        """修改车次"""
        # 检查自动更新状态
        if self._check_auto_update_status():
            return

        # 获取当前车次信息
        current_text = item.text()
        match = re.match(r'车次: (\d+)  时间: (.+)', current_text)
        if not match:
            self.show_center_message("错误", "无法解析当前车次信息", QMessageBox.Icon.Critical)
            return

        current_train_id = match.group(1)
        current_time = match.group(2)

        # 创建输入对话框
        dialog = QDialog()
        dialog.setWindowTitle("修改车次")
        dialog.setFixedSize(250, 150)

        layout = QVBoxLayout(dialog)

        # 车次号输入
        train_id_layout = QHBoxLayout()
        train_id_label = QLabel("车次号:")
        self.train_id_input = QLineEdit()
        self.train_id_input.setText(current_train_id)
        train_id_layout.addWidget(train_id_label)
        train_id_layout.addWidget(self.train_id_input)

        # 时间输入
        time_layout = QHBoxLayout()
        time_label = QLabel("时间:")
        self.time_input = QLineEdit()
        self.time_input.setText(current_time)
        time_layout.addWidget(time_label)
        time_layout.addWidget(self.time_input)

        # 确认按钮
        confirm_btn = QPushButton("确认")
        confirm_btn.clicked.connect(dialog.accept)

        # 添加到布局
        layout.addLayout(train_id_layout)
        layout.addLayout(time_layout)
        layout.addWidget(confirm_btn)

        # 显示对话框
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_train_id = self.train_id_input.text().strip()
            new_time = self.time_input.text().strip()

            # 验证输入
            if not new_train_id.isdigit() or len(new_train_id) < 3 or len(new_train_id) > 5:
                self.show_center_message("错误", "请输入3-5位数字车次号", QMessageBox.Icon.Critical)
                return

            # 验证时间不能为空
            if not new_time.strip():
                self.show_center_message("错误", "时间不能为空", QMessageBox.Icon.Critical)
                return

            # 保存原始数据用于历史记录
            original_train_data = self.processed_schedule[self.selected_station][direction][index]

            # 更新数据结构 - 统一格式
            full_id = f"0{new_train_id}"
            new_train_data = {
                full_id: new_time,
                '__display__': new_train_id,  # 保存用户输入的原始格式
                '__source__': 'manual'
            }
            self.processed_schedule[self.selected_station][direction][index] = new_train_data

            # 记录修改历史
            self.modification_history['modified'].append({
                'station': self.selected_station,
                'direction': direction,
                'index': index,
                'original_train_data': original_train_data,
                'new_train_data': new_train_data
            })

            # 重新填充列表
            self._populate_schedule_lists(self.selected_station, self.processed_schedule)

            # 显示成功信息
            self.show_center_message("成功", "车次修改成功", QMessageBox.Icon.Information)
    def _delete_train(self, list_widget, item, index, direction):
        """删除车次"""
        # 检查自动更新状态
        if self._check_auto_update_status():
            return

        # 确认删除（居中显示）
        reply = self.show_center_question(
            "确认删除",
            "确定要删除选中的车次吗？"
        )

        if reply == QMessageBox.StandardButton.Yes:
            # 保存删除的数据用于历史记录
            deleted_train_data = self.processed_schedule[self.selected_station][direction][index]

            # 从数据结构中删除
            del self.processed_schedule[self.selected_station][direction][index]

            # 记录修改历史
            self.modification_history['deleted'].append({
                'station': self.selected_station,
                'direction': direction,
                'index': index,
                'train_data': deleted_train_data
            })

            # 重新填充列表
            self._populate_schedule_lists(self.selected_station, self.processed_schedule)

            # 显示成功信息
            self.show_center_message("成功", "车次删除成功", QMessageBox.Icon.Information)

    def _reload_schedule(self):
        """重新加载原始数据"""
        # 检查自动更新状态
        if self._check_auto_update_status():
            return

        # 确认重新加载（居中显示）
        reply = self.show_center_question(
            "确认重新加载",
            "确定要重新加载原始数据吗？所有修改将被放弃。"
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.original_schedule:
                # 恢复原始数据
                self.processed_schedule = json.loads(json.dumps(self.original_schedule))

                # 清空修改历史
                self.modification_history = {
                    'added': [],
                    'modified': [],
                    'deleted': []
                }

                # 重新填充列表
                self._populate_schedule_lists(self.selected_station, self.processed_schedule)

                # 显示成功信息
                self.show_center_message("成功", "已重新加载原始数据", QMessageBox.Icon.Information)
            else:
                self.show_center_message("提示", "没有原始数据可加载", QMessageBox.Icon.Information)

    def _update_current_time(self):
        """更新当前时间显示"""
        if hasattr(self, 'time_label'):
            current_time = datetime.now().strftime("%H:%M:%S")
            self.time_label.setText(f"当前时间: {current_time}")

    def _populate_schedule_lists(self, station, schedule_data):
        """填充上下行车次列表并根据当前时间定位"""
        current_time = datetime.now().time()

        # 收集"生效且未过期"的特殊乘客提醒（使用 fmt_train_id 统一格式）
        now = datetime.now()
        special_train_info = {
            fmt_train_id(r['train_id']): r['platform_gate']
            for r in self.reminders
            if r['active'] and r['type'] == 'special_passenger' and r['reminder_time'] >= now
        }
        special_train_ids = set(special_train_info.keys())

        # 更新当前时间显示
        self._update_current_time()

        # 清空列表
        self.upbound_list.clear()
        self.downbound_list.clear()

        # 上下行统一处理函数
        def fill_list(trains, list_widget, direction):
            scroll_pos = None
            for idx, train in enumerate(trains):
                # ========== 修改：解析统一格式 ==========
                if '__display__' in train:
                    # 统一格式：{full_id: time, '__display__': display_id, '__source__': source}
                    display_train_id = train['__display__']
                    is_manual = (train.get('__source__') == 'manual')
                    # 找到实际的车次ID和时间（非特殊键）
                    for key, val in train.items():
                        if not key.startswith('__'):
                            train_id = key
                            time_str = val
                            break
                else:
                    # 兼容旧格式（直接从Excel导入的原始数据）
                    train_id = list(train.keys())[0]
                    time_str = list(train.values())[0]
                    display_train_id = train_id[-4:].zfill(4)
                    is_manual = False
                # =======================================

                # 1. 解析时间对象（可能失败）
                time_obj = None
                try:
                    time_obj = datetime.strptime(time_str, "%H:%M:%S").time()
                except ValueError:
                    pass

                # 2. 黄埔新港时间判断（仅上行用）
                has_huangpu_xingang_time = False
                if direction == '上行' and time_obj and station != '黄埔新港':
                    huangpu_xingang_schedule = schedule_data.get('黄埔新港', {}).get('上行', [])
                    for hp_train in huangpu_xingang_schedule:
                        # 兼容新格式
                        if isinstance(hp_train, dict) and '__display__' in hp_train:
                            for key, val in hp_train.items():
                                if not key.startswith('__'):
                                    hp_train_id = key
                                    hp_time_str = val
                                    break
                        else:
                            hp_train_id = list(hp_train.keys())[0]
                            hp_time_str = list(hp_train.values())[0]

                        if hp_train_id == train_id:
                            try:
                                datetime.strptime(hp_time_str, "%H:%M:%S").time()
                                has_huangpu_xingang_time = True
                                break
                            except ValueError:
                                pass
                    # if has_huangpu_xingang_time:
                    #     break

                # 3. 判断是否新增/修改（用于红色标记）
                is_modified = is_manual

                # 4. 拼显示文本
                if direction == '上行':
                    if is_modified:
                        display_text = f"车次: {display_train_id}  时间: {time_str}"
                    else:
                        if time_obj and not has_huangpu_xingang_time:
                            display_text = f"车次: {display_train_id}  时间: {time_str}  文"
                        else:
                            display_text = f"车次: {display_train_id}  时间: {time_str}"
                else:  # 下行永不拼"文"
                    display_text = f"车次: {display_train_id}  时间: {time_str}"

                # 5. 追加站台门（提醒统一用 fmt_train_id 匹配）
                gate = special_train_info.get(fmt_train_id(train_id))
                if gate:
                    display_text += f"  {gate}"

                # 6. 创建 Item
                item = QListWidgetItem(display_text)
                item.setData(ItemRole_HasLiveSpecialReminder, fmt_train_id(train_id) in special_train_ids)

                # 7. 设颜色
                if fmt_train_id(train_id) in special_train_ids:
                    if direction == '上行':
                        item.setForeground(QColor(255, 165, 0))  # 上行橙色
                    else:
                        item.setForeground(QColor("#3498db"))  # 下行蓝色
                elif is_modified:
                    item.setForeground(QColor("red"))  # 新增/修改
                else:
                    if time_obj:
                        if (datetime.combine(datetime.min, time_obj) + timedelta(minutes=1)).time() < current_time:
                            item.setForeground(Qt.GlobalColor.gray)
                    else:
                        item.setForeground(Qt.GlobalColor.darkGray)
                        item.setToolTip("非标准时间格式")

                list_widget.addItem(item)

                # 8. 记录第一个未过期车次位置
                if scroll_pos is None and time_obj:
                    if (datetime.combine(datetime.min, time_obj) + timedelta(minutes=1)).time() >= current_time:
                        scroll_pos = list_widget.count() - 1

            return scroll_pos

        # 上行
        upbound_trains = schedule_data.get(station, {}).get('上行', [])
        up_scroll = fill_list(upbound_trains, self.upbound_list, '上行')
        if up_scroll is not None:
            self.upbound_list.scrollToItem(self.upbound_list.item(up_scroll), QListWidget.ScrollHint.PositionAtTop)
            self.upbound_list.item(up_scroll).setSelected(True)

        # 下行
        downbound_trains = schedule_data.get(station, {}).get('下行', [])
        down_scroll = fill_list(downbound_trains, self.downbound_list, '下行')
        if down_scroll is not None:
            self.downbound_list.scrollToItem(self.downbound_list.item(down_scroll),
                                             QListWidget.ScrollHint.PositionAtTop)
            self.downbound_list.item(down_scroll).setSelected(True)

    def _is_manual_modified(self, station, direction, index, train_id):
        """判断车次是否为手动添加或修改的"""
        # 检查 added 历史
        for add in self.modification_history['added']:
            if (add['station'] == station and
                    add['direction'] == direction and
                    add['index'] == index):
                return True

        # 检查 modified 历史
        for mod in self.modification_history['modified']:
            if (mod['station'] == station and
                    mod['direction'] == direction and
                    mod['index'] == index):
                # 确认是这个车次
                new_data = mod.get('new_train_data', {})
                if train_id in new_data:
                    return True

        return False

    def import_timetable(self):
        """时刻表导入入口"""
        shared_path = r"\\10.106.115.200\淘金中心站新共享\01 安全模块\【1】中心站安全文件盒（A类）\A1 行车组织（仅电子版）\02 运营时刻表\01 五号线"

        try:
            if os.path.exists(shared_path):
                file_path, _ = QFileDialog.getOpenFileName(
                    self,
                    "选择运营时刻表Excel文件",
                    shared_path,
                    "Excel Files (*.xlsx *.xls);;All Files (*.*)"
                )

                if file_path:
                    self._process_timetable_file(file_path)
                    return
                else:
                    # 用户取消了共享文件选择，直接返回，不打开本地文件
                    return
            else:
                self.show_center_message("警告", "共享路径无法访问，将打开本地文件选择", QMessageBox.Icon.Warning)

        except Exception as e:
            self.show_center_message("错误", f"访问共享路径失败：{str(e)}\n将打开本地文件选择",
                                     QMessageBox.Icon.Critical)

        # 只有当共享路径不存在或访问失败时，才会打开本地文件选择
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择运营时刻表Excel文件",
            os.getcwd(),
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )

        if file_path:
            self._process_timetable_file(file_path)

    def _process_timetable_file(self, file_path):
        """封装文件处理逻辑"""
        try:
            self.processed_schedule = self.process_timetable_data(file_path)
            # 保存原始数据用于重新加载
            self.original_schedule = json.loads(json.dumps(self.processed_schedule))

            # 清空修改历史
            self.modification_history = {
                'added': [],
                'modified': [],
                'deleted': []
            }

            # 获取文件名
            file_name = os.path.basename(file_path)

            # 使用正则表达式提取大写字母、数字、_和-的组合
            import re
            match = re.search(r'([A-Z0-9_-]+)', file_name)
            if match:
                display_name = match.group(1)
            else:
                # 如果没有匹配到，则使用不带扩展名的文件名
                display_name = os.path.splitext(file_name)[0]

            # 更新按钮文字和样式
            self.btn_func3.setText(display_name)
            self.btn_func3.setToolTip(file_name)  # 工具提示显示完整文件名
            self.btn_func3.setStyleSheet("""
                QPushButton {
                    color: #ff0000;
                    background-color: transparent;
                    border: 1px solid #dddddd;
                    border-radius: 3px;
                    padding: 2px;
                    text-align: center;
                }
                QPushButton:hover {
                    background-color: #f0f0f0;
                    border-color: #bbbbbb;
                }
                QPushButton:pressed {
                    background-color: #e0e0e0;
                }
            """)

            self.show_center_message(
                "成功",
                f"时刻表导入成功！"
            )

            # 如果已选择车站，直接用print输出该车站的时刻表数据，供后续车次查询使用
            if self.selected_station and self.selected_station in self.processed_schedule:
                print(f"\n===== {self.selected_station}站 时刻表数据（供车次查询使用） =====")
                print(f"上行方向数据: {self.processed_schedule[self.selected_station]['上行']}")
                print(f"下行方向数据: {self.processed_schedule[self.selected_station]['下行']}")
                print("===============================================\n")

        except Exception as e:
            self.show_center_message(
                "错误",
                f"请选择正确运营时刻表",
                QMessageBox.Icon.Critical
            )

    def process_timetable_data(self, file_path):
        """处理地铁运营时刻表Excel文件"""
        df = pd.read_excel(file_path)
        total_columns = len(df.columns)

        TRAIN_ID_MARKER = "TRAINID"

        marker_rows, marker_cols = (df == TRAIN_ID_MARKER).values.nonzero()
        if len(marker_rows) == 0 or len(marker_cols) == 0:
            raise ValueError("Excel文件中未找到'TRAINID'标记，无法识别车次ID列")

        STATION_LIST = [
            '滘口', '坦尾', '中山八', '西场', '西村', '广州火车站', '小北', '淘金', '区庄', '动物园',
            '杨箕', '五羊邨', '珠江新城', '猎德', '潭村', '员村', '科韵路', '车陂南', '东圃', '三溪',
            '鱼珠', '大沙地', '大沙东', '文冲', '双沙', '庙头', '夏园', '保盈大道', '夏港', '黄埔新港'
        ]

        station_schedule = {
            station: {'上行': [], '下行': []}
            for station in STATION_LIST
        }

        for current_station in STATION_LIST:
            station_column = marker_cols[0]
            station_rows = df.index[df.iloc[:, station_column] == current_station].tolist()

            if not station_rows:
                self.show_center_message("提示", f"Excel中未找到车站：{current_station}，将跳过该车站",
                                         QMessageBox.Icon.Warning)
                continue

            for station_index in range(len(station_rows)):
                if station_index * 2 + 1 >= len(marker_rows):
                    self.show_center_message("提示", f"车站{current_station}的TRAINID组数不足，将停止处理该车站",
                                             QMessageBox.Icon.Warning)
                    break

                station_row = station_rows[station_index]

                downbound_index = station_index * 2
                downbound_train_id_row = marker_rows[downbound_index]
                upbound_train_id_row = marker_rows[downbound_index + 1]

                downbound_time_row = station_row
                upbound_time_row = station_row if current_station == '滘口' else station_row + 1

                if upbound_time_row >= len(df) or downbound_time_row >= len(df):
                    self.show_center_message("提示", f"车站{current_station}的时间行超出数据范围，将跳过",
                                             QMessageBox.Icon.Warning)
                    continue

                # 处理上行方向数据
                for col in range(station_column + 1, total_columns):
                    train_id = df.iloc[upbound_train_id_row, col]
                    departure_time = df.iloc[upbound_time_row, col]

                    if pd.notna(train_id) and pd.notna(departure_time):
                        if isinstance(departure_time, pd.Timestamp):
                            departure_time = departure_time.strftime("%H:%M:%S")
                        else:
                            departure_time = str(departure_time).strip()

                        full_id = str(train_id).strip()
                        # 统一格式：{完整ID: time, '__display__': 显示格式, '__source__': 来源}
                        station_schedule[current_station]['上行'].append({
                            full_id: departure_time,
                            '__display__': full_id[-4:].zfill(4),  # Excel导入固定4位
                            '__source__': 'excel'
                        })

                # 处理下行方向数据
                for col in range(station_column - 1, -1, -1):
                    train_id = df.iloc[downbound_train_id_row, col]
                    departure_time = df.iloc[downbound_time_row, col]

                    if pd.notna(train_id) and pd.notna(departure_time):
                        if isinstance(departure_time, pd.Timestamp):
                            departure_time = departure_time.strftime("%H:%M:%S")
                        else:
                            departure_time = str(departure_time).strip()

                        full_id = str(train_id).strip()
                        station_schedule[current_station]['下行'].append({
                            full_id: departure_time,
                            '__display__': full_id[-4:].zfill(4),
                            '__source__': 'excel'
                        })

        return station_schedule

    def query_train_schedule(self):
        """车次查询功能"""
        if not self.processed_schedule:
            self.show_center_message("提示", "请先导入时刻表", QMessageBox.Icon.Information)
            return

        if not self.selected_station:
            self.show_center_message("提示", "请先选择车站", QMessageBox.Icon.Information)
            return

        # 创建查询对话框并保持引用
        self.train_query_dialog = QDialog(self)
        self.train_query_dialog.setWindowTitle(f"{self.selected_station}站 车次查询")
        self.train_query_dialog.setFixedSize(220, 150)  # 缩小宽度，保持高度

        layout = QVBoxLayout(self.train_query_dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        # 输入提示标签 - 增大字体
        input_label = QLabel("请输入4位车次号:")
        input_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        input_label.setStyleSheet("font-size: 13px; font-weight: bold;")  # 增大字体并加粗
        layout.addWidget(input_label)

        # 输入框 - 增大字体
        self.train_id_input = QLineEdit()
        self.train_id_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.train_id_input.setPlaceholderText("例如: 1234")
        self.train_id_input.setStyleSheet("""
            QLineEdit {
                font-size: 13px;  # 增大字体
                border: 1px solid #ccc;
                border-radius: 3px;
                padding: 3px;
            }
        """)
        layout.addWidget(self.train_id_input)

        # 查询按钮 - 增大字体
        confirm_btn = QPushButton("查询")
        confirm_btn.setFixedHeight(28)  # 稍微增加高度
        confirm_btn.clicked.connect(self._perform_train_query)
        confirm_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;  # 增大字体
                background-color: #f0f0f0;
                color: #333333;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        layout.addWidget(confirm_btn)

        # 统一结果显示标签 - 增大字体
        self.result_label = QLabel()
        self.result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.result_label.setWordWrap(True)
        self.result_label.setStyleSheet("font-size: 12px;")  # 增大字体
        self.result_label.setVisible(False)
        layout.addWidget(self.result_label)

        # 居中显示对话框
        screen = QApplication.primaryScreen().availableGeometry()
        dialog_size = self.train_query_dialog.size()
        self.train_query_dialog.move(
            (screen.width() - dialog_size.width()) // 2,
            (screen.height() - dialog_size.height()) // 2
        )

        self.train_query_dialog.exec()

    def _perform_train_query(self):
        """执行车次查询"""
        try:
            input_text = self.train_id_input.text().strip()
            self.result_label.setVisible(False)

            # 验证输入 - 修改为支持3-5位数字
            if not input_text.isdigit() or len(input_text) < 3 or len(input_text) > 5:
                self.result_label.setText("请输入3-5位数字车次号")
                self.result_label.setStyleSheet("color: red; font-size: 11px;")
                self.result_label.setVisible(True)
                return

            full_train_id = input_text.zfill(5) if len(input_text) != 4 else f"0{input_text}"

            # 查询时刻表 - 修改为使用原始时刻表进行查询
            if self.original_schedule:
                station_schedule = self.original_schedule.get(self.selected_station, {})
            else:
                station_schedule = self.processed_schedule.get(self.selected_station, {})

            found = False
            direction = ""
            departure_time = ""

            # 检查上行方向
            for train in station_schedule.get('上行', []):
                if full_train_id in train:
                    direction = "上行方向"
                    departure_time = train[full_train_id]
                    found = True
                    break

            # 检查下行方向
            if not found:
                for train in station_schedule.get('下行', []):
                    if full_train_id in train:
                        direction = "下行方向"
                        departure_time = train[full_train_id]
                        found = True
                        break

            # 显示结果
            if found:
                self.result_label.setText(f"时间: {departure_time}")
                self.result_label.setStyleSheet("""
                    color: green; 
                    font-size: 14px;  # 增大字体到14px
                    font-weight: bold;  # 加粗显示
                """)
            else:
                self.result_label.setText(f"未找到车次 {full_train_id}")
                self.result_label.setStyleSheet("color: red; font-size: 11px;")

            self.result_label.setVisible(True)

        except Exception as e:
            self.result_label.setText(f"查询出错: {str(e)}")
            self.result_label.setStyleSheet("color: red; font-size: 11px;")
            self.result_label.setVisible(True)

    def schedule_query(self):
        """修改时刻表查询按钮的点击事件"""
        if not self.processed_schedule:
            self.show_center_message("提示", "请先导入时刻表", QMessageBox.Icon.Information)
            return

        if not self.selected_station:
            self.show_center_message("提示", "请先选择车站", QMessageBox.Icon.Information)
            return

        self.show_schedule_dialog(self.selected_station, self.processed_schedule)

    def _query_all_stations_stop_time(self, train_id, direction):
        """查询该车次在所有站点的停站时间"""
        if not self.processed_schedule:
            self.show_center_message("提示", "请先导入时刻表", QMessageBox.Icon.Information)
            return

        # 构造完整车次号
        full_train_id = train_id.zfill(5)  # 504 -> 00504, 0504 -> 00504

        # 存储该车次在各站点的停站时间
        stop_times = []

        # 根据方向确定站点顺序
        if direction == '上行':
            # 上行方向：滘口 -> 黄埔新港
            stations_to_check = self.STATION_LIST
        else:  # 下行方向
            # 下行方向：黄埔新港 -> 滘口
            stations_to_check = list(reversed(self.STATION_LIST))

        # 遍历所有站点查找该车次
        for station in stations_to_check:
            station_schedule = self.processed_schedule.get(station, {})
            direction_schedule = station_schedule.get(direction, [])

            for train in direction_schedule:
                if full_train_id in train:
                    stop_times.append({
                        'station': station,
                        'time': train[full_train_id]
                    })
                    break

        # 创建结果对话框
        dialog = QDialog(self)
        dialog.setWindowTitle(f"车次 {train_id} 次 全线各站时间")
        dialog.setFixedSize(300, 400)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(5)

        # 添加方向信息
        direction_label = QLabel(f"方向: {direction}")
        direction_label.setStyleSheet("font-size: 14px; font-weight: bold; text-align: center; margin-bottom: 5px;")
        layout.addWidget(direction_label)

        # 创建结果列表
        result_list = QListWidget()
        result_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 3px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
        """)

        if stop_times:
            # 添加结果到列表
            for stop in stop_times:
                item = QListWidgetItem(f"{stop['station']} - {stop['time']}")
                result_list.addItem(item)
        else:
            item = QListWidgetItem("未找到该车次的停站信息")
            item.setForeground(Qt.GlobalColor.gray)
            result_list.addItem(item)

        layout.addWidget(result_list)

        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.setFixedHeight(30)
        close_btn.clicked.connect(dialog.accept)
        close_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;
                background-color: #f0f0f0;
                color: #333;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        layout.addWidget(close_btn)

        # 居中显示对话框
        screen = QApplication.primaryScreen().availableGeometry()
        dialog_size = dialog.size()
        dialog.move(
            (screen.width() - dialog_size.width()) // 2,
            (screen.height() - dialog_size.height()) // 2
        )

        dialog.exec()

    # 新增：处理共享路径超时的函数
    def _check_shared_path_with_timeout(self, timeout=5):
        """检查共享路径是否可用，带超时功能"""
        import threading

        result = [False]  # 使用列表来在闭包中修改值

        def check_path():
            try:
                # 尝试访问共享路径
                os.listdir(self.SHARED_PATH)
                result[0] = True
            except:
                result[0] = False

        # 创建并启动线程
        thread = threading.Thread(target=check_path)
        thread.start()
        thread.join(timeout)  # 等待指定的超时时间

        return result[0]

    # 新增：自动加载时刻表的函数
    def _auto_load_timetable(self):
        """根据当前爬取的时刻表文本自动加载对应的Excel文件"""
        if not self.current_timetable_text:
            self.show_center_message("错误", "没有可处理的时刻表信息", QMessageBox.Icon.Critical)
            return

        # 创建进度对话框
        progress_dialog = QProgressDialog("正在查找匹配的时刻表文件...", "取消", 0, 0, self)
        progress_dialog.setWindowTitle("自动加载")
        progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.show()

        # 检查共享路径是否可用（5秒超时）
        shared_available = self._check_shared_path_with_timeout(5)

        search_path = self.SHARED_PATH if shared_available else (
            sys._MEIPASS if getattr(sys, 'frozen', False) else os.getcwd())
        search_subdirs = shared_available  # 只有共享路径才搜索子目录

        # 创建文件搜索线程
        self.file_search_thread = FileSearchThread(
            self.current_timetable_text,
            search_path,
            search_subdirs
        )

        # 连接信号槽
        def on_file_found(file_path):
            progress_dialog.close()
            self.file_search_thread.stop()
            # 处理找到的文件
            self._process_timetable_file(file_path)

        def on_file_not_found():
            progress_dialog.close()
            self.show_center_message(
                "未找到文件",
                f"在 {search_path} 中未找到匹配 {self.current_timetable_text} 的Excel文件",
                QMessageBox.Icon.Warning
            )

        def on_error(error_msg):
            progress_dialog.close()
            self.show_center_message("错误", error_msg, QMessageBox.Icon.Critical)

        def on_cancel():
            self.file_search_thread.stop()
            progress_dialog.close()

        self.file_search_thread.found.connect(on_file_found)
        self.file_search_thread.not_found.connect(on_file_not_found)
        self.file_search_thread.error.connect(on_error)
        progress_dialog.canceled.connect(on_cancel)

        # 启动搜索线程
        self.file_search_thread.start()

    def get_timetable_from_web(self):
        """从网页获取时刻表数据"""
        # 创建进度对话框
        progress_dialog = QProgressDialog("正在获取时刻表数据中...", None, 0, 0, self)
        progress_dialog.setWindowTitle("时刻表获取")
        progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.setCancelButton(None)
        progress_dialog.show()

        # 创建并启动爬取线程
        self.crawler_thread = CrawlerThread()

        # 连接信号槽
        def on_finished(result):
            progress_dialog.close()
            today, timetable = result
            if today and timetable:
                # 保存当前爬取的时刻表文本
                self.current_timetable_text = timetable

                # 显示结果对话框
                dialog = QDialog(self)
                dialog.setWindowTitle("时刻表获取结果")
                dialog.setFixedSize(200, 120)  # 增大窗口以容纳新按钮

                layout = QVBoxLayout(dialog)
                layout.setContentsMargins(10, 10, 10, 10)
                layout.setSpacing(10)

                # 日期标签
                date_label = QLabel(f"日期: {today}")
                date_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                date_label.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 2px;")
                layout.addWidget(date_label)

                # 时刻表标签
                timetable_label = QLabel(timetable)
                timetable_label.setStyleSheet(
                    "font-family: 'Courier New Roman', monospace; font-size: 16px; color: red;")
                timetable_label.setWordWrap(True)
                timetable_label.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignTop)
                layout.addWidget(timetable_label)

                # 新增：自动加载按钮
                auto_load_btn = QPushButton("自动加载时刻表")
                auto_load_btn.setFixedSize(120, 30)
                auto_load_btn.clicked.connect(lambda: [dialog.accept(), self._auto_load_timetable()])
                auto_load_btn.setStyleSheet("""
                    QPushButton {
                        font-size: 12px;
                        background-color: #4CAF50;
                        color: white;
                        border: none;
                        border-radius: 3px;
                    }
                    QPushButton:hover {
                        background-color: #45a049;
                    }
                """)
                layout.addWidget(auto_load_btn, alignment=Qt.AlignmentFlag.AlignCenter)

                # 居中显示对话框
                dialog.adjustSize()
                screen_geometry = QApplication.primaryScreen().availableGeometry()
                screen_center = screen_geometry.center()
                dialog.move(
                    screen_center.x() - dialog.width() // 2,
                    screen_center.y() - dialog.height() // 2
                )

                dialog.exec()
            else:
                self.show_center_message("提示", "未获取到时刻表数据", QMessageBox.Icon.Information)

        def on_error(error_msg):
            progress_dialog.close()
            self.show_center_message("错误", f"获取时刻表数据失败：{error_msg}", QMessageBox.Icon.Critical)

        self.crawler_thread.finished.connect(on_finished)
        self.crawler_thread.error.connect(on_error)

        # 启动线程
        self.crawler_thread.start()

    def quit_application(self):
        """完全退出应用程序"""
        # 停止可能在运行的文件搜索线程
        if hasattr(self, 'file_search_thread') and self.file_search_thread.isRunning():
            self.file_search_thread.stop()

        # 关闭本地服务器
        if hasattr(self, 'instance_checker') and self.instance_checker.local_server:
            self.instance_checker.local_server.close()

        # 释放互斥锁
        if hasattr(self, 'instance_checker') and self.instance_checker.mutex_handle:
            ctypes.windll.kernel32.CloseHandle(self.instance_checker.mutex_handle)

        # 完全退出应用程序
        QApplication.quit()
        sys.exit(0)

    def _set_reminder(self, list_widget, item, index, direction):
        """设置提醒功能"""
        # 检查自动更新状态
        if self._check_auto_update_status():
            return

        # 获取当前车次信息
        current_text = item.text()
        match = re.match(r'车次: (\d+)  时间: (.+)', current_text)
        if not match:
            self.show_center_message("错误", "无法解析当前车次信息", QMessageBox.Icon.Critical)
            return

        current_train_id = match.group(1)
        current_time = match.group(2)

        # 获取原始显示格式
        display_train_id = current_train_id

        if self.processed_schedule and self.selected_station:
            station_schedule = self.processed_schedule.get(self.selected_station, {})
            trains = station_schedule.get(direction, [])

            for train in trains:
                if isinstance(train, dict) and '__display__' in train:
                    for key, val in train.items():
                        if not key.startswith('__'):
                            full_train_id = key
                            if fmt_train_id(full_train_id) == fmt_train_id(current_train_id):
                                display_train_id = train.get('__display__', current_train_id)
                                break
                else:
                    full_train_id = list(train.keys())[0]
                    if fmt_train_id(full_train_id) == fmt_train_id(current_train_id):
                        display_train_id = full_train_id[-4:].zfill(4)
                        break

        # 检查时间格式
        try:
            datetime.strptime(current_time, "%H:%M:%S")
            is_standard_time = True
        except ValueError:
            if "文" in current_time:
                is_standard_time = True
            else:
                is_standard_time = False

        if is_standard_time:
            self._set_special_passenger_reminder(current_train_id, current_time, direction, display_train_id)
        else:
            self._set_non_stop_reminder(current_train_id, current_time, direction, index, display_train_id)

    def _set_special_passenger_reminder(self, train_id, train_time, direction, display_train_id):
        """设置特殊乘客提醒"""
        dialog = QDialog(self.schedule_dialog)
        dialog.setWindowTitle("设置提醒")
        dialog.setFixedSize(280, 260)
        dialog.setWindowModality(Qt.WindowModality.WindowModal)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        train_info_label = QLabel(f"车次: {display_train_id}")
        train_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(train_info_label)

        time_info_label = QLabel(f"到达时间: {train_time}")
        time_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(time_info_label)

        reminder_time_layout = QHBoxLayout()
        reminder_time_label = QLabel("提前提醒时间:")
        reminder_time_label.setStyleSheet("font-size: 12px;")
        self.reminder_time_input = QLineEdit()
        self.reminder_time_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.reminder_time_input.setPlaceholderText("例如: 3")
        self.reminder_time_input.setText("3")
        self.reminder_time_input.setStyleSheet("font-size: 12px;")
        reminder_time_unit = QLabel("分钟")
        reminder_time_unit.setStyleSheet("font-size: 12px;")
        reminder_time_layout.addWidget(reminder_time_label)
        reminder_time_layout.addWidget(self.reminder_time_input)
        reminder_time_layout.addWidget(reminder_time_unit)
        layout.addLayout(reminder_time_layout)

        platform_gate_layout = QHBoxLayout()
        platform_gate_label = QLabel("站台门编号:")
        platform_gate_label.setStyleSheet("font-size: 12px;")
        self.platform_gate_combo = QComboBox()
        for i in range(1, 19):
            self.platform_gate_combo.addItem(f"{i}#")
        self.platform_gate_combo.setCurrentText("18#")
        self.platform_gate_combo.setStyleSheet("font-size: 12px;")
        platform_gate_layout.addWidget(platform_gate_label)
        platform_gate_layout.addWidget(self.platform_gate_combo)
        layout.addLayout(platform_gate_layout)

        passenger_type_layout = QHBoxLayout()
        passenger_type_label = QLabel("乘客类别:")
        passenger_type_label.setStyleSheet("font-size: 12px;")
        self.passenger_type_combo = QComboBox()
        passenger_types = [
            "视力不便",
            "行动不便",
            "轮椅需要踏板",
            "轮椅不需要踏板",
            "其他特殊乘客"
        ]
        self.passenger_type_combo.addItems(passenger_types)
        self.passenger_type_combo.setCurrentText("轮椅不需要踏板")
        self.passenger_type_combo.setStyleSheet("font-size: 12px;")
        passenger_type_layout.addWidget(passenger_type_label)
        passenger_type_layout.addWidget(self.passenger_type_combo)
        layout.addLayout(passenger_type_layout)

        self.notes_layout = QHBoxLayout()
        notes_label = QLabel("备注:")
        notes_label.setStyleSheet("font-size: 12px;")
        self.notes_input = QLineEdit()
        self.notes_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.notes_input.setPlaceholderText("请输入乘客特征")
        self.notes_input.setStyleSheet("font-size: 12px;")
        self.notes_layout.addWidget(notes_label)
        self.notes_layout.addWidget(self.notes_input)
        layout.addLayout(self.notes_layout)

        self._toggle_notes_visibility("轮椅不需要踏板")
        self.passenger_type_combo.currentTextChanged.connect(self._toggle_notes_visibility)

        confirm_btn = QPushButton("确认设置")
        confirm_btn.setFixedHeight(30)
        confirm_btn.clicked.connect(dialog.accept)
        confirm_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(confirm_btn)

        screen = QApplication.primaryScreen().availableGeometry()
        dialog_size = dialog.size()
        dialog.move(
            (screen.width() - dialog_size.width()) // 2,
            (screen.height() - dialog_size.height()) // 2
        )

        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                reminder_minutes = int(self.reminder_time_input.text().strip())
                platform_gate = self.platform_gate_combo.currentText()
                passenger_type = self.passenger_type_combo.currentText()

                notes = ""
                if passenger_type == "其他特殊乘客":
                    notes = self.notes_input.text().strip()
                    if not notes:
                        self.show_center_message("提示", "请输入特殊乘客特征", QMessageBox.Icon.Warning)
                        return

                if reminder_minutes < 1 or reminder_minutes > 60:
                    self.show_center_message("错误", "提醒时间请设置在1-60分钟之间", QMessageBox.Icon.Critical)
                    return

                try:
                    time_str = re.sub(r'\s*文', '', train_time).strip() if "文" in train_time else train_time
                    train_time_obj = datetime.strptime(time_str, "%H:%M:%S").time()
                except ValueError:
                    self.show_center_message("错误", "无法解析车次时间格式", QMessageBox.Icon.Critical)
                    return

                now = datetime.now()
                train_datetime = datetime.combine(now.date(), train_time_obj)
                reminder_datetime = train_datetime - timedelta(minutes=reminder_minutes)

                if reminder_datetime < now:
                    self.show_center_message(
                        "时间错误",
                        f"提醒时间 ({reminder_datetime.strftime('%H:%M:%S')}) 早于当前时间 ({now.strftime('%H:%M:%S')})，请重新设置！",
                        QMessageBox.Icon.Warning
                    )
                    return

                reminder = {
                    'type': 'special_passenger',
                    'train_id': train_id,
                    'display_train_id': display_train_id,
                    'train_time': train_time,
                    'reminder_time': reminder_datetime,
                    'platform_gate': platform_gate,
                    'passenger_type': passenger_type,
                    'notes': notes,
                    'direction': direction,
                    'station': self.selected_station,
                    'active': True
                }

                self.reminders.append(reminder)

                success_message = f"提醒设置成功！\n车次: {display_train_id}\n到达时间: {train_time}\n提前提醒: {reminder_minutes}分钟\n站台门: {platform_gate}\n乘客类别: {passenger_type}"
                if notes:
                    success_message += f"\n备注: {notes}"

                self.show_center_message("成功", success_message, QMessageBox.Icon.Information)

            except ValueError:
                self.show_center_message("错误", "请输入有效的数字作为提醒时间", QMessageBox.Icon.Critical)
            except Exception as e:
                self.show_center_message("错误", f"设置提醒失败: {str(e)}", QMessageBox.Icon.Critical)

    def _toggle_notes_visibility(self, passenger_type):
        """切换备注框的显示/隐藏"""
        if passenger_type == "其他特殊乘客":
            # 显示备注框
            for i in range(self.notes_layout.count()):
                widget = self.notes_layout.itemAt(i).widget()
                if widget:
                    widget.show()
        else:
            # 隐藏备注框并清空内容
            for i in range(self.notes_layout.count()):
                widget = self.notes_layout.itemAt(i).widget()
                if widget:
                    widget.hide()
            if hasattr(self, 'notes_input'):
                self.notes_input.clear()

    def _set_non_stop_reminder(self, train_id, train_time, direction, index, display_train_id):
        """设置不停站提醒"""
        reply = self.show_center_question(
            "设置不停站提醒",
            f"确定要为车次 {display_train_id} 设置不停站提醒吗？\n该提醒将在上一趟列车到达后立即触发。"
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            previous_train_time = self._find_previous_train_time_by_index(direction, index)
            now = datetime.now()

            if previous_train_time:
                reminder_datetime = datetime.combine(now.date(), previous_train_time)
            else:
                reminder_datetime = now

            if reminder_datetime < now:
                self.show_center_message(
                    "时间错误",
                    f"提醒时间 ({reminder_datetime.strftime('%H:%M:%S')}) 早于当前时间 ({now.strftime('%H:%M:%S')})，请重新设置！\n\n提示：上一趟列车时间已过，请手动设置提醒时间。",
                    QMessageBox.Icon.Warning
                )
                return

            reminder = {
                'type': 'non_stop',
                'train_id': display_train_id,
                'train_time': train_time,
                'reminder_time': reminder_datetime,
                'direction': direction,
                'station': self.selected_station,
                'active': True
            }
            self.reminders.append(reminder)

            self.show_center_message(
                "成功",
                f"不停站提醒设置成功！\n"
                f"车次: {display_train_id}\n"
                f"列车信息: {train_time}\n"
                f"方向: {direction}\n"
                f"提醒时间: {reminder_datetime.strftime('%H:%M:%S')}",
                QMessageBox.Icon.Information
            )

        except Exception as e:
            self.show_center_message("错误", f"设置不停站提醒失败: {str(e)}",
                                     QMessageBox.Icon.Critical)

    def _find_previous_train_time_by_index(self, direction, current_index):
        """根据当前车次的索引查找上一趟列车的时间（修改点1：新增方法）"""
        if not self.processed_schedule or not self.selected_station or current_index <= 0:
            return None

        station_schedule = self.processed_schedule.get(self.selected_station, {})
        trains = station_schedule.get(direction, [])

        # 获取当前车次的上一趟列车（索引减1）
        if current_index - 1 < len(trains):
            previous_train = trains[current_index - 1]
            for _, time_str in previous_train.items():
                try:
                    # 尝试解析时间格式
                    return datetime.strptime(time_str, "%H:%M:%S").time()
                except ValueError:
                    # 非标准时间格式，跳过
                    continue

        return None

    def _add_all_non_stop_reminders(self, station, schedule_data):
        """一键添加所有不停站提醒"""
        if not schedule_data or not station:
            self.show_center_message("错误", "无法获取时刻表数据", QMessageBox.Icon.Critical)
            return

        reply = self.show_center_question(
            "一键添加所有不停站提醒",
            "确定添加所有不停站提醒？仅限6:30-23:00且晚于当前时间的。"
        )

        if reply != QMessageBox.StandardButton.Yes:
            return

        now = datetime.now()
        current_time = now.time()

        start_time = datetime.strptime("06:30:00", "%H:%M:%S").time()
        end_time = datetime.strptime("23:00:00", "%H:%M:%S").time()

        total_found = 0
        added_count = 0
        skipped_count = 0

        for direction in ['上行', '下行']:
            trains = schedule_data.get(station, {}).get(direction, [])

            for index, train in enumerate(trains):
                for train_id, time_str in train.items():
                    if train_id.startswith('__'):
                        continue

                    # 修复：确保Excel导入的车次显示为4位格式
                    # 如果 __display__ 不存在，使用 fmt_train_id 后再补零到4位
                    display_train_id = train.get('__display__') or fmt_train_id(train_id).zfill(4)

                    try:
                        time_obj = datetime.strptime(time_str, "%H:%M:%S").time()
                        continue
                    except ValueError:
                        total_found += 1

                        previous_train_time = self._find_previous_train_time_by_index(direction, index)

                        if previous_train_time:
                            if start_time <= previous_train_time <= end_time:
                                previous_train_datetime = datetime.combine(now.date(), previous_train_time)
                                reminder_datetime = previous_train_datetime

                                if reminder_datetime > now:
                                    reminder = {
                                        'type': 'non_stop',
                                        'train_id': display_train_id,
                                        'train_time': time_str,
                                        'reminder_time': reminder_datetime,
                                        'direction': direction,
                                        'station': station,
                                        'active': True
                                    }
                                    self.reminders.append(reminder)
                                    added_count += 1
                                else:
                                    skipped_count += 1
                            else:
                                skipped_count += 1
                        else:
                            skipped_count += 1

        if total_found == 0:
            self.show_center_message("提示", "未发现不停站列车", QMessageBox.Icon.Information)
        else:
            message = f"共发现 {total_found} 个不停站列车\n"
            message += f"成功添加 {added_count} 个提醒\n"
            if skipped_count > 0:
                message += f"跳过 {skipped_count} 个提醒（时间不在范围内或已过期）"

            self.show_center_message("一键添加完成", message, QMessageBox.Icon.Information)

    def _find_previous_train_time(self, direction):
        """查找上一趟列车的时间（保留原方法，用于其他用途）"""
        if not self.processed_schedule or not self.selected_station:
            return None

        current_time = datetime.now().time()
        station_schedule = self.processed_schedule.get(self.selected_station, {})
        trains = station_schedule.get(direction, [])

        # 查找当前时间之前的最后一趟列车
        previous_time = None
        for train in trains:
            for _, time_str in train.items():
                try:
                    # 尝试解析时间格式
                    train_time = datetime.strptime(time_str, "%H:%M:%S").time()

                    # 如果这个时间在当前时间之前，且比之前找到的时间更接近当前时间
                    if train_time < current_time:
                        if not previous_time or train_time > previous_time:
                            previous_time = train_time
                except ValueError:
                    # 非标准时间格式，跳过
                    continue

        return previous_time

    def check_reminders(self):
        """检查是否有到期的提醒"""
        now = datetime.now()
        reminders_to_remove = []

        for i, reminder in enumerate(self.reminders):
            if reminder['active'] and now >= reminder['reminder_time']:
                # 根据提醒类型显示不同的弹窗
                if reminder['type'] == 'special_passenger':
                    self.show_special_passenger_reminder_popup(reminder)
                elif reminder['type'] == 'non_stop':
                    self.show_non_stop_reminder_popup(reminder)
                reminders_to_remove.append(i)

        # 移除已触发的提醒
        for i in reversed(reminders_to_remove):
            self.reminders.pop(i)

    def show_special_passenger_reminder_popup(self, reminder):
        """显示特殊乘客提醒弹窗，置顶显示"""
        # 创建提醒弹窗
        popup = QDialog()
        popup.setWindowTitle("🚨 特殊乘客提醒 🚨")

        # 根据是否有备注调整窗口大小
        has_notes = reminder.get('notes', '')
        popup.setFixedSize(300, 180 if has_notes else 160)  # 有备注时增加高度

        # 设置窗口为置顶
        popup.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Dialog
        )

        # 设置弹窗样式 - 黄色背景
        popup.setStyleSheet("""
            QDialog {
                background-color: #fff3cd;
                border: 2px solid #ffc107;
                border-radius: 5px;
            }
            QLabel {
                color: #856404;
                font-family: "Microsoft YaHei UI";
            }
            QPushButton {
                background-color: #ffc107;
                color: #856404;
                border: none;
                border-radius: 3px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0a800;
            }
        """)

        layout = QVBoxLayout(popup)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # 添加提醒内容
        direction_label = QLabel(f"方向: {reminder['direction']}")
        direction_label.setStyleSheet("font-size: 14px;")
        layout.addWidget(direction_label)

        display_id = reminder.get('display_train_id', reminder['train_id'])
        train_label = QLabel(f"车次: {display_id}")
        train_label.setStyleSheet("font-size: 14px;")
        layout.addWidget(train_label)

        time_label = QLabel(f"到达时间: {reminder['train_time']}")
        time_label.setStyleSheet("font-size: 14px;")
        layout.addWidget(time_label)

        gate_label = QLabel(f"站台门: {reminder['platform_gate']}")
        gate_label.setStyleSheet("font-size: 14px;")
        layout.addWidget(gate_label)

        passenger_label = QLabel(f"乘客类别: {reminder['passenger_type']}")
        passenger_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(passenger_label)

        # ===== 新增：显示备注（如果有）=====
        if has_notes:
            notes_label = QLabel(f"备注: {reminder['notes']}")
            notes_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #d9534f;")  # 红色突出显示
            layout.addWidget(notes_label)

        # 确认按钮
        confirm_btn = QPushButton("我已了解")
        confirm_btn.setFixedHeight(30)
        confirm_btn.clicked.connect(popup.accept)
        layout.addWidget(confirm_btn)

        # 居中显示弹窗
        screen = QApplication.primaryScreen().availableGeometry()
        popup_size = popup.size()
        popup.move(
            (screen.width() - popup_size.width()) // 2,
            (screen.height() - popup_size.height()) // 2
        )

        # 显示弹窗
        popup.exec()

    def show_non_stop_reminder_popup(self, reminder):
        """显示不停站提醒弹窗，置顶显示，红色背景"""
        # 创建提醒弹窗
        popup = QDialog()
        popup.setWindowTitle("🚨 不停站提醒 🚨")
        popup.setFixedSize(280, 140)  # 适当的窗口大小

        # 设置窗口为置顶
        popup.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Dialog
        )

        # 设置弹窗样式 - 红色背景
        popup.setStyleSheet("""
            QDialog {
                background-color: #f8d7da;
                border: 2px solid #dc3545;
                border-radius: 5px;
            }
            QLabel {
                color: #721c24;
                font-family: "Microsoft YaHei UI";
            }
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 5px 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
        """)

        layout = QVBoxLayout(popup)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # 添加提醒内容
        warning_label = QLabel("⚠️ 注意！有列车不停站通过！")
        warning_label.setStyleSheet("font-size: 16px; font-weight: bold; text-align: center;")
        layout.addWidget(warning_label)

        direction_label = QLabel(f"方向: {reminder['direction']}")
        direction_label.setStyleSheet("font-size: 14px; text-align: center;")
        layout.addWidget(direction_label)

        display_id = reminder.get('display_train_id', reminder['train_id'])
        train_label = QLabel(f"车次: {display_id}")
        train_label.setStyleSheet("font-size: 14px; text-align: center;")
        layout.addWidget(train_label)

        # 确认按钮
        confirm_btn = QPushButton("确认")
        confirm_btn.setFixedHeight(30)
        confirm_btn.clicked.connect(popup.accept)
        layout.addWidget(confirm_btn)

        # 居中显示弹窗
        screen = QApplication.primaryScreen().availableGeometry()
        popup_size = popup.size()
        popup.move(
            (screen.width() - popup_size.width()) // 2,
            (screen.height() - popup_size.height()) // 2
        )

        # 显示弹窗
        popup.exec()

    def show_reminders_panel(self):
        """显示已设置提醒面板"""
        # 创建提醒面板对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("已设置提醒管理")
        dialog.setFixedSize(600, 400)  # 适当的窗口大小

        # 设置对话框字体
        font = QFont("Microsoft YaHei UI", 10)
        dialog.setFont(font)

        # 主布局
        main_layout = QVBoxLayout(dialog)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # 标题
        title_label = QLabel("已设置提醒列表")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; text-align: center; margin-bottom: 5px;")
        main_layout.addWidget(title_label)

        # 创建提醒列表表格
        self.reminders_table = QTableWidget()
        self.reminders_table.setColumnCount(8)  # 增加一列用于操作
        self.reminders_table.setHorizontalHeaderLabels([
            "提醒类型", "车站", "方向", "车次", "到达时间", "提醒时间", "状态", "操作"
        ])

        # 设置表格样式
        self.reminders_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                border-radius: 3px;
                gridline-color: #ddd;
            }
            QTableWidget::item {
                padding: 5px;
                border-bottom: 1px solid #eee;
            }
            QTableWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                color: #333;
                padding: 5px;
                border: 1px solid #ddd;
                text-align: center;
            }
            QPushButton {
                background-color: #f0f0f0;
                color: #333;
                border: 1px solid #ddd;
                border-radius: 3px;
                padding: 3px 8px;
                font-size: 9px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton[delete="true"] {
                background-color: #f8d7da;
                color: #721c24;
                border-color: #f5c6cb;
            }
            QPushButton[delete="true"]:hover {
                background-color: #f5c6cb;
            }
        """)

        # 设置列宽
        column_widths = [100, 80, 80, 80, 80, 80, 60, 80]
        for i, width in enumerate(column_widths):
            self.reminders_table.setColumnWidth(i, width)

        # 设置表格为不可编辑
        self.reminders_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

        # 添加表格到布局
        main_layout.addWidget(self.reminders_table)

        # 添加当前时间显示
        self.reminders_time_label = QLabel()
        self.reminders_time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.reminders_time_label.setStyleSheet("font-size: 14px; color: #666;")
        main_layout.addWidget(self.reminders_time_label)

        # 更新当前时间
        self._update_reminders_time()

        # 添加时间更新定时器
        self.reminders_time_timer = QTimer()
        self.reminders_time_timer.setInterval(1000)
        self.reminders_time_timer.timeout.connect(self._update_reminders_time)
        self.reminders_time_timer.start()

        # 自动刷新提醒列表
        self._refresh_reminders_table()

        # 添加删除全部按钮（新增功能）
        delete_all_btn = QPushButton("删除全部")
        delete_all_btn.setFixedSize(100, 30)
        delete_all_btn.clicked.connect(self._delete_all_reminders)
        delete_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 3px;
                padding: 5px 10px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
        """)

        # 添加按钮布局
        button_layout = QHBoxLayout()
        button_layout.addWidget(delete_all_btn)  # 添加删除全部按钮
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        # 对话框关闭时停止定时器
        def on_dialog_closed():
            self.reminders_time_timer.stop()

        dialog.finished.connect(on_dialog_closed)

        # 居中显示对话框
        dialog.adjustSize()
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        screen_center = screen_geometry.center()
        dialog.move(
            screen_center.x() - dialog.width() // 2,
            screen_center.y() - dialog.height() // 2
        )

        dialog.exec()

    def _refresh_reminders_table(self):
        """刷新提醒列表表格"""
        # 清空表格
        self.reminders_table.setRowCount(0)

        # 添加提醒数据
        now = datetime.now()
        for i, reminder in enumerate(self.reminders):
            # 跳过已过期但未被清理的提醒
            if not reminder['active'] or (reminder['active'] and now >= reminder['reminder_time']):
                continue

            row_position = self.reminders_table.rowCount()
            self.reminders_table.insertRow(row_position)

            # 提醒类型
            type_item = QTableWidgetItem("特殊乘客" if reminder['type'] == 'special_passenger' else "不停站")
            type_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.reminders_table.setItem(row_position, 0, type_item)

            # 车站
            station_item = QTableWidgetItem(reminder['station'])
            station_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.reminders_table.setItem(row_position, 1, station_item)

            # 方向
            direction_item = QTableWidgetItem(reminder['direction'])
            direction_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.reminders_table.setItem(row_position, 2, direction_item)

            # 车次 - 修改：优先使用 display_train_id，如果没有则使用 train_id，不再做 fmt_train_id 转换
            display_id = reminder.get('display_train_id') or reminder.get('train_id', '')
            train_id_item = QTableWidgetItem(display_id)
            train_id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.reminders_table.setItem(row_position, 3, train_id_item)

            # 到达时间
            train_time_item = QTableWidgetItem(reminder['train_time'])
            train_time_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.reminders_table.setItem(row_position, 4, train_time_item)

            # 提醒时间
            reminder_time_item = QTableWidgetItem(reminder['reminder_time'].strftime("%H:%M:%S"))
            reminder_time_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.reminders_table.setItem(row_position, 5, reminder_time_item)

            # 状态
            status_item = QTableWidgetItem("待触发")
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            status_item.setForeground(QColor("green"))
            self.reminders_table.setItem(row_position, 6, status_item)

            # 操作按钮 - 修改按钮样式
            button_widget = QWidget()
            button_layout = QHBoxLayout(button_widget)
            button_layout.setContentsMargins(0, 0, 0, 0)
            button_layout.setSpacing(2)

            # 修改按钮
            modify_btn = QPushButton("修改")
            modify_btn.setFixedSize(40, 20)
            modify_btn.clicked.connect(lambda checked, idx=i: self._modify_reminder(idx))
            button_layout.addWidget(modify_btn)

            # 删除按钮 - 添加delete属性
            delete_btn = QPushButton("删除")
            delete_btn.setFixedSize(40, 20)
            delete_btn.setProperty("delete", "true")  # 添加属性用于样式控制
            delete_btn.clicked.connect(lambda checked, idx=i: self._delete_reminder(idx))
            button_layout.addWidget(delete_btn)

            self.reminders_table.setCellWidget(row_position, 7, button_widget)

        # 如果没有提醒，显示提示信息
        if self.reminders_table.rowCount() == 0:
            empty_item = QTableWidgetItem("暂无已设置的提醒")
            empty_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            empty_item.setForeground(QColor("gray"))
            self.reminders_table.insertRow(0)
            self.reminders_table.setItem(0, 0, empty_item)
            self.reminders_table.setSpan(0, 0, 1, 8)  # 合并所有列

    def _update_reminders_time(self):
        """更新当前时间显示"""
        if hasattr(self, 'reminders_time_label'):
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.reminders_time_label.setText(f"当前时间: {current_time}")

    def _modify_reminder(self, index):
        """修改提醒"""
        if index < 0 or index >= len(self.reminders):
            self.show_center_message("错误", "无效的提醒索引", QMessageBox.Icon.Critical)
            return

        reminder = self.reminders[index]

        # 根据提醒类型显示不同的修改对话框
        if reminder['type'] == 'special_passenger':
            self._modify_special_passenger_reminder(index, reminder)
        elif reminder['type'] == 'non_stop':
            self._modify_non_stop_reminder(index, reminder)

    def _modify_special_passenger_reminder(self, index, reminder):
        """修改特殊乘客提醒"""
        dialog = QDialog()
        dialog.setWindowTitle("修改提醒")
        dialog.setFixedSize(280, 260)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        display_id = reminder.get('display_train_id', reminder['train_id'])
        train_info_label = QLabel(f"车次: {display_id}")
        train_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(train_info_label)

        time_info_label = QLabel(f"到达时间: {reminder['train_time']}")
        time_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(time_info_label)

        reminder_time_layout = QHBoxLayout()
        reminder_time_label = QLabel("提前提醒时间:")
        reminder_time_label.setStyleSheet("font-size: 12px;")
        self.reminder_time_input = QLineEdit()
        self.reminder_time_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.reminder_time_input.setPlaceholderText("例如: 3")

        now = datetime.now()
        time_str = reminder['train_time'].replace("  文", "").strip() if "文" in reminder['train_time'] else reminder[
            'train_time']
        train_datetime = datetime.combine(now.date(), datetime.strptime(time_str, "%H:%M:%S").time())
        reminder_datetime = reminder['reminder_time']
        advance_minutes = (train_datetime - reminder_datetime).total_seconds() // 60
        self.reminder_time_input.setText(str(int(advance_minutes)))

        self.reminder_time_input.setStyleSheet("font-size: 12px;")
        reminder_time_unit = QLabel("分钟")
        reminder_time_unit.setStyleSheet("font-size: 12px;")
        reminder_time_layout.addWidget(reminder_time_label)
        reminder_time_layout.addWidget(self.reminder_time_input)
        reminder_time_layout.addWidget(reminder_time_unit)
        layout.addLayout(reminder_time_layout)

        platform_gate_layout = QHBoxLayout()
        platform_gate_label = QLabel("站台门编号:")
        platform_gate_label.setStyleSheet("font-size: 12px;")
        self.platform_gate_combo = QComboBox()
        for i in range(1, 19):
            self.platform_gate_combo.addItem(f"{i}#")
        self.platform_gate_combo.setCurrentText(reminder['platform_gate'])
        self.platform_gate_combo.setStyleSheet("font-size: 12px;")
        platform_gate_layout.addWidget(platform_gate_label)
        platform_gate_layout.addWidget(self.platform_gate_combo)
        layout.addLayout(platform_gate_layout)

        passenger_type_layout = QHBoxLayout()
        passenger_type_label = QLabel("乘客类别:")
        passenger_type_label.setStyleSheet("font-size: 12px;")
        self.passenger_type_combo = QComboBox()
        passenger_types = [
            "视力不便",
            "行动不便",
            "轮椅需要踏板",
            "轮椅不需要踏板",
            "其他特殊乘客"
        ]
        self.passenger_type_combo.addItems(passenger_types)
        self.passenger_type_combo.setCurrentText(reminder['passenger_type'])
        self.passenger_type_combo.setStyleSheet("font-size: 12px;")
        passenger_type_layout.addWidget(passenger_type_label)
        passenger_type_layout.addWidget(self.passenger_type_combo)
        layout.addLayout(passenger_type_layout)

        self.notes_layout = QHBoxLayout()
        notes_label = QLabel("备注:")
        notes_label.setStyleSheet("font-size: 12px;")
        self.notes_input = QLineEdit()
        self.notes_input.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.notes_input.setPlaceholderText("请输入特殊乘客特征")
        self.notes_input.setStyleSheet("font-size: 12px;")
        self.notes_input.setText(reminder.get('notes', ''))
        self.notes_layout.addWidget(notes_label)
        self.notes_layout.addWidget(self.notes_input)
        layout.addLayout(self.notes_layout)

        self._toggle_notes_visibility(reminder['passenger_type'])
        self.passenger_type_combo.currentTextChanged.connect(self._toggle_notes_visibility)

        confirm_btn = QPushButton("确认修改")
        confirm_btn.setFixedHeight(30)
        confirm_btn.clicked.connect(dialog.accept)
        confirm_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(confirm_btn)

        screen = QApplication.primaryScreen().availableGeometry()
        dialog_size = dialog.size()
        dialog.move(
            (screen.width() - dialog_size.width()) // 2,
            (screen.height() - dialog_size.height()) // 2
        )

        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                reminder_minutes = int(self.reminder_time_input.text().strip())
                platform_gate = self.platform_gate_combo.currentText()
                passenger_type = self.passenger_type_combo.currentText()

                notes = ""
                if passenger_type == "其他特殊乘客":
                    notes = self.notes_input.text().strip()
                    if not notes:
                        self.show_center_message("提示", "请输入特殊乘客特征", QMessageBox.Icon.Warning)
                        return

                if reminder_minutes < 1 or reminder_minutes > 60:
                    self.show_center_message("错误", "提醒时间请设置在1-60分钟之间", QMessageBox.Icon.Critical)
                    return

                try:
                    time_str = reminder['train_time']
                    # 处理"文"字标记：移除"文"及其前面的空白字符（空格或制表符）
                    if "文" in time_str:
                        time_str = re.sub(r'\s*文', '', time_str).strip()
                    train_time_obj = datetime.strptime(time_str, "%H:%M:%S").time()
                except ValueError:
                    self.show_center_message("错误", "无法解析车次时间格式", QMessageBox.Icon.Critical)
                    return

                now = datetime.now()
                train_datetime = datetime.combine(now.date(), train_time_obj)
                reminder_datetime = train_datetime - timedelta(minutes=reminder_minutes)

                if reminder_datetime < now:
                    self.show_center_message(
                        "时间错误",
                        f"提醒时间 ({reminder_datetime.strftime('%H:%M:%S')}) 早于当前时间 ({now.strftime('%H:%M:%S')})，请重新设置！",
                        QMessageBox.Icon.Warning
                    )
                    return

                self.reminders[index]['reminder_time'] = reminder_datetime
                self.reminders[index]['platform_gate'] = platform_gate
                self.reminders[index]['passenger_type'] = passenger_type
                self.reminders[index]['notes'] = notes

                self._refresh_reminders_table()

                success_message = f"提醒修改成功！\n车次: {display_id}\n提醒时间: {reminder_datetime.strftime('%H:%M:%S')}"
                if notes:
                    success_message += f"\n特征备注: {notes}"

                self.show_center_message("成功", success_message, QMessageBox.Icon.Information)

            except ValueError:
                self.show_center_message("错误", "请输入有效的数字作为提醒时间", QMessageBox.Icon.Critical)
            except Exception as e:
                self.show_center_message("错误", f"修改提醒失败: {str(e)}", QMessageBox.Icon.Critical)

    def _modify_non_stop_reminder(self, index, reminder):
        """修改不停站提醒"""
        dialog = QDialog()
        dialog.setWindowTitle("修改不停站提醒")
        dialog.setFixedSize(280, 180)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        display_id = reminder.get('display_train_id', reminder['train_id'])
        train_info_label = QLabel(f"车次: {display_id}")
        train_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(train_info_label)

        time_info_label = QLabel(f"列车信息: {reminder['train_time']}")
        time_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(time_info_label)

        direction_info_label = QLabel(f"方向: {reminder['direction']}")
        direction_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(direction_info_label)

        reminder_time_layout = QHBoxLayout()
        reminder_time_label = QLabel("提醒时间:")
        reminder_time_label.setStyleSheet("font-size: 12px;")
        self.reminder_time_edit = QLineEdit()
        self.reminder_time_edit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.reminder_time_edit.setText(reminder['reminder_time'].strftime("%H:%M:%S"))
        self.reminder_time_edit.setStyleSheet("font-size: 12px;")
        reminder_time_layout.addWidget(reminder_time_label)
        reminder_time_layout.addWidget(self.reminder_time_edit)
        layout.addLayout(reminder_time_layout)

        confirm_btn = QPushButton("确认修改")
        confirm_btn.setFixedHeight(30)
        confirm_btn.clicked.connect(dialog.accept)
        confirm_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(confirm_btn)

        screen = QApplication.primaryScreen().availableGeometry()
        dialog_size = dialog.size()
        dialog.move(
            (screen.width() - dialog_size.width()) // 2,
            (screen.height() - dialog_size.height()) // 2
        )

        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                reminder_time_str = self.reminder_time_edit.text().strip()

                try:
                    reminder_time_obj = datetime.strptime(reminder_time_str, "%H:%M:%S").time()
                except ValueError:
                    self.show_center_message("错误", "请输入正确的时间格式（HH:MM:SS）", QMessageBox.Icon.Critical)
                    return

                now = datetime.now()
                reminder_datetime = datetime.combine(now.date(), reminder_time_obj)

                if reminder_datetime < now:
                    self.show_center_message(
                        "时间错误",
                        f"提醒时间 ({reminder_datetime.strftime('%H:%M:%S')}) 早于当前时间 ({now.strftime('%H:%M:%S')})，请重新设置！",
                        QMessageBox.Icon.Warning
                    )
                    return

                self.reminders[index]['reminder_time'] = reminder_datetime

                self._refresh_reminders_table()

                self.show_center_message(
                    "成功",
                    f"不停站提醒修改成功！\n车次: {display_id}\n提醒时间: {reminder_datetime.strftime('%H:%M:%S')}",
                    QMessageBox.Icon.Information
                )

            except Exception as e:
                self.show_center_message("错误", f"修改提醒失败: {str(e)}", QMessageBox.Icon.Critical)

    def _delete_reminder(self, index):
        """删除提醒"""
        if index < 0 or index >= len(self.reminders):
            self.show_center_message("错误", "无效的提醒索引", QMessageBox.Icon.Critical)
            return

        reminder = self.reminders[index]

        # 确认删除
        reply = self.show_center_question(
            "确认删除",
            f"确定要删除车次 {reminder['train_id']} 的提醒吗？"
        )

        if reply == QMessageBox.StandardButton.Yes:
            # 删除提醒
            self.reminders.pop(index)

            # 刷新提醒列表
            self._refresh_reminders_table()

            # 显示成功信息
            self.show_center_message("成功", "提醒已删除", QMessageBox.Icon.Information)

    def _delete_all_reminders(self):
        """删除全部提醒（新增功能）"""
        if not self.reminders:
            self.show_center_message("提示", "没有已设置的提醒", QMessageBox.Icon.Information)
            return

        # 确认删除全部
        reply = self.show_center_question(
            "确认删除全部",
            f"确定要删除所有 {len(self.reminders)} 个提醒吗？"
        )

        if reply == QMessageBox.StandardButton.Yes:
            # 清空所有提醒
            self.reminders.clear()

            # 刷新提醒列表
            self._refresh_reminders_table()

            # 显示成功信息
            self.show_center_message("成功", "所有提醒已删除", QMessageBox.Icon.Information)

    def _send_message_via_glink(self, list_widget, item, index, direction):
        """通过Glink发送消息"""
        if self._check_auto_update_status():
            return

        current_user = os.getlogin()
        has_permission = False

        if current_user == 'dongwuyuanzhan':
            permission_file = r"\\淘金站行值a\陈孟熙\yp\Project\时刻表助手\权限\zoo.txt"
            if os.path.exists(permission_file):
                has_permission = True
        elif current_user == 'taojinzhan':
            permission_file = r"E:\陈孟熙\yp\Project\时刻表助手\权限\taojin.txt"
            if os.path.exists(permission_file):
                has_permission = True
        else:
            has_permission = False

        if not has_permission:
            self.show_center_message("权限不足", "您没有使用Glink发送消息的权限", QMessageBox.Icon.Warning)
            return

        current_text = item.text()
        match = re.match(r'车次: (\d+)  时间: (.+)', current_text)
        if not match:
            self.show_center_message("错误", "无法解析当前车次信息", QMessageBox.Icon.Critical)
            return

        current_train_id = match.group(1)
        current_time = match.group(2)

        display_train_id = current_train_id

        if self.processed_schedule and self.selected_station:
            station_schedule = self.processed_schedule.get(self.selected_station, {})
            trains = station_schedule.get(direction, [])

            for train in trains:
                if isinstance(train, dict) and '__display__' in train:
                    for key, val in train.items():
                        if not key.startswith('__'):
                            full_train_id = key
                            if fmt_train_id(full_train_id) == fmt_train_id(current_train_id):
                                display_train_id = train.get('__display__', current_train_id)
                                break
                else:
                    full_train_id = list(train.keys())[0]
                    if fmt_train_id(full_train_id) == fmt_train_id(current_train_id):
                        display_train_id = full_train_id[-4:].zfill(4)
                        break

        dialog = QDialog()
        dialog.setWindowTitle("Glink发送消息")
        dialog.setFixedSize(300, 220)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        train_info_label = QLabel(f"车次: {display_train_id}")
        train_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(train_info_label)

        time_info_label = QLabel(f"到达时间: {current_time}")
        time_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(time_info_label)

        direction_info_label = QLabel(f"方向: {direction}")
        direction_info_label.setStyleSheet("font-size: 13px; font-weight: bold;")
        layout.addWidget(direction_info_label)

        receiver_layout = QHBoxLayout()
        receiver_label = QLabel("接收人:")
        receiver_label.setStyleSheet("font-size: 12px;")
        self.receiver_combo = QComboBox()

        all_stations = [
            '滘口', '坦尾', '中山八', '西场', '西村', '广州火车站', '小北', '淘金', '区庄', '动物园',
            '杨箕', '五羊邨', '珠江新城', '猎德', '潭村', '员村', '科韵路', '车陂南', '东圃', '三溪',
            '鱼珠', '大沙地', '大沙东', '文冲', '双沙', '庙头', '夏园', '保盈大道', '夏港', '黄埔新港'
        ]

        filtered_stations = []
        if hasattr(self, 'selected_station') and self.selected_station in all_stations:
            current_index = all_stations.index(self.selected_station)

            if direction == '下行':
                filtered_stations = all_stations[:current_index + 1]
            elif direction == '上行':
                filtered_stations = all_stations[current_index:]

        if not filtered_stations:
            filtered_stations = all_stations

        self.receiver_combo.addItems(filtered_stations)

        if hasattr(self, 'selected_station') and self.selected_station in filtered_stations:
            self.receiver_combo.setCurrentText(self.selected_station)
        self.receiver_combo.setStyleSheet("font-size: 12px;")
        receiver_layout.addWidget(receiver_label)
        receiver_layout.addWidget(self.receiver_combo)
        layout.addLayout(receiver_layout)

        platform_gate_layout = QHBoxLayout()
        platform_gate_label = QLabel("站台门编号:")
        platform_gate_label.setStyleSheet("font-size: 12px;")
        self.glink_platform_gate_combo = QComboBox()
        for i in range(1, 19):
            self.glink_platform_gate_combo.addItem(f"{i}#")
        self.glink_platform_gate_combo.setCurrentText("18#")
        self.glink_platform_gate_combo.setStyleSheet("font-size: 12px;")
        platform_gate_layout.addWidget(platform_gate_label)
        platform_gate_layout.addWidget(self.glink_platform_gate_combo)
        layout.addLayout(platform_gate_layout)

        wheelchair_layout = QHBoxLayout()
        wheelchair_label = QLabel("乘客类别:")
        wheelchair_label.setStyleSheet("font-size: 12px;")
        self.wheelchair_combo = QComboBox()
        wheelchair_types = [
            "视力不便",
            "行动不便",
            "轮椅需要踏板",
            "轮椅不需要踏板",
            "其他特殊乘客"
        ]
        self.wheelchair_combo.addItems(wheelchair_types)
        self.wheelchair_combo.setCurrentText("轮椅不需要踏板")
        self.wheelchair_combo.setStyleSheet("font-size: 12px;")
        wheelchair_layout.addWidget(wheelchair_label)
        wheelchair_layout.addWidget(self.wheelchair_combo)
        layout.addLayout(wheelchair_layout)

        confirm_btn = QPushButton("发送消息")
        confirm_btn.setFixedHeight(30)
        confirm_btn.clicked.connect(dialog.accept)
        confirm_btn.setStyleSheet("""
            QPushButton {
                font-size: 12px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        layout.addWidget(confirm_btn)

        screen = QApplication.primaryScreen().availableGeometry()
        dialog_size = dialog.size()
        dialog.move(
            (screen.width() - dialog_size.width()) // 2,
            (screen.height() - dialog_size.height()) // 2
        )

        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                receiver = self.receiver_combo.currentText()
                platform_gate = self.glink_platform_gate_combo.currentText()
                wheelchair_type = self.wheelchair_combo.currentText()

                message = f"{display_train_id}次，{platform_gate}，{wheelchair_type}"

                reply = self.show_center_question(
                    "Glink发送信息",
                    f"确定要发送以下信息吗？\n\n{message}\n\n接收人: {receiver}"
                )

                if reply == QMessageBox.StandardButton.Yes:
                    progress_dialog = QProgressDialog("正在发送消息...", None, 0, 0, self)
                    progress_dialog.setWindowTitle("Glink发送")
                    progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
                    progress_dialog.setMinimumDuration(0)
                    progress_dialog.setCancelButton(None)
                    progress_dialog.show()

                    self.glink = Glink()

                    current_user = os.getlogin()
                    if current_user == 'dongwuyuanzhan':
                        self.glink.username = 'dongwuyuanzhan'
                        self.glink.password = 'Dwy57*010'
                        self.glink.MY_USER_ID = '80cb0e95-d022-44e2-9972-7ecddbf0df99'
                        self.glink.MY_USER_NAME = '动物园站'

                    self.glink_thread = GlinkThread(self.glink)

                    def on_login_result(success, login_message):
                        progress_dialog.close()
                        if success:
                            progress_dialog.setLabelText("正在发送消息...")
                            progress_dialog.show()

                            self.glink_thread.set_task("send_message",
                                                       message=message,
                                                       target_user=receiver)
                            self.glink_thread.send_result.connect(on_send_result)
                            self.glink_thread.start()
                        else:
                            self.show_center_message("错误", f"Glink登录失败: {login_message}",
                                                     QMessageBox.Icon.Critical)

                    def on_send_result(success, message):
                        progress_dialog.close()
                        if success:
                            self.show_center_message("成功", "消息发送成功！", QMessageBox.Icon.Information)
                        else:
                            self.show_center_message("错误", f"消息发送失败: {message}", QMessageBox.Icon.Critical)

                    self.glink_thread.set_task("login")
                    self.glink_thread.login_result.connect(on_login_result)
                    self.glink_thread.start()

            except Exception as e:
                self.show_center_message("错误", f"发送消息失败: {str(e)}", QMessageBox.Icon.Critical)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("Microsoft YaHei UI")
    app.setFont(font)

    # 确保应用在关闭最后一个窗口时不退出
    app.setQuitOnLastWindowClosed(False)

    window = TrainScheduleApp()
    window.show()
    sys.exit(app.exec())