import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
                           QLabel, QSpinBox, QMessageBox, QProgressBar, QDialog,
                           QTextBrowser, QSplitter, QComboBox, QSlider, QMenuBar,
                           QMenu, QInputDialog, QFileDialog, QScrollArea, QFrame,
                           QSizePolicy, QToolButton, QToolTip)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl, QSize, QRect, QPoint
from PyQt6.QtGui import (QTextDocument, QDesktopServices, QIcon, QCursor, QAction)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineProfile, QWebEngineSettings
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import datetime
import pickle
import base64
import html
import email
from email.header import decode_header
import re
from bs4 import BeautifulSoup
import json
import shutil
from datetime import datetime
import pytz

# 定义Gmail API需要的权限范围
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels'
]

def decode_base64url(data):
    """解码base64url格式的数据"""
    # 补全填充字符
    pad = len(data) % 4
    if pad:
        data += '=' * (4 - pad)
    return base64.urlsafe_b64decode(data)

def get_email_content(service, msg_id):
    """获取邮件内容的改进版本"""
    try:
        # 获取完整邮件数据
        message = service.users().messages().get(userId='me', id=msg_id, format='raw').execute()
        
        # 解码邮件原始数据
        msg_bytes = base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
        mime_msg = email.message_from_bytes(msg_bytes)
        
        # 初始化变量
        html_content = None
        plain_content = None
        image_parts = {}
        
        def process_image(part):
            """处理图片附件"""
            content_id = part.get("Content-ID")
            if content_id:
                # 移除尖括号
                content_id = content_id.strip("<>")
                image_data = part.get_payload(decode=True)
                image_type = part.get_content_type()
                # 将图片数据转换为base64
                image_b64 = base64.b64encode(image_data).decode()
                image_parts[content_id] = f"data:{image_type};base64,{image_b64}"
                return True
            return False
            
        def get_content_from_part(part):
            """从邮件部分获取内容"""
            content_type = part.get_content_type()
            if content_type.startswith('image/'):
                process_image(part)
                return None
            elif content_type == 'text/html':
                return part.get_payload(decode=True).decode()
            elif content_type == 'text/plain':
                return part.get_payload(decode=True).decode()
            return None
            
        # 遍历所有部分来获取内容
        if mime_msg.is_multipart():
            for part in mime_msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                    
                # 尝试处理图片
                if part.get_content_type().startswith('image/'):
                    process_image(part)
                    continue
                    
                # 获取文本内容
                if part.get_content_type() == 'text/html' and not html_content:
                    html_content = get_content_from_part(part)
                elif part.get_content_type() == 'text/plain' and not plain_content:
                    plain_content = get_content_from_part(part)
        else:
            content = mime_msg.get_payload(decode=True).decode()
            if mime_msg.get_content_type() == 'text/html':
                html_content = content
            else:
                plain_content = content

        # 处理HTML内容中的内嵌图片
        if html_content and image_parts:
            for cid, data_url in image_parts.items():
                html_content = html_content.replace(f'cid:{cid}', data_url)
                
        # 优先使用HTML内容
        if html_content:
            return html_content
        elif plain_content:
            # 将纯文本转换为HTML格式
            escaped_text = html.escape(plain_content)
            return f'<pre style="white-space: pre-wrap; font-family: Arial, sans-serif;">{escaped_text}</pre>'
        else:
            return "<p>此邮件没有可显示的内容</p>"
            
    except Exception as e:
        error_msg = str(e)
        return f'''
        <div style="color: #721c24; background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 4px;">
            <h3>获取邮件内容时发生错误</h3>
            <p>{html.escape(error_msg)}</p>
        </div>
        '''

class AccountManager:
    """账号管理器"""
    def __init__(self):
        self.accounts_dir = 'accounts'
        self.accounts_file = os.path.join(self.accounts_dir, 'accounts.json')
        self.current_account = None
        self.accounts = self.load_accounts()
        os.makedirs(self.accounts_dir, exist_ok=True)

    def load_accounts(self):
        """加载已保存的账号信息"""
        if os.path.exists(self.accounts_file):
            try:
                with open(self.accounts_file, 'r') as f:
                    return json.load(f)
            except:
                return {}
        return {}

    def save_accounts(self):
        """保存账号信息"""
        with open(self.accounts_file, 'w') as f:
            json.dump(self.accounts, f)

    def add_account(self, email):
        """添加新账号"""
        if not os.path.exists('credentials.json'):
            raise Exception('找不到credentials.json文件，请确保文件存在且位于正确位置')

        # 创建账号专用目录
        account_dir = os.path.join(self.accounts_dir, email)
        os.makedirs(account_dir, exist_ok=True)
        
        # 设置token文件路径
        token_path = os.path.join(account_dir, 'token.json')
        
        # 使用统一的credentials.json获取新用户的token
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        
        # 保存token
        token_data = {
            'token': creds.token,
            'refresh_token': creds.refresh_token,
            'token_uri': creds.token_uri,
            'client_id': creds.client_id,
            'client_secret': creds.client_secret,
            'scopes': creds.scopes
        }
        
        with open(token_path, 'w') as token_file:
            json.dump(token_data, token_file)
        
        # 保存账号信息
        self.accounts[email] = {
            'token_path': token_path,
            'email': email
        }
        self.save_accounts()

    def switch_account(self, email):
        """切换当前账号"""
        if email in self.accounts:
            self.current_account = email
            return True
        return False

    def get_current_credentials(self):
        """获取当前账号的Gmail凭证"""
        if not self.current_account or self.current_account not in self.accounts:
            raise Exception('请先选择一个账号')
            
        account = self.accounts[self.current_account]
        return self.get_credentials(account['token_path'])

    def get_credentials(self, token_path):
        """获取指定账号的凭证"""
        if not os.path.exists(token_path):
            raise Exception(f'找不到token文件：{token_path}')
            
        try:
            with open(token_path, 'r') as token_file:
                token_data = json.load(token_file)
                
            creds = Credentials(
                token=token_data['token'],
                refresh_token=token_data['refresh_token'],
                token_uri=token_data['token_uri'],
                client_id=token_data['client_id'],
                client_secret=token_data['client_secret'],
                scopes=token_data['scopes']
            )
            
            # 如果凭证过期且可以刷新
            if creds.expired and creds.refresh_token:
                creds.refresh(Request())
                # 更新token文件
                token_data['token'] = creds.token
                with open(token_path, 'w') as token_file:
                    json.dump(token_data, token_file)
                    
            return creds
            
        except Exception as e:
            raise Exception(f'读取token文件失败：{str(e)}')

class EmailContentDialog(QDialog):
    """邮件内容对话框"""
    def __init__(self, email_data, parent=None):
        super().__init__(parent)
        self.email_data = email_data
        self.zoom_levels = ['50%', '75%', '100%', '125%', '150%', '175%', '200%']
        self.current_zoom = 1.0  # 1.0 = 100%
        self.initUI()

    def initUI(self):
        self.setWindowTitle('邮件内容')
        self.setGeometry(200, 200, 1000, 800)
        
        layout = QVBoxLayout()
        
        # 创建工具栏
        toolbar = QHBoxLayout()
        
        # 添加缩放控制
        zoom_label = QLabel('缩放:')
        self.zoom_combo = QComboBox()
        self.zoom_combo.addItems(self.zoom_levels)
        self.zoom_combo.setCurrentText('100%')
        self.zoom_combo.currentTextChanged.connect(self.change_zoom)
        
        toolbar.addWidget(zoom_label)
        toolbar.addWidget(self.zoom_combo)
        toolbar.addStretch()
        
        layout.addLayout(toolbar)
        
        # 创建邮件头部信息
        header_html = f'''
        <div style="background-color: #f8f9fa; padding: 10px; border-radius: 5px; margin-bottom: 10px;">
            <p style="margin: 5px 0;"><b>发件人:</b> {self.email_data["from"]}</p>
            <p style="margin: 5px 0;"><b>主题:</b> {self.email_data["subject"]}</p>
            <p style="margin: 5px 0;"><b>时间:</b> {self.email_data["date"]}</p>
        </div>
        '''
        
        # 创建 WebView
        self.web_view = QWebEngineView()
        
        # 配置 WebView 设置
        profile = QWebEngineProfile.defaultProfile()
        settings = self.web_view.settings()
        
        # 启用必要的设置
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalStorageEnabled, True)
        
        # 组合头部和内容
        full_html = f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {{
                    margin: 0;
                    padding: 16px;
                    background-color: white;
                }}
                img {{
                    max-width: 100%;
                }}
                a {{
                    color: #1a73e8;
                }}
            </style>
        </head>
        <body>
            {header_html}
            <div class="email-content">
                {self.email_data["content"]}
            </div>
        </body>
        </html>
        '''
        
        # 加载内容
        self.web_view.setHtml(full_html)
        
        # 处理链接点击
        self.web_view.page().linkHovered.connect(self.handle_link_hover)
        self.web_view.page().profile().setHttpUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
        
        layout.addWidget(self.web_view)
        self.setLayout(layout)

    def change_zoom(self, zoom_level):
        """更改内容缩放级别"""
        try:
            zoom = float(zoom_level.rstrip('%')) / 100.0
            self.current_zoom = zoom
            self.web_view.setZoomFactor(zoom)
        except ValueError:
            pass

    def handle_link_hover(self, url):
        """处理链接悬停"""
        if url:
            QToolTip.showText(QCursor.pos(), url)
        else:
            QToolTip.hideText()

    def closeEvent(self, event):
        """关闭事件处理"""
        self.web_view.setHtml("")  # 清除内容
        event.accept()

class EmailCard(QFrame):
    """邮件卡片组件"""
    clicked = pyqtSignal(object)  # 发送邮件数据

    def __init__(self, email_data, parent=None):
        super().__init__(parent)
        self.email_data = email_data
        self.initUI()
        
    def initUI(self):
        self.setObjectName("emailCard")
        self.setStyleSheet("""
            #emailCard {
                background-color: white;
                border-radius: 8px;
                margin: 5px;
                padding: 10px;
            }
            #emailCard:hover {
                background-color: #f8f9fa;
                border: 1px solid #e9ecef;
            }
            QLabel {
                color: #212529;
            }
            .sender {
                font-weight: bold;
                font-size: 14px;
            }
            .subject {
                font-size: 13px;
            }
            .date {
                color: #6c757d;
                font-size: 12px;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # 发件人和时间行
        header_layout = QHBoxLayout()
        
        sender_label = QLabel(self.email_data['from'])
        sender_label.setProperty("class", "sender")
        sender_label.setWordWrap(True)
        header_layout.addWidget(sender_label)
        
        date_str = self.format_date(self.email_data['date'])
        date_label = QLabel(date_str)
        date_label.setProperty("class", "date")
        date_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        header_layout.addWidget(date_label)
        
        layout.addLayout(header_layout)
        
        # 主题行
        subject_label = QLabel(self.email_data['subject'])
        subject_label.setProperty("class", "subject")
        subject_label.setWordWrap(True)
        layout.addWidget(subject_label)
        
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

    def format_date(self, date_str):
        """格式化日期显示"""
        try:
            # 解析邮件日期
            parsed_date = email.utils.parsedate_to_datetime(date_str)
            
            # 转换为本地时间
            local_tz = datetime.now().astimezone().tzinfo
            local_date = parsed_date.astimezone(local_tz)
            
            # 获取当前时间
            now = datetime.now(local_tz)
            
            # 如果是今天的邮件，只显示时间
            if local_date.date() == now.date():
                return local_date.strftime("%H:%M")
            # 如果是昨天的邮件
            elif local_date.date() == (now - datetime.timedelta(days=1)).date():
                return "昨天 " + local_date.strftime("%H:%M")
            # 如果是今年的邮件
            elif local_date.year == now.year:
                return local_date.strftime("%m月%d日")
            # 其他情况显示完整日期
            else:
                return local_date.strftime("%Y年%m月%d日")
        except:
            return date_str

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self.email_data)

class EmailListWidget(QScrollArea):
    """邮件列表容器"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        
    def initUI(self):
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: #f8f9fa;
            }
            QScrollBar:vertical {
                border: none;
                background: #f8f9fa;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #dee2e6;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #adb5bd;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)
        
        # 创建容器widget
        container = QWidget()
        self.layout = QVBoxLayout(container)
        self.layout.setSpacing(2)
        self.layout.setContentsMargins(10, 10, 10, 10)
        self.layout.addStretch()
        
        self.setWidget(container)
    
    def clear(self):
        """清空列表"""
        while self.layout.count() > 1:  # 保留最后的stretch
            item = self.layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
    
    def add_email(self, email_data):
        """添加邮件卡片"""
        card = EmailCard(email_data)
        self.layout.insertWidget(self.layout.count() - 1, card)
        return card

class ModernButton(QPushButton):
    """现代风格按钮"""
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
            QPushButton:disabled {
                background-color: #6c757d;
            }
        """)

class EmailFetcher(QThread):
    """邮件获取线程"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, account_manager, max_results):
        super().__init__()
        self.account_manager = account_manager
        self.max_results = max_results

    def run(self):
        try:
            os.environ['HTTP_PROXY'] = 'http://127.0.0.1:7890'
            os.environ['HTTPS_PROXY'] = 'http://127.0.0.1:7890'

            # 使用当前账号的凭证获取服务
            creds = self.account_manager.get_current_credentials()
            service = build('gmail', 'v1', credentials=creds)

            results = service.users().messages().list(
                userId='me',
                maxResults=self.max_results,
                labelIds=['INBOX']
            ).execute()

            messages = results.get('messages', [])
            email_list = []
            
            # 减少进度条更新频率，每处理 5 封邮件更新一次
            update_interval = max(1, len(messages) // 10)  # 将总进度分成10份
            last_update = 0

            for i, message in enumerate(messages):
                # 只在达到更新间隔时才发送进度信号
                if i - last_update >= update_interval:
                    self.progress.emit(int((i + 1) / len(messages) * 100))
                    last_update = i

                msg = service.users().messages().get(
                    userId='me',
                    id=message['id'],
                    format='metadata',
                    metadataHeaders=['From', 'Subject', 'Date']
                ).execute()

                headers = msg['payload']['headers']
                subject = next((header['value'] for header in headers if header['name'].lower() == 'subject'), '无主题')
                from_email = next((header['value'] for header in headers if header['name'].lower() == 'from'), '未知发件人')
                date = next((header['value'] for header in headers if header['name'].lower() == 'date'), '未知时间')

                content = get_email_content(service, message['id'])

                email_list.append({
                    'id': message['id'],
                    'from': from_email,
                    'subject': subject,
                    'date': date,
                    'content': content
                })

            # 确保最后显示100%
            self.progress.emit(100)
            self.finished.emit(email_list)

        except Exception as e:
            self.error.emit(str(e))

class GmailClient(QMainWindow):
    def __init__(self):
        super().__init__()
        self.account_manager = AccountManager()
        self.email_list = []
        self.initUI()
        self.apply_style()

    def initUI(self):
        self.setWindowTitle('Gmail 客户端')
        self.setGeometry(100, 100, 1200, 800)

        # 创建菜单栏
        self.create_menu_bar()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)

        # 顶部工具栏
        toolbar = QWidget()
        toolbar.setObjectName("toolbar")
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(20, 10, 20, 10)

        # 当前账号显示
        self.account_label = QLabel('当前账号: 未选择')
        self.account_label.setObjectName("accountLabel")
        toolbar_layout.addWidget(self.account_label)

        toolbar_layout.addStretch()

        # 邮件数量选择
        count_layout = QHBoxLayout()
        count_layout.setSpacing(5)
        count_label = QLabel('显示数量:')
        self.count_spinner = QSpinBox()
        self.count_spinner.setRange(1, 100)
        self.count_spinner.setValue(20)
        count_layout.addWidget(count_label)
        count_layout.addWidget(self.count_spinner)
        toolbar_layout.addLayout(count_layout)

        # 刷新按钮
        refresh_btn = ModernButton('刷新邮件')
        refresh_btn.clicked.connect(self.fetch_emails)
        toolbar_layout.addWidget(refresh_btn)

        layout.addWidget(toolbar)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 邮件列表
        self.email_list_widget = EmailListWidget()
        layout.addWidget(self.email_list_widget)

    def apply_style(self):
        """应用全局样式"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
            }
            #toolbar {
                background-color: white;
                border-bottom: 1px solid #dee2e6;
            }
            #accountLabel {
                font-size: 14px;
                color: #495057;
            }
            #progressBar {
                border: none;
                background-color: #e9ecef;
                height: 2px;
                margin: 0px;
            }
            #progressBar::chunk {
                background-color: #007bff;
            }
            QLabel {
                color: #212529;
            }
            QSpinBox {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 4px;
                background: white;
            }
            QSpinBox:hover {
                border-color: #80bdff;
            }
            QMenuBar {
                background-color: white;
                border-bottom: 1px solid #dee2e6;
            }
            QMenuBar::item {
                padding: 8px 12px;
                background-color: transparent;
            }
            QMenuBar::item:selected {
                background-color: #e9ecef;
            }
            QMenu {
                background-color: white;
                border: 1px solid #dee2e6;
                padding: 5px;
            }
            QMenu::item {
                padding: 8px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #e9ecef;
            }
            QMenu::separator {
                height: 1px;
                background-color: #dee2e6;
                margin: 5px 0px;
            }
            QMenu::item:checked {
                background-color: #e9ecef;
            }
            QMenu::indicator {
                width: 16px;
                height: 16px;
                margin-right: 8px;
            }
            QMenu::indicator:checked {
                image: url(check.png);  /* 如果有选中图标的话 */
            }
        """)

    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 账号菜单
        self.account_menu = menubar.addMenu('账号')
        self.update_account_menu()

    def update_account_menu(self):
        """更新账号菜单"""
        self.account_menu.clear()
        
        # 添加账号选项
        add_account = QAction('添加新账号...', self)
        add_account.triggered.connect(self.add_account)
        self.account_menu.addAction(add_account)
        
        if self.account_manager.accounts:
            # 添加分隔线
            self.account_menu.addSeparator()
            
            # 添加已登录的账号列表
            for email in self.account_manager.accounts.keys():
                account_action = QAction(email, self)
                account_action.setCheckable(True)
                account_action.setChecked(email == self.account_manager.current_account)
                account_action.triggered.connect(lambda checked, e=email: self.switch_account(e))
                self.account_menu.addAction(account_action)

    def add_account(self):
        """添加新账号"""
        email, ok = QInputDialog.getText(self, '添加账号', '请输入Gmail邮箱地址:')
        if ok and email:
            if not email.endswith('@gmail.com'):
                QMessageBox.warning(self, '警告', '请输入有效的Gmail邮箱地址')
                return
                
            try:
                self.account_manager.add_account(email)
                self.account_manager.switch_account(email)
                self.update_account_label()
                self.update_account_menu()  # 更新账号菜单
                self.fetch_emails()
                QMessageBox.information(self, '成功', f'账号 {email} 添加成功！')
            except Exception as e:
                QMessageBox.critical(self, '错误', f'添加账号失败：{str(e)}')

    def switch_account(self, email):
        """切换账号"""
        try:
            self.account_manager.switch_account(email)
            self.update_account_label()
            self.update_account_menu()  # 更新账号菜单
            self.fetch_emails()
        except Exception as e:
            QMessageBox.critical(self, '错误', f'切换账号失败：{str(e)}')

    def update_account_label(self):
        """更新当前账号显示"""
        account_name = self.account_manager.current_account or '未选择'
        self.account_label.setText(f'当前账号: {account_name}')

    def fetch_emails(self):
        """获取邮件列表"""
        if not self.account_manager.current_account:
            QMessageBox.warning(self, '警告', '请先选择一个账号')
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        self.thread = EmailFetcher(self.account_manager, self.count_spinner.value())
        self.thread.progress.connect(self.update_progress)
        self.thread.finished.connect(self.display_emails)
        self.thread.error.connect(self.show_error)
        self.thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def display_emails(self, email_list):
        self.email_list = email_list
        self.email_list_widget.clear()
        
        for email_data in email_list:
            card = self.email_list_widget.add_email(email_data)
            card.clicked.connect(self.show_email_content)

        self.progress_bar.setVisible(False)

    def show_error(self, error_msg):
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, '错误', f'获取邮件时发生错误：\n{error_msg}')

    def show_email_content(self, email_data):
        dialog = EmailContentDialog(email_data, self)
        dialog.exec()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格作为基础
    client = GmailClient()
    client.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main() 