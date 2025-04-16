from langchain_community.tools.gmail import get_gmail_credentials
from langchain_community.tools.gmail.utils import build_resource_service
from googleapiclient.discovery import build
import datetime
import socket
import os

def list_recent_emails(max_results=10):
    """获取最近的邮件"""
    try:
        # 设置代理环境变量
        os.environ['HTTP_PROXY'] = 'http://127.0.0.1:7890'
        os.environ['HTTPS_PROXY'] = 'http://127.0.0.1:7890'
        
        # 设置较长的超时时间
        socket.setdefaulttimeout(300)
        
        # 获取Gmail API的凭证
        credentials = get_gmail_credentials(
            token_file="token.json",
            scopes=["https://mail.google.com/"],
            client_secrets_file="credentials.json",
        )
        
        # 构建Gmail服务
        service = build('gmail', 'v1', credentials=credentials)
        
        print(f'正在获取最近的 {max_results} 封邮件...\n')
        
        # 获取邮件列表
        results = service.users().messages().list(
            userId='me',
            maxResults=max_results,
            labelIds=['INBOX']  # 只获取收件箱的邮件
        ).execute()
        
        messages = results.get('messages', [])
        
        if not messages:
            print('没有找到任何邮件。')
            return
        
        # 遍历并显示每封邮件的详细信息
        for i, message in enumerate(messages, 1):
            msg = service.users().messages().get(
                userId='me',
                id=message['id'],
                format='metadata',
                metadataHeaders=['From', 'Subject', 'Date']
            ).execute()
            
            # 获取邮件头信息
            headers = msg['payload']['headers']
            subject = next((header['value'] for header in headers if header['name'].lower() == 'subject'), '无主题')
            from_email = next((header['value'] for header in headers if header['name'].lower() == 'from'), '未知发件人')
            date = next((header['value'] for header in headers if header['name'].lower() == 'date'), '未知时间')
            
            print(f'邮件 #{i}')
            print(f'发件人: {from_email}')
            print(f'主题: {subject}')
            print(f'时间: {date}')
            print('-' * 80)
            
    except Exception as error:
        print(f'发生错误: {error}')
        print('\n提示：如果遇到连接错误，请检查：')
        print('1. 是否已经启动代理软件（如 Clash）')
        print('2. 代理端口是否正确（默认使用7890端口）')
        print('3. 网络连接是否正常')
        print('4. 是否能正常访问 Google 服务')

def main():
    # 设置要获取的邮件数量
    max_results = 10
    list_recent_emails(max_results)

if __name__ == '__main__':
    main() 