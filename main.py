import imaplib
import email
from openai import OpenAI
import configparser
from datetime import datetime
import os

# 加载配置文件
config = configparser.ConfigParser()
config.read('.config')

# IMAP配置
IMAP_SERVER = config['EMAIL']['IMAP_SERVER']
EMAIL_ACCOUNT = config['EMAIL']['EMAIL_ACCOUNT']
EMAIL_PASSWORD = config['EMAIL']['EMAIL_PASSWORD']

# OpenAI API配置
client = OpenAI(api_key=config['OPENAI']['API_KEY'])

# 连接IMAP服务器
def connect_imap():
    print("Connecting to IMAP server...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    print("Connected to IMAP server.")
    return mail

# 拉取当前日期的邮件
def fetch_emails(mail, limit=100):
    print("Selecting inbox...")
    mail.select('inbox')
    print("Fetching all emails...")
    result, data = mail.search(None, 'ALL')
    
    if result != 'OK':
        print("No emails found.")
        return []
    
    email_ids = data[0].split()
    emails = []
    
    print(f"Found {len(email_ids)} emails.")
    
    # 分页获取邮件
    for i in range(0, len(email_ids), limit):
        batch_ids = email_ids[i:i + limit]
        print(f"Processing emails {i + 1} to {min(i + limit, len(email_ids))}...")
        
        for email_id in batch_ids:
            try:
                result, message_data = mail.fetch(email_id, '(RFC822)')
                if result != 'OK':
                    print(f"Failed to fetch email ID {email_id}.")
                    continue
                
                msg = email.message_from_bytes(message_data[0][1])
                emails.append(msg)

            except Exception as e:
                print(f"Error fetching email ID {email_id}: {e}")

        # 限制处理的邮件数量
        if len(emails) >= limit:
            break

    return emails

# 调用OpenAI API进行分类
def classify_email(content):
    print("Classifying email content...")
    response = client.chat.completions.create(
        model='gpt-3.5-turbo',
        messages=[
            {"role": "user", "content": f"Classify the following email content:\n\n{content}"}
        ]
    )
    return response.choices[0].message.content

# 主函数
def main():
    mail = connect_imap()
    emails = fetch_emails(mail, limit=10)  # 每次处理100封邮件
    
    # 打印调试信息
    if not emails:
        print("No emails to process.")
        return
    
    for msg in emails:
        subject = msg["subject"]
        # 检查内容是否存在
        if msg.is_multipart():
            content = ""
            for part in msg.walk():
                if part.get_content_type() == "text/plain" and part.get_content_disposition() != "attachment":
                    content = part.get_payload(decode=True).decode(errors='ignore')
                    break
        else:
            content = msg.get_payload(decode=True).decode(errors='ignore') if msg.get_payload() else ""

        print(f'Email Subject: {subject}')
        print(f'Email Content: {content}\n')
        
        # 如果需要进行分类，取消注释下面一行
        # classification = classify_email(content)
        # print(f'Classification: {classification}\n')

if __name__ == "__main__":
    main()
