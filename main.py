import imaplib
import email
from openai import OpenAI
import configparser

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
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    return mail

# 拉取最近的邮件
def fetch_emails(mail, limit=10):
    mail.select('inbox')
    result, data = mail.search(None, 'ALL')
    
    if result != 'OK':
        return []
    
    email_ids = data[0].split()
    email_ids = email_ids[-limit:]  # 获取最近的邮件ID
    emails = []

    for email_id in email_ids:
        try:
            result, message_data = mail.fetch(email_id, '(RFC822)')
            if result != 'OK':
                continue
            
            msg = email.message_from_bytes(message_data[0][1])
            emails.append(msg)

        except Exception as e:
            print(f"Error fetching email ID {email_id}: {e}")

    return emails

# 提取邮件内容
def extract_email_content(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and part.get_content_disposition() != "attachment":
                return part.get_payload(decode=True).decode(errors='ignore')
    else:
        return msg.get_payload(decode=True).decode(errors='ignore') if msg.get_payload() else ""
    return ""

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
    emails = fetch_emails(mail, limit=10)  # 获取最近的10封邮件
    
    if not emails:
        print("No emails to process.")
        return
    
    for msg in emails:
        subject = msg["subject"]
        content = extract_email_content(msg)

        # print(f'Email Subject: {subject}')
        print(f'Email Content: {content}\n\n')

        # content = classify_email(content="2333,322")
        # print(content)

if __name__ == "__main__":
    main()
