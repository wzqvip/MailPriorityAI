import imaplib
import email
from email.header import decode_header

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
    print(msg)
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and part.get_content_disposition() != "attachment":
                return part.get_payload(decode=True).decode(errors='ignore')
    else:
        return msg.get_payload(decode=True).decode(errors='ignore') if msg.get_payload() else ""
    return ""

# 解码邮件头字段
def decode_header_value(value):
    decoded_parts = decode_header(value)
    decoded_value = ''
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            decoded_value += part.decode(encoding or 'utf-8')
        else:
            decoded_value += part
    return decoded_value

# 提取邮件头信息
def extract_email_headers(msg):
    from_addr = decode_header_value(msg.get("From", ""))
    to_addr = decode_header_value(msg.get("To", ""))
    cc_addr = decode_header_value(msg.get("Cc", ""))
    subject = decode_header_value(msg.get("Subject", ""))
    date = msg.get("Date", "")

    headers = {
        "发件人": from_addr,
        "收件人": to_addr,
        "抄送": cc_addr,
        "主题": subject,
        "日期": date
    }
    return headers

# 调用OpenAI API进行分类
def classify_email(content):
    print("Classifying email content...")
    response = client.chat.completions.create(
        messages=[
        {
            "role": "system",
            "content": "你是一个日程智能助理，下面是一封邮件，请根据邮件的内容进行以下分类和判断：\n\n\
                1. 判断邮件的类型（活动宣传、学校事务、学术信息、垃圾邮件、日常通知等）。\n\
                2. 评估邮件的重要级别，包括以下几类：“必须完成”、“重要通知”、“一般通知”、“回复必要”等。\
                对于如下情况：选课提醒、学术指导、体育测试、考试信息等，标记为“必须完成”；对课程讲座、项目宣讲、活动宣传等邮件，标记为“重要通知”或“一般通知”视内容重要性而定；对于娱乐性质的活动，如音乐会、社团宣传等，标记为“一般通知”。\n\n\
                3. 如果邮件需要回复，请总结回复的关键点。\n\
                4. 提取日程信息，包括日期和时间（例如：xx月xx日 xx时-xx时），以及活动的地点和主题（例如，2024年11月6日 13:00 在xxx举办xxx活动）。\n\n\
                输出格式如下：\n\n\
                类型: xxxx\n\
                重要级: xxx\n\
                总结: 总结邮件内容，用几句话描述邮件的主要内容。\n\
                日程（如果有）: (例如：2024年11月6日 13:00 在xxx举办 xxx活动)"
        },
        {
            "role": "user",
            "content": content
        }
    ],
    model="gpt-4o-mini",
    )
    return response.choices[0].message.content

# 主函数
def main():
    mail = connect_imap()
    emails = fetch_emails(mail, limit=50)  # 获取最近的10封邮件
    
    if not emails:
        print("No emails to process.")
        return
    
    for msg in emails:
        subject = msg["subject"]
        content = extract_email_content(msg)

        # print(f'Email Subject: {subject}')
        # print(f'Email Content: {content}\n\n')
        content.encode("utf-8")

        # content = classify_email(content=content)
        print(content)
        print("\n ======================\n ")

if __name__ == "__main__":
    main()
