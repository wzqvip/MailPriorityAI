import imaplib
import email
from email.header import decode_header
from openai import OpenAI
import configparser

import tkinter as tk
from tkinter import ttk, messagebox

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
# 拉取指定范围内的邮件
def fetch_emails(mail, limit=3, start=0):
    mail.select('inbox')
    result, data = mail.search(None, 'ALL')
    
    if result != 'OK':
        return []
    
    email_ids = data[0].split()

    # 如果start超出了实际数量，则设置为从第一封开始
    if start >= len(email_ids):
        start = 0

    # 获取从start位置开始的limit封邮件
    email_ids = email_ids[-(start + limit):-start] if start > 0 else email_ids[-limit:]
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
def classify_email(headers, content):
    print("Classifying email content...")

    # 将邮件头和内容合并为一个字符串输入
    combined_content = f"""
    发件人: {headers['发件人']}
    收件人: {headers['收件人']}
    抄送: {headers['抄送']}
    主题: {headers['主题']}
    日期: {headers['日期']}
    内容: {content}
    """

    response = client.chat.completions.create(
        messages=[
            {
                "role": "system",
                "content": "你是一个日程智能助理，下面是一封邮件，请根据邮件的内容进行以下分类和判断：\n\n\
                    1. 判断邮件的类型（活动宣传、学校事务、学术信息、垃圾邮件、日常通知等）。\n\
                    2. 评估邮件的重要级别，包括以下几类：“必须完成”、“重要通知”、“一般通知”、“回复必要”等。 如果只是讲座或者活动或者宣传等请标记为一般通知。 如果是学校事务例如 极端天气，调课，放假，施工，是重要。\n\
                    3. 如果邮件需要回复，请总结回复的关键点。\n\
                    4. 提取日程信息，包括日期和时间（例如：xx月xx日 xx时-xx时），以及活动的地点和主题（例如，2024年11月6日 13:00 在xxx举办xxx活动）。\n\n\
                    输出格式如下，如果没有日程的话就不要添加那一行：\n\n\
                    类型: xxxx\n\
                    重要级: xxx\n\
                    发件人: xxx\n\
                    收件人(我 tacoin或者wangzq开头的就是我，或者是全体本科生/xx书院之类的)\n\
                    总结: 总结邮件内容，用几句话描述邮件的主要内容。\n\
                    日程: (例如：2024年11月6日 13:00 在xxx举办 xxx活动)"
            },
            {
                "role": "user",
                "content": combined_content
            }
        ],
        model="gpt-4o-mini",
    )
    return response.choices[0].message.content

# 主函数
# 创建GUI应用
class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("邮件助手")

        # 配置布局
        top_frame = tk.Frame(root)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        tk.Label(top_frame, text="起始邮件编号:").pack(side=tk.LEFT, padx=5)
        self.start_entry = tk.Entry(top_frame, width=5)
        self.start_entry.pack(side=tk.LEFT)
        self.start_entry.insert(0, "0")

        tk.Label(top_frame, text="邮件数量:").pack(side=tk.LEFT, padx=5)
        self.limit_entry = tk.Entry(top_frame, width=5)
        self.limit_entry.pack(side=tk.LEFT)
        self.limit_entry.insert(0, "10")

        load_button = tk.Button(top_frame, text="加载邮件", command=self.load_emails)
        load_button.pack(side=tk.LEFT, padx=10)

        # 创建TreeView表格
        columns = ("类型", "重要级", "发件人", "收件人", "总结", "日程")
        self.tree = ttk.Treeview(root, columns=columns, show="headings")
        
        # 设置列标题
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True)

    def load_emails(self):
        try:
            start = int(self.start_entry.get())
            limit = int(self.limit_entry.get())
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")
            return

        mail = connect_imap()
        emails = fetch_emails(mail, limit=limit, start=start)

        # 清空当前表格内容
        for item in self.tree.get_children():
            self.tree.delete(item)

        for msg in emails:
            headers = extract_email_headers(msg)
            content = extract_email_content(msg)
            classification = classify_email(headers, content)

            # 解析分类结果并插入到TreeView
            type_info, priority, sender, recipient, summary, schedule = self.parse_classification(classification)
            self.tree.insert("", tk.END, values=(type_info, priority, sender, recipient, summary, schedule))

    def parse_classification(self, classification):
        # 解析分类结果字符串为各个字段
        lines = classification.splitlines()
        info = {"类型": "", "重要级": "", "发件人": "", "收件人": "", "总结": "", "日程": ""}
        print("debug+++++++")
        print(lines)
        print("enddebug---")

        for line in lines:
            for key in info.keys():
                if line.startswith(f"{key}:"):
                    info[key] = line[len(key)+1:].strip()

        return info["类型"], info["重要级"], info["发件人"], info["收件人"], info["总结"], info["日程"]

# 主程序
if __name__ == "__main__":
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()