import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from openai import OpenAI
import os
import sys
import configparser
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from datetime import datetime, timedelta
import time

# 加载配置文件
# 获取当前程序目录
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller创建临时文件夹，并将路径存储在 _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# 加载配置文件
config_path = resource_path('.config')
config = configparser.ConfigParser()
config.read(config_path)


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

# 拉取指定范围内的邮件
def fetch_emails(mail, limit=3, start=0, start_date=None, end_date=None):
    mail.select('inbox')

    # 按日期或数量拉取邮件
    if start_date and end_date:
        # 按日期拉取，构造 IMAP 搜索条件
        since_date = datetime.strptime(start_date, "%Y-%m-%d").strftime("%d-%b-%Y")
        before_date = datetime.strptime(end_date, "%Y-%m-%d").strftime("%d-%b-%Y")
        result, data = mail.search(None, f'SINCE {since_date}', f'BEFORE {before_date}')
    else:
        # 按数量拉取
        result, data = mail.search(None, 'ALL')

    if result != 'OK':
        return []

    email_ids = data[0].split()
    
    # 如果是按数量拉取，限制邮件数量
    if not start_date and not end_date:
        email_ids = email_ids[-limit:] if limit > 0 else email_ids

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

    date = parsedate_to_datetime(date).strftime("%Y-%m-%d %H:%M:%S") if date else ""

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

        # 进度显示标签
        self.progress_label = tk.Label(top_frame, text="")  # 初始化为空
        self.progress_label.pack(side=tk.LEFT, padx=5)

       # 起始日期和结束日期默认值
        default_start_date = (datetime.now() - timedelta(days=3)).strftime("%Y-%m-%d")
        default_end_date = datetime.now().strftime("%Y-%m-%d")

        # 起始日期输入框
        tk.Label(top_frame, text="起始日期 (YYYY-MM-DD):").pack(side=tk.LEFT, padx=5)
        self.start_date_entry = tk.Entry(top_frame, width=16)
        self.start_date_entry.insert(0, default_start_date)  # 设置默认值
        self.start_date_entry.pack(side=tk.LEFT)

        # 结束日期输入框
        tk.Label(top_frame, text="结束日期 (YYYY-MM-DD):").pack(side=tk.LEFT, padx=5)
        self.end_date_entry = tk.Entry(top_frame, width=16)
        self.end_date_entry.insert(0, default_end_date)  # 设置默认值
        self.end_date_entry.pack(side=tk.LEFT)

        # 直接加载邮件按钮
        load_by_date_button = tk.Button(top_frame, text="加载邮件", command=self.load_emails_by_date)
        load_by_date_button.pack(side=tk.LEFT, padx=5)

        # 创建TreeView表格
        columns = ("类型", "重要级", "日期", "发件人", "收件人", "总结", "日程")
        self.tree = ttk.Treeview(root, columns=columns, show="headings")

        # 设置列标题及其宽度，并绑定排序事件
        for col in columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_column(_col, False))
            if col == "类型" or col == "重要级":
                self.tree.column(col, width=80, anchor="center")
            elif col == "日期":
                self.tree.column(col, width=120, anchor="center")
            elif col == "发件人" or col == "收件人":
                self.tree.column(col, width=120, anchor="center")
            elif col == "总结":
                self.tree.column(col, width=500, anchor="w")
            elif col == "日程":
                self.tree.column(col, width=200, anchor="center")

        self.tree.pack(fill=tk.BOTH, expand=True)

        # 记录排序的状态
        self.sorting_order = {col: False for col in columns}

        # 初始化处理计数器
        self.processed_count = 0

        # 添加实时监听复选框
        self.enable_listen_var = tk.BooleanVar()
        self.enable_listen_var.set(False)
        self.listen_checkbox = tk.Checkbutton(
            top_frame, text="启用实时监听", variable=self.enable_listen_var, command=self.toggle_listen
        )
        self.listen_checkbox.pack(side=tk.LEFT, padx=5)

        # 启动监听线程的标识
        self.idle_thread = None

    # 启用或禁用监听的函数
    def toggle_listen(self):
        if self.enable_listen_var.get():
            print("实时监听已开启")
            self.start_idle_thread()
        else:
            print("实时监听已关闭")
            self.stop_idle_thread()

    # 启动监听线程
    def start_idle_thread(self):
        if self.idle_thread is None or not self.idle_thread.is_alive():
            self.idle_thread = threading.Thread(target=self.idle_mailbox, daemon=True)
            self.idle_thread.start()

    # 停止监听线程
    def stop_idle_thread(self):
        if self.idle_thread and self.idle_thread.is_alive():
            self.idle_thread = None  # 设置为 None，让 idle_mailbox 自然结束

                # 执行 IDLE 监听的函数
    def idle_mailbox(self):
        while self.idle_thread:
            try:
                # 连接到 IMAP 服务器并执行 IDLE
                mail = connect_imap()
                mail.select('inbox')
                
                # IDLE 命令等待新邮件
                mail.send(b'0001 IDLE\r\n')
                response = mail.readline()
                
                # 当有新的邮件时，服务器会通知
                if b'EXISTS' in response:
                    # 退出 IDLE 模式
                    mail.send(b'DONE\r\n')
                    mail.readline()
                    
                    # 处理新邮件
                    self.handle_new_mail()
                    
                # 等待 10 秒后重新进入 IDLE，保持连接
                time.sleep(10)

            except Exception as e:
                print("Error in idle_mailbox:", e)
                time.sleep(60)  # 出现错误后等待 60 秒再重新连接

    # 处理新邮件的函数
    def handle_new_mail(self):
        print("新邮件到达，正在处理...")
        
        # 加载最新邮件
        mail = connect_imap()
        emails = fetch_emails(mail, limit=1)  # 只加载最新的一封邮件

        # 将新邮件插入表格
        for i, msg in enumerate(emails):
            threading.Thread(target=self.process_email, args=(msg, i + 1, len(emails))).start()


    def sort_column(self, col, reverse):
        # 获取列中的所有项
        data = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]

        # 判断数据类型并排序
        try:
            data.sort(key=lambda t: float(t[0]), reverse=reverse)  # 数字排序
        except ValueError:
            data.sort(key=lambda t: t[0], reverse=reverse)  # 字符串排序

        # 清空当前表格内容
        for index, (val, item) in enumerate(data):
            self.tree.move(item, '', index)

        # 切换排序顺序
        self.sorting_order[col] = not reverse

        # 更新列标题显示排序状态
        self.tree.heading(col, text=col + (" ▲" if not reverse else " ▼"),
                          command=lambda: self.sort_column(col, not reverse))

    def load_emails(self):
        mail = connect_imap()

        # 检查是否按日期拉取
        if self.fetch_by_date.get():
            start_date = self.start_date_entry.get()
            end_date = self.end_date_entry.get()

            # 验证日期格式
            try:
                datetime.strptime(start_date, "%Y-%m-%d")
                datetime.strptime(end_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("错误", "请输入有效的日期格式 (YYYY-MM-DD)")
                return

            emails = fetch_emails(mail, start_date=start_date, end_date=end_date)
        else:
            # 按数量拉取
            try:
                start = int(self.start_entry.get())
                limit = int(self.limit_entry.get())
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字")
                return

            emails = fetch_emails(mail, limit=limit, start=start)

        # 重置处理计数器并清空表格
        self.processed_count = 0
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 使用线程处理每封邮件
        for i, msg in enumerate(emails):
            threading.Thread(target=self.process_email, args=(msg, i + 1, len(emails))).start()

    # 按日期拉取邮件的函数
    def load_emails_by_date(self):
        mail = connect_imap()

        start_date = self.start_date_entry.get()
        end_date = self.end_date_entry.get()

        # 验证日期格式
        try:
            datetime.strptime(start_date, "%Y-%m-%d")
            datetime.strptime(end_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的日期格式 (YYYY-MM-DD)")
            return

        # 获取符合日期范围的邮件
        emails = fetch_emails(mail, start_date=start_date, end_date=end_date)

        # 清空表格并重置计数器
        self.processed_count = 0
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 使用线程处理每封邮件
        for i, msg in enumerate(emails):
            threading.Thread(target=self.process_email, args=(msg, i + 1, len(emails))).start()

    def process_email(self, msg, current, total):
        headers = extract_email_headers(msg)
        content = extract_email_content(msg)
        classification = classify_email(headers, content)

        # 解析分类结果并插入到TreeView
        type_info, priority, sender, recipient, summary, schedule = self.parse_classification(classification)

        # 在主线程中更新TreeView和进度标签
        self.root.after(0, self.update_ui, type_info, priority, headers["日期"], sender, recipient, summary, schedule, current, total)

    def update_ui(self, type_info, priority, date, sender, recipient, summary, schedule, current, total):
        # 插入数据
        row_id = self.tree.insert("", 0, values=(type_info, priority, date, sender, recipient, summary, schedule))

        # 根据重要性设置行颜色
        if priority == "必须完成":
            self.tree.item(row_id, tags=('red',))
        elif priority == "重要通知":
            self.tree.item(row_id, tags=('blue',))

        # 添加样式
        # 在 TreeView 中创建标签并设置颜色
        self.tree.tag_configure('red', background='#FFC0C0')  # 淡红色
        self.tree.tag_configure('blue', background='#ADD8E6')  # 淡蓝色

        # 更新进度标签 FIXME: 这里计数有问题，修了之后可以用ifelse改成加载完成。
        self.processed_count += 1
        self.progress_label.config(text=f"加载共计{total}") 

    def parse_classification(self, classification):
        # 解析分类结果字符串为各个字段
        lines = classification.splitlines()
        info = {"类型": "", "重要级": "", "发件人": "", "收件人": "", "总结": "", "日程": ""}

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
