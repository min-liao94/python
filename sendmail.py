import os
import pandas as pd
import shutil
from openpyxl import load_workbook
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


# 取得當前腳本所在目錄
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# 讀取 list_all.xlsx 檔案
list_file = "list_all.xlsx"

# 讀取 Excel 檔案
try:
    df = pd.read_excel(list_file)

except Exception as e:
    print(f"讀取檔案錯誤：{e}")
    exit(1)

# 讀取 Excel 檔案所需的資訊內容取代template_資訊問卷
#template_file ="template_資訊問卷.txt"
first_col = df.iloc[:, 0] # 工號
second_col = df.iloc[:, 1] # 姓名
third_col = df.iloc[:, 2] # 密碼
fourth_col = df.iloc[:, 3] # 信箱
subject = "2025年資訊發展中心服務滿意度調查(請同仁協助)"
# 寄件設定
sender_email = "noreply@ivt.tw"
smtp_password = "ahpei8jaeThoJo8t"
smtp_host = "mail10.ivt.tw"
smtp_port = 465

def build_body(name, account, password):
    """建立郵件內容"""
    return (
        f"Dear {name} 您好：\n\n"
        "未完成填寫的同仁們請協助填寫表單，歡迎將您的意見傳達給資訊發展中心。\n"
        "若已填完表單者，可忽略此信件，謝謝。\n"
        "此滿意度調查將於12/10日 17:00 進行收件\n"
        "請同仁依操作說明操作\n"
        "若有任何問題，隨時可聯繫資訊發展中心工程師怡琳 min (分機175)。\n\n"
        "感謝您的協助！\n\n"
        "Regards,\n"
        "資訊發展中心"
    )

def find_attachment(gonghao):
    base_name = "解決方案"
    file_dir = os.path.join(script_dir, "file")
    for ext in (".pdf",):
        candidate = os.path.join(file_dir, base_name + ext)
        if os.path.exists(candidate):
            return candidate
    return None

missing_attachments = []

for i in range(len(df)):
    gonghao = str(first_col.iloc[i]).strip()
    name = str(second_col.iloc[i]).strip()
    user_password = str(third_col.iloc[i]).strip()
    email = str(fourth_col.iloc[i]).strip()

    if not email or email.lower() in ["nan", "none"]:
        print(f"跳過：工號 {gonghao} 缺少 email")
        continue

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = email
    msg["Subject"] = subject
    msg.attach(MIMEText(build_body(name, gonghao, user_password), "plain", "utf-8"))

    # 附加 file 資料夾中的 PDF：問卷教學手冊.pdf
    attachment_path = find_attachment(gonghao)
    if not attachment_path:
        warning = (
            f"警告：找不到附件 file/解決方案.pdf（工號：{gonghao}），"
            "已跳過寄信並請儘速補齊檔案。"
        )
        print(warning)
        missing_attachments.append(gonghao)
        continue

    with open(attachment_path, "rb") as f:
        attach_part = MIMEApplication(f.read())
        attach_part.add_header(
            "Content-Disposition",
            "attachment",
            filename=os.path.basename(attachment_path),
        )
        msg.attach(attach_part)

    try:
        with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
            server.login(sender_email, smtp_password)
            server.sendmail(sender_email, [email], msg.as_string())
            print(f"郵件已成功寄出至 {email}")
    except Exception as e:
        print(f"寄信失敗（{email}）：{e}")


if missing_attachments:
    print(
        "提醒：以下工號因缺少附件而未寄出，請檢查 file/解決方案.pdf 是否存在："
    )
    print(", ".join(missing_attachments))

