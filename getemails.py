import csv
import win32com.client
from datetime import datetime

# 初始化 Outlook 应用
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 选择收件箱
inbox = outlook.GetDefaultFolder(6)  # 6 是收件箱的索引
messages = inbox.Items

# 询问用户输入年份
year = input("请输入要导出的邮件年份（例如：2023）：").strip()

# 验证年份输入
try:
    year = int(year)
except ValueError:
    print("无效的年份输入，请输入有效的数字年份。")
    exit()

# 打开 CSV 文件以写入
filename = f"outlook_emails_{year}.csv"
with open(filename, "w", newline="", encoding="utf-8") as csvfile:
    csvwriter = csv.writer(csvfile)
    # 写入表头
    csvwriter.writerow([
        "SenderName",         # 发件人姓名
        "SenderEmailAddress", # 发件人邮箱地址
        "To",                 # 收件人邮箱地址
        "CC",                 # 抄送邮箱地址
        "Subject",            # 邮件标题
        "ReceivedTime",       # 接收时间
        "Attachments"         # 附件名称
    ])

    # 遍历所有邮件并根据年份过滤
    count = 0
    for message in messages:
        try:
            # 跳过无效邮件
            if not message:
                continue

            # 获取邮件接收时间
            received_time = message.ReceivedTime
            if received_time.year == year:
                # 安全获取每个字段，避免字段为空或异常
                sender_name = getattr(message, "SenderName", "Unknown")
                sender_email = getattr(message, "SenderEmailAddress", "Unknown")
                to = getattr(message, "To", "")
                cc = getattr(message, "CC", "")
                subject = getattr(message, "Subject", "No Subject")

                # 获取附件名称列表
                attachments = []
                if message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        attachments.append(attachment.FileName)
                attachments_str = "; ".join(attachments)  # 多个附件用分号分隔

                # 写入邮件信息到 CSV
                csvwriter.writerow([
                    sender_name,
                    sender_email,
                    to,
                    cc,
                    subject,
                    received_time,
                    attachments_str
                ])
                count += 1

        except AttributeError as e:
            # 捕获属性错误，例如某些字段不存在
            print(f"跳过邮件（无法访问某些字段）：{e}")
        except Exception as e:
            # 捕获其他未知错误
            print(f"无法处理邮件：{e}")

print(f"{count} 封来自 {year} 年的邮件已成功导出到 {filename} 文件中。")
