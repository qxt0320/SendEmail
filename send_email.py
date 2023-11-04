import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# 假定 read_excel_file 和 create_grades_entry 函数的实现是存在的
from read_excel import read_excel_file
from email_content_generator import generate_email_content, create_grades_entry


# 设置邮件内容
def send_email(student_name, student_id, email_content):
    sender_email = "894790421@qq.com"  # 发送邮件的 QQ 邮箱地址
    sender_password = "ejcelpfxahtnbbhg"  # QQ 邮箱的授权码
    student_email = str(student_id) + '@fzu.edu.cn'  # 生成邮箱号

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = student_email
    message['Subject'] = '叮咚！您的成绩单来啦！'

    # 使用 HTML 格式发送
    message.attach(MIMEText(email_content, 'html'))

    # 连接到 QQ 邮箱的 SMTP 服务器并发送邮件
    mail_server = smtplib.SMTP('smtp.qq.com', 587)  # 使用 QQ 邮箱的 SMTP 服务器，端口号 587
    mail_server.starttls()  # 使用 starttls() 增加安全性
    mail_server.login(sender_email, sender_password)
    mail_server.send_message(message)
    mail_server.quit()


