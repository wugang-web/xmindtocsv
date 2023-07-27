import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


def demo_email(msg_from, passwd, msg_to, text_content, file_path=None):
    # msg_from/msg_to   发送方邮箱/收件人邮箱
    # passwd   邮箱的授权码

    msg = MIMEMultipart()

    subject = "邮件大标题"
    text_content = "邮件内容"
    text = MIMEText(text_content)
    msg.attach(text)

    if file_path:  # 附件所在的位置
        demoFile = file_path
        demoFile = MIMEApplication(open(demoFile, 'rb').read())
        demoFile.add_header('Content-Disposition', 'attachment', filename=docFile)
        msg.attach(demoFile)

    msg['Subject'] = subject
    msg['From'] = msg_from
    msg['To'] = msg_to

    try:
        s = smtplib.SMTP_SSL("smtp.qq.com", 465)
        s.login(msg_from, passwd)
        s.sendmail(msg_from, msg_to, msg.as_string())
        print("successful")
    except smtplib.SMTPException:
        print("failure")
    finally:
        s.quit()


if __name__ == '__main__':
    demo_email('发送方的邮箱', '授权码', '接收方的邮箱', '内容',
               file_path='/Users/without/Downloads/2020-11-02-18_39_103.xlsx')  # file_path 添加附件的路径，如果没有附件 file_path 写None