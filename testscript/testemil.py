'''发送给********邮箱.com'''
from email.mime.application import MIMEApplication
import smtplib
import os


# msg = MIMEMultipart()
#
# subject = "邮件大标题"
# text_content = "邮件内容"
# text = MIMEText(text_content)
# msg.attach(text)


def send_email(new_log):
    '''
    发送邮箱
    :param new_log: 最新的报告
    :return:
    '''

    file = open(new_log, 'rb')
    mail_content = file.read()
    # print(mail_content)
    file.close()

    # 发送方用户信息
    send_user = 'gang.wu@keliangtek.com'
    send_password = 'wugangG2023'

    # 发送和接收
    sendUser = 'gang.wu@keliangtek.com'
    receive = 'gang.wu@keliangtek.com'

    # 邮件内容
    send_subject = 'QuiLab4自动化测试报告'
    msg = MIMEApplication(mail_content, 'rb')
    msg['Subject'] = send_subject

    # msg1 = MIMEMultipart()
    text_content = "邮件内容"
    # text = MIMEText(text_content)
    # msg.attach(text)
    # msg.items(text_content)
    # print(new_log)
    msg.add_header('Content-Disposition', 'attachment', filename=new_log)
    try:

        # 登录服务器
        smt = smtplib.SMTP('imap.exmail.qq.com')

        # helo 向服务器标识用户身份
        smt.helo('imap.exmail.qq.com')
        # 服务器返回确认结果
        smt.ehlo('imap.exmail.qq.com')

        smt.login(send_user, send_password)
        print('正在准备发送邮件。')
        # print(msg.as_string())
        smt.sendmail(sendUser, receive, msg.as_string())
        smt.quit()
        print('邮件发送成功。')

    except Exception as e:
        print('邮件发送失败：', e)


def new_report(report_dir):
    '''
    获取最新报告
    :param report_dir: 报告文件路径
    :return: file ---最新报告文件路径
    '''

    # 返回指定路径下的文件和文件夹列表。
    lists = os.listdir(report_dir)
    listLog = []
    print(lists)
    for i in lists:
        if os.path.splitext(i)[1] == '.pdf':
            # print(len(i))
            # print(i)
            listLog.append(i)
    print(listLog)
    print(listLog[-1])
    fileNewLog = os.path.join(report_dir, listLog[-1])
    print(fileNewLog)
    return fileNewLog


if __name__ == '__main__':
    os.system(r"MyTest5\bin\Debug\MyTest5.exe")
    # 报告路径
    test_report = r'MyTest5\bin\Debug\Reports\SingleReportPDF\QuiKLab'
    # 获取最新测试报告
    newLog = new_report(test_report)
    # 发送邮件报告
    send_email(newLog)
    os.system("taskkill /f /im cmd.exe")
