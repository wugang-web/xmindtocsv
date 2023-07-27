from appium import webdriver
from selenium.webdriver.common.keys import Keys
import time


import ctypes, sys

import ctypes
import sys

#
# def is_admin():
#     try:
#         return ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         return False


# def elevate():
#     if not is_admin():
#         # 提升权限
#         ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
#         sys.exit()



class App():
    def __init__(self):
        # 检查并提升管理员权限
        desired_caps = {}
        # do the things you want.
        desired_caps['app'] = r"E:\soft\QuiKLAB4\QuiKLab5.0\QuiKLab.exe"

        # desired_caps['app'] = r"C:\Windows\System32\notepad.exe"

        driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities=desired_caps)
        driver.wait_activity(3)

        # wait = WebDriverWait(driver, 30)
        #driver.find_element_by_accessibility_id("LoginDialogEx.widget_pic_center.widget_user_login.pushButtonLogin").click()
        # driver.find_element_by_accessibility_id("LoginDialogEx.widget_top.widget_toolbar.pushButton_close").click()
        # time.sleep(10)
        driver.find_element_by_name("文件").click()
        driver.find_element_by_name("新建项目").click()
        # driver.find_element_by_accessibility_id("LoginDialogEx.widget_pic_center.widget_user_login.pushButtonLogin").click()
App()
# Re-run the program with admin rights
# ctypes.windll.shell32.ShellExecuteW(None,"runas", sys.executable, __file__, None, 1)
