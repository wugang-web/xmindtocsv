import ctypes, sys, os, time

from appium import webdriver
from selenium.webdriver.common.keys import Keys
import time


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

desired_caps = {}
def admin_exe() :
    if is_admin():
        print("admin_exe函数内，以管理员权限运行")
        time.sleep(1)

        desired_caps['app'] = r"E:\soft\QuiKLAB4\QuiKLab5.0\QuiKLabIDE.exe"
        driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities=desired_caps)

        driver.find_element_by_name("登  录").click()
    else:
        if sys.version_info[0] == 3:
            print('admin_exe函数内，还没有管理员权限')
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)

# Windows 系统



if __name__ == '__main__':
    print("admin_exe前")
    os.system('runas /user:Administrator "python 新建文本文档.py"')
    print("admin_exe后")













class L:



    d