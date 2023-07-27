

from appium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import subprocess

# CLIENT_LOCATION 是定义的一个存放地址的属性
# subprocess.Popen([CLIENT_LOCATION])

import ctypes, sys

import os
# -*- coding:utf-8 -*-
# import sys, ctypes, os
#
#
# def __set_run_as_admin():
#     def is_admin():
#         try:
#             return ctypes.windll.shell32.IsUserAnAdmin()
#         except:
#             return False
#
#     if is_admin():
#         None
#     else:
#         if sys.version_info[0] == 3:
#             ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
#             # 想要执行的代码


class NotepadTests():


    # Windows 系统
    #os.system('runas /user:Administrator "E:\soft\QuiKLAB4\QuiKLab5.0\QuiKLabIDE.exe"')



    time.sleep(4)
    def setUpClass(self):
        desired_caps = {}
        #os.system(r'runas /user:Administrator  start "" /d "E:\soft\QuiKLAB4\QuiKLab5.0\"  "QuiKLabIDE.exe"')
        #subprocess.Popen(["E:\soft\QuiKLAB4\QuiKLab5.0\QuiKLabIDE.exe"])
        #desired_caps['app'] = r"E:\soft\QuiKLAB4\QuiKLab5.0\QuiKLabIDE.exe"

        desired_caps['app'] = r"E:\soft\QuiKLAB4\QuiKLab5.0\QuiKLabIDE.exe"

        #desired_caps['app'] = r"C:\Windows\System32\notepad.exe"
        self.driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities=desired_caps)


    #@classmethod
    def tearDownClass(self):
        self.driver.quit()

    def test_edit(self):
        self.driver.find_element_by_name("登  录").click()
        time.sleep(4)
        self.driver.find_element_by_name("文件").click()
        time.sleep(4)
        self.driver.find_element_by_xpath(
            '//MenuItem[starts-with(@Name, "保存(S)")]').click()
        time.sleep(4)
        self.driver.find_element_by_xpath(
            '//Pane[starts-with(@ClassName, "Address Band Root")]').find_element_by_xpath(
            '//ProgressBar[starts-with(@ClassName, "msctls_progress32")]').click()
        time.sleep(4)
        self.driver.find_element_by_xpath(
            '//Edit[starts-with(@Name, "地址")]').send_keys(
            r"D:\test" + Keys.ENTER)
        self.driver.find_element_by_accessibility_id(
            'FileNameControlHost').send_keys("note_test.txt")
        self.driver.find_element_by_name("保存(S)").click()
        self.driver.find_element_by_name("关闭").click()


if __name__ == "__main__":
    #__set_run_as_admin()
    B= NotepadTests()
    B.setUpClass()
    B.test_edit()
    #suite = unittest.TestLoader().loadTestsFromTestCase(NotepadTests)
    #unittest.TextTestRunner(verbosity=2).run(suite)