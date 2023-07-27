import sys
import clr

# make Ranorex module available, needs before the `import Ranorex`
sys.path.append('E:\\Ranorex\\Bin\\x64\\')
clr.AddReference('Ranorex.Core')
import Ranorex

Ranorex.Host.Local.RunApplication('D:\BaiduNetdiskDownload\Notepad++.7.6.1.bin.x64\notepad++.exe')

# apps = [c for c in Ranorex.Host.Local.Children if "My App" in c.ToString()]
# if len(apps) != 1:
#     print("starting of 'My App' somehow failed, quitting now")
#     sys.exit(1)
#
# app = apps[0]
# app.PressKeys('{LMenu down}{Fkey}{LMenu up}') # presses Alt-F -> e.g. opens the file menu