import os
import re

import winreg
jetbrains_toolbox_path = r"C:\Users\Edouard\AppData\Local\JetBrains\Toolbox"
from pathlib import Path
detected_programs = []
for d in Path(jetbrains_toolbox_path).joinpath("apps").iterdir():
    if d.is_dir():
        if d.name != "Toolbox":
            detected_programs.append(d.name)

program = "PyCharm"
login = os.getlogin()
program_path = rf"C:\Users\{login}\AppData\Local\JetBrains\Toolbox\apps\PyCharm-P"
# program_path = rf"C:\Users\{login}\AppData\Local\JetBrains\Toolbox\apps\PyCharm-P\ch-0\212.5457.59\bin\pycharm64.exe"
program_path = Path(program_path)
if program_path.exists():
    channels = {}
    for f_channel in program_path.iterdir():
        m = re.match("ch-(\d)",f_channel.name)
        if m and f_channel.is_dir():
            for dir in f_channel.iterdir():
                if dir.joinpath("bin/pycharm64.exe").exists():
                    channels[int(m.groups()[0])]=str(dir.joinpath("bin/pycharm64.exe"))
                    break
print(channels)
last_channel = max([i for i in channels.keys()])
print(last_channel)
program_path = channels[last_channel]
print(program_path)
# Remove old entries
try:
    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f'*\shell\Open in {program}\command')
    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f'*\shell\Open in {program}')
    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f'Directory\shell\Open directory in {program}\command')
    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f'Directory\shell\Open directory in {program}')
    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f'Directory\\backgroundshell\Open directory in {program}\command')
    winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f'Directory\\background\shell\Open directory in {program}')
except:
    print("Not able to delete keys")
# exit()

# File entries
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f'*\shell\Open in {program}')
winreg.SetValue(key, '', winreg.REG_SZ, f"Open in &{program}")
winreg.SetValueEx(key,"Icon",0,winreg.REG_EXPAND_SZ,f"{program_path},0")
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f'*\shell\Open in {program}\command')
winreg.SetValue(key, '', winreg.REG_SZ, f"{program_path} \"%1\"")
#Folder entries
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f'Directory\shell\Open directory in {program}')
winreg.SetValue(key, '', winreg.REG_SZ, f"Open directory in &{program}")
winreg.SetValueEx(key,"Icon",0,winreg.REG_EXPAND_SZ,f"{program_path},0")
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f'Directory\shell\Open directory in {program}\command')
winreg.SetValue(key, '', winreg.REG_SZ, f"{program_path} \"%1\"")
#Folder background entries
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f'Directory\\background\shell\Open directory in {program}')
winreg.SetValue(key, '', winreg.REG_SZ, f"Open directory in &{program}")
winreg.SetValueEx(key,"Icon",0,winreg.REG_EXPAND_SZ,f"{program_path},0")
key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f'Directory\\backgroundshell\Open directory in {program}\command')
winreg.SetValue(key, '', winreg.REG_SZ, f"{program_path} \"%1\"")

