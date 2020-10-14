import os
import string

ALL_FILE_PATH= "C:\\"#Office 安装盘符（这里是默认安装路径）

def GetSpreadsheetPath():
	#返回系统中SPREADSHEETCOMPARE的路径
	keyword = "SPREADSHEETCOMPARE.EXE"
	spread_sheet_path = ""
	spread_sheet_path_list = []
	
	log_file.write("spread_sheet_path_list:\n")
	for home, dirs, files in os.walk(ALL_FILE_PATH):
		for file_name in files:
			if file_name.find(keyword) != -1:
				spread_sheet_path = os.path.join(home, file_name)
				spread_sheet_path_list.append(spread_sheet_path)
				log_file.write(spread_sheet_path+'\n')

	for spread_sheet_path in spread_sheet_path_list:
		if spread_sheet_path.find("Microsoft Office\\root\\"):
			log_file.write("\nselected path:\n")
			log_file.write(spread_sheet_path+"\n")
			return spread_sheet_path

def GenCompareBat(spread_sheet_path):
	#生成bat，用于比较表格，返回Bat路径
	bat_context = f"""@echo off
chcp 65001
set toolpath=%~dp0
echo %toolpath%
echo %~1> "%toolpath%temp.txt"
echo %~2>> "%toolpath%temp.txt"
"{spread_sheet_path}" "%toolpath%temp.txt"
"""

	with open("compare.bat","w") as bat_file:
		bat_file.write(bat_context)

	log_file.write("\nsuccessfully GenBat.\npath:%s"%(os.path.abspath('.')+"\\compare.bat"))
	return os.path.abspath('.')+"\\compare.bat"

def SetRegKey(bat_path):
	#通过设置注册表，将Bat的路径设置为SVN比较用的程序
	import win32api
	import win32con
	reg_key_list = ['.xlsx','.xls']
	reg_root = win32con.HKEY_CURRENT_USER
	reg_path = r"Software\\TortoiseSVN\\DiffTools"
	reg_flags = win32con.WRITE_OWNER|win32con.KEY_WOW64_64KEY|win32con.KEY_ALL_ACCESS
	
	log_file.write("\nstart write reg\n")

	
	key_path, _ = win32api.RegCreateKeyEx(reg_root, reg_path, reg_flags)
	for reg_key in reg_key_list:
		win32api.RegSetValueEx(key_path, reg_key, 0, win32con.REG_SZ, bat_path+' %base %mine')
		#winreg.SetValue(key_path, reg_key, winreg.REG_SZ, bat_path+' %base %mine')
		log_file.write('SetValue success, key:%s\n'%reg_key)
	win32api.RegCloseKey(key_path)
	print(bat_path)


if __name__ == '__main__':
	log_file = open("installer.txt","w")
	spread_sheet_path = GetSpreadsheetPath()
	bat_path = GenCompareBat(spread_sheet_path)
	SetRegKey(bat_path)

	log_file.close()
