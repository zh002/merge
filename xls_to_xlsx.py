#转换脚本
import win32com.client as win32
import os

# 另存为xlsx的文件路径
xlsx_file = r"F:\merge\excel\xlsx"
xls_file = r"F:\merge\excel\xls"
for file in os.scandir(xls_file):

    suffix = file.name.split(".")[-1]
    if file.is_dir():
        pass
    else:
        if suffix == "xls":
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(file.path)
            # xlsx文件夹路径\\文件名x
            wb.SaveAs(xlsx_file + "\\" + file.name + "x", FileFormat=51)
            wb.Close()
