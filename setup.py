import cx_Freeze
import sys
import os

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

os.environ['TCL_LIBRARY'] = r'C:\Python37\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Python37\tcl\tk8.6'

executables = [cx_Freeze.Executable("sales.py", base=base)]

packagess    = ["subprocess", "tkinter", "os", "openpyxl", "pandas", "numpy", "copy", "datetime", "xlrd", "xlsxwriter"]

include_files = [r"C:\Python37\DLLs\tcl86t.dll",
                 r"C:\Python37\DLLs\tk86t.dll",
                 r"C:\Users\Khaled\Desktop\New folder\Sales Reports Maker with Python\Database"]

cx_Freeze.setup(
    name="Mr-AbdElrahman",
    version="1.00",
    description="Sales Reports Maker",
    options={"build_exe": {"packages": packagess, "include_files": include_files}},
    executables=executables
)
