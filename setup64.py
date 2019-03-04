import cx_Freeze
import sys
import os

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

os.environ['TCL_LIBRARY'] = r'Files\64\tcl8.6'
os.environ['TK_LIBRARY'] = r'Files\64\tk8.6'

shortcut_table = [
    ("DesktopShortcut",        # Shortcut
     "DesktopFolder",          # Directory_
     "ElManar-Office",           # Name
     "TARGETDIR",              # Component_
     "[TARGETDIR]ElManar-Office.exe",# Target
     None,                     # Arguments
     None,                     # Description
     None,                     # Hotkey
     None,                     # Icon
     None,                     # IconIndex
     None,                     # ShowCmd
     'TARGETDIR'               # WkDir
     )
    ]

msi_data = {"Shortcut": shortcut_table}
bdist_msi_options = {'data': msi_data}

executables = [cx_Freeze.Executable("ElManar-Office.py", base=base, icon='icon.ico')]

packagess    = ["subprocess", "tkinter", "os", "openpyxl", "pandas", "numpy", "copy", "xlrd", "xlsxwriter", "datetime"]

include_files = [r"Files\64\tcl86t.dll",
                 r"Files\64\tk86t.dll",
                 r"Database",
                 r"icon.ico"]

cx_Freeze.setup(
    name="ElManarOffice",
    version="1.00",
    description="ElManar-Office Program designed by Khaled Mohamed",
    options={"bdist_msi": bdist_msi_options, "build_exe": {"packages": packagess, "include_files": include_files}},
    executables=executables
)
