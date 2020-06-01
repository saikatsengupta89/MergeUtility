import sys
from cx_Freeze import setup, Executable

include_files= ['autorun.inf']
base= None

if sys.platform == "win32":
    base= "Win32GUI"

setup (
    name="WorkWithPDF",
    version="0.1",
    description="Merging Specified PDFs",
    options= {'build.exe': {'include_files':include_files}},
    executables= [Executable("MergePDFS.py", base=base)]
)
