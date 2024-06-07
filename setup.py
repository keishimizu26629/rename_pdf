import sys
from cx_Freeze import setup, Executable

base = None

if sys.platform == 'win32':
    base = 'Win32GUI'

script = "renamePdf.py"
icon = "icons/icon.ico"


includes = [
    "os",
    "glob",
    "re",
    "math",
    "copy",
    "pdfminer",
    "csv",
    "datetime",
    "dateutil",
    "io",
    "PyPDF2",
    "reportlab",
    "tkinter",
    "PIL",
]

excludes = [
    "numpy",
    "pandas",
    "lxml",
    "locket",
    "PyQt5",
    "matplotlib",
    "babel",
    "bcrypt",
    "brotli",
    "debugpy",
    "jedi",
    "markupsafe",
    "nacl",
    "notebook",
    "pygments",
    "sphinx",
    "tornado",
    "win32com",
    "wx",
    "zmq",
    # "tkinter",
]


# sample.pyをexe化します
executables = [
    Executable(
        script=script,
        base=base,
        icon=icon  # アイコンファイルを指定します
    )
]

build_exe_options = {
    "includes": includes,  # 必要なパッケージをここに追加
    "excludes": excludes,  # 除外するパッケージをここに追加
    "include_files": [icon],  # アイコンファイルを含める
}

setup(name = 'renamePdf',
    version = '1.2',
    description = 'Convert the PDF.',
    options={"build_exe": build_exe_options},
    executables = executables
)
