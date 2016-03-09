import sys
from cx_Freeze import setup, Executable
product_name = "QR Code Generator"
bdist_msi_options = {
    'add_to_path': False,
    'initial_target_dir': r'[ProgramFilesFolder]\%s' % (product_name),
    }

build_exe_options = {"packages": ["os"], "includes": ["Tkinter","lxml._elementpath"], "include_files" : [('template','template'),('qrcode.ico','qrcode.ico')]}

# GUI applications require a different base on Windows
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

exe = Executable(script='qrCodeGenDoc.py',
                 base=base,
                 icon='qrcode.ico',
                )

setup(name=product_name,
      version='1.0',
      description='Generates QR Codes from Excel Sheets',
      executables=[exe],
      options={
          'bdist_msi': bdist_msi_options,
          'build_exe': build_exe_options})
