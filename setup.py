from cx_Freeze import setup, Executable
import sys

##productName = "XMLGenerator"
##if 'bdist_msi' in sys.argv:
##    sys.argv += ['--initial-target-dir', 'C:\InstallDir\\' + productName]
##    sys.argv += ['--install-script', 'install.py']
##
exe = Executable(
      script="xmlgen.py",
      base="Win32GUI",
      targetName="XMLGenerator.exe"
     )
setup(
      name="XMLGenerator",
      version="2.0",
      author="Christian Nash",
      description="XML Generator for creating testing XMLs",
      executables=[exe],
      ) 
