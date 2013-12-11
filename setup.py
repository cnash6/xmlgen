from cx_Freeze import setup, Executable
import sys

includefiles = ['SAMPLE_FILES\SAMPLE_CSV.csv', 'SAMPLE_FILES\SAMPLE_CSV_VERY_LARGE.csv', 'SAMPLE_FILES\SAMPLE_XML.xml']

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
      options = {'build_exe': {'include_files':includefiles}}, 

      executables=[exe],
      ) 
