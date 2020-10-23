from cx_Freeze import setup, Executable  
import os.path
import sys
#Permet d'éviter une erreur de type "KeyError: TCL_LIBRARY"
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__)) 
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6') 
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6') 
options = {     'build_exe': {         'include_files':[             os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),             os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),          ],     },}  
#Si vous souhaitez pouvoir exporter sur un autre système d'exploitation, ces lignes sont nécessaires. 
base = None 
  
#Paramètres de l'exécutable 
executables = [
        Executable(script = "strainUpdate.py",copyright= "Copyright © 2020 Christophe.navarro", icon = "fungi.ico", base = base )
]
buildOptions = dict( 
        includes = ["json","openpyxl", "requests", "scrapy", "lxml", "xlrd"],
        include_files = ["fungi.ico", "fungi.bmp", "fungi2.bmp"] 
)
 
setup(  name = "strainUpdate",
        version = "0.1" , 
        description = "Scraping site Global Catalog of MicroOrganism " , 
        options = dict(build_exe = buildOptions), 
        executables = executables)