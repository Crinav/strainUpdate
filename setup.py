from cx_Freeze import setup, Executable  
import os.path

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
        includes = ["datetime","openpyxl", "requests", "scrapy", "lxml", "xlrd","re", "os", ],
        include_files = ["fungi.ico", "LICENSE.txt", "README.txt", "user.txt"] 
)
 
setup(  name = "strainUpdate",
        version = "1.0" , 
        description = "Scraping site Global Catalog of MicroOrganism " , 
        options = dict(build_exe = buildOptions), 
        executables = executables)