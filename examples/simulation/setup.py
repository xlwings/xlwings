from cx_Freeze import setup, Executable


build_exe_options = {'packages': ['win32com', 'xlwings'],
                     'excludes': ['scipy', 'email', 'xml', 'pandas', 'Tkinter', 'Tkconstants', 'pydoc', 'tcl',
                                  'tk', 'matplotlib', 'PIL', 'nose', 'setuptools', 'xlrd', 'xlwt', 'PyQt4', 'markdown',
                                  'IPython', 'docutils'],
                     'optimize': 2}


setup(name = 'simulation',
      version = '0.1.0',
      options = {'build_exe': build_exe_options},
      executables = [Executable('simulation.py')])