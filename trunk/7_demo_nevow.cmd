set python_home=c:\python24
set main=pyconsole_nevow.py
set pythonpath=%~dp0
start %python_home%\python.exe %python_home%\Scripts\twistd.py -noy %~dp0\%main%
