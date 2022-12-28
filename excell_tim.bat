@echo on
if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit
call C:\ProgramData\Anaconda3\Scripts\activate.bat
C:\ProgramData\Anaconda3\python.exe  "C:\Users\rafaelvilela\Desktop\MEGAsync\Code\planilha\ui2.py"
exit