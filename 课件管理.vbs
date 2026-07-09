Set WshShell = CreateObject("WScript.Shell")

' Project directory (hardcoded to avoid wrong-path issues)
projectDir = "C:\Users\14005\WorkBuddy\2026-07-03-17-01-41"
pyPath = "C:\Users\14005\.workbuddy\binaries\python\envs\xiaoshuzhai\Scripts\python.exe"

' Set UTF-8 encoding for Python I/O
WshShell.Environment("Process").Item("PYTHONIOENCODING") = "utf-8"

' Launch menu in CMD window
WshShell.Run "cmd /c cd /d """ & projectDir & """ && chcp 65001 >nul && """ & pyPath & """ ppt_manager.py menu", 1, True
