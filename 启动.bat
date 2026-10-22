@echo off
rem Launch the VBS that shares this file's base name (启动.vbs), quoted for spaces.
set "ENTRY=%~dpn0.vbs"
if not exist "%ENTRY%" (
  echo Cannot find launcher VBS:
  echo %ENTRY%
  pause
  exit /b 1
)
start "" wscript.exe "%ENTRY%"
exit /b 0
