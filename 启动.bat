@echo off
rem Launcher fallback: delegates to the VBS entry so no PowerShell console appears
for %%f in ("%~dp0*.vbs") do start "" wscript.exe "%%f"
exit /b
