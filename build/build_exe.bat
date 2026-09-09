@echo off
setlocal
"%~dp0..\.venv\Scripts\python.exe" "%~dp0build_exe.py" %*
exit /b %errorlevel%
