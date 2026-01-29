@echo off
echo Rebuilding EMIS Tool...
pyinstaller --onefile --noconsole emis_tool.py
echo Build Complete!
pause