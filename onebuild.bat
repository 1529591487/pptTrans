rmdir build /S /Q 
rmdir dist /S /Q 
pyinstaller  --onefile  -p "ui" -p "..\\Public" main.py