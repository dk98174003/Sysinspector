cd C:\Python\sysinspector
pyinstaller --icon=sysinspector.ico --add-data "sysinspector.ico;." sysinspector.py -n Sysinspector --onefile --noconfirm --noconsole --windowed
del sysinspector.spec
rmdir /S /Q build
copy dist\*.exe .
rmdir /S /Q dist

