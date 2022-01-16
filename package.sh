#pyinstaller -F -w -n  e2w  -i e2w.ico main_window.py
mkdir e2w-v0.1.0
cp dist/e2w.exe e2w-v0.1.0
cp -r resources e2w-v0.1.0/
zip -q -r e2w-v0.1.0.zip  e2w-v0.1.0



