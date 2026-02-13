@echo off
setlocal

python -m pip install --upgrade pip
pip install -r requirements_macos_app.txt pyinstaller

pyinstaller --noconfirm --clean --onefile --windowed --name SpendTagger spending_tagger_app.py

echo.
echo Build completed.
echo Output: dist\SpendTagger.exe
