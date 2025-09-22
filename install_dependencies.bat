@echo off
chcp 65001 >nul
echo Instalando dependências do projeto...
winget install 9NCVDN91XZQP
pip install --upgrade pip
pip install selenium
pip install xlwings
pip install openpyxl
pip install pywin32
pip install pandas
pip install keyboard
pause