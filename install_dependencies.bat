@echo off
chcp 65001 >nul
echo Instalando dependências do projeto...
winget install 9NCVDN91XZQP
pip install --upgrade pip
pip install -r requirements.txt
pause