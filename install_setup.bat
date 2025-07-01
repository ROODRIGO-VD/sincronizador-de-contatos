@echo off
echo ============================================
echo Instalando cx_Freeze (se ainda não tiver)...
pip install cx_Freeze

echo.
echo ============================================
echo Iniciando o processo de build com cx_Freeze...
python setup.py build

echo.
echo ============================================
echo Build finalizado. Verifique a pasta "build".
pause
