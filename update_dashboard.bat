@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Atualizando o dashboard de indicadores...
echo.
python update_dashboard.py
echo.
echo ============================================
echo Pressione qualquer tecla para fechar esta janela.
echo ============================================
pause >nul
