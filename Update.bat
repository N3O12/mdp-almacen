@echo off
echo Eliminando version anterior...
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
if exist "__pycache__" rmdir /s /q __pycache__

echo Generando nueva version...
python -m PyInstaller build.py

echo Proceso completado!
pause