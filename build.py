import PyInstaller.__main__
import sys
import os

# Obtener la ruta del icono
icon_path = os.path.join('static', 'img', 'icon.ico')

PyInstaller.__main__.run([
    'desktop_app.py',
    '--name=AlmacenMunicipal',
    '--onefile',
    f'--icon={icon_path}',
    '--noconfirm',
    '--add-data=templates;templates',
    '--add-data=static;static',
    '--add-data=config.ini;.',
    '--hidden-import=babel.numbers',
    '--hidden-import=babel.dates',
    '--collect-data=docx',
    '--collect-data=openpyxl',
    '--collect-data=babel'
])
