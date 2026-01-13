import webbrowser
import sys
import os
from threading import Timer
from main import app

def open_browser():
    webbrowser.open('http://127.0.0.1:5000/')

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        # Si es un ejecutable
        application_path = os.path.dirname(sys.executable)
    else:
        # Si es script Python
        application_path = os.path.dirname(os.path.abspath(__file__))
        
    # Configurar la ruta base
    os.chdir(application_path)
    
    # Abrir navegador después de 1.5 segundos
    Timer(1.5, open_browser).start()
    
    # Iniciar Flask
    app.run(port=5000)
