# Archivo: main.py
# Ubicación: raíz del proyecto
# Descripción: Punto de entrada principal para la aplicación del bot

import tkinter as tk
from ui_manager import UIManager
import sys


def main():
    """Función principal que inicia la aplicación"""
    # Configurar codificación UTF-8 para todo el sistema
    # Verificar que stdout y stderr no sean None (pueden serlo en PyInstaller sin consola)
    try:
        if sys.stdout is not None and hasattr(sys.stdout, 'encoding'):
            if sys.stdout.encoding != 'utf-8':
                sys.stdout.reconfigure(encoding='utf-8')

        if sys.stderr is not None and hasattr(sys.stderr, 'encoding'):
            if sys.stderr.encoding != 'utf-8':
                sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        # Si hay algún error configurando encoding, continuar de todas formas
        pass

    root = tk.Tk()
    root.title("Bot de Correo")
    root.geometry("800x800")

    # Inicializar la interfaz de usuario
    app = UIManager(root)

    # Iniciar el bucle principal
    root.mainloop()


if __name__ == "__main__":
    main()