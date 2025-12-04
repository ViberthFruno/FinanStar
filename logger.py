# Archivo: logger.py
# Ubicación: raíz del proyecto
# Descripción: Sistema de registro para mostrar mensajes en la interfaz

import tkinter as tk
import datetime


class Logger:
    def __init__(self):
        """Inicializa el sistema de registro"""
        self.text_widget = None

    def set_text_widget(self, text_widget):
        """Establece el widget de texto donde se mostrarán los logs"""
        self.text_widget = text_widget

    def log(self, message, level="INFO"):
        """Registra un mensaje con un nivel específico"""
        # Obtener la hora actual
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Formatear el mensaje de log
        log_message = f"[{now}] [{level}] {message}\n"

        # Mostrar en consola (UTF-8)
        print(log_message, end="")

        # Mostrar en el widget de texto si está disponible
        if self.text_widget:
            # Configurar el color según el nivel
            tag = f"tag_{level.lower()}"

            # Desactivar la edición
            self.text_widget.config(state=tk.NORMAL)

            # Insertar el mensaje al final
            self.text_widget.insert(tk.END, log_message, tag)

            # Configurar colores según el nivel
            if level == "ERROR":
                self.text_widget.tag_config(tag, foreground="red")
            elif level == "WARNING":
                self.text_widget.tag_config(tag, foreground="orange")
            elif level == "INFO":
                self.text_widget.tag_config(tag, foreground="blue")

            # Hacer scroll automático al final
            self.text_widget.see(tk.END)

            # Desactivar la edición nuevamente
            self.text_widget.config(state=tk.DISABLED)