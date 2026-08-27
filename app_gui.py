#!/usr/bin/env python3
import sys
import os
import traceback

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    log_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
    log_path = os.path.join(log_dir, "titulador_error.log")
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write("=== ERROR LOG ===\n" + err_msg + "\n")
    except Exception:
        pass
    try:
        import tkinter.messagebox
        tkinter.messagebox.showerror("Error al iniciar Titulador", f"Se produjo un error al iniciar la aplicación:\n\n{err_msg}")
    except Exception:
        pass

sys.excepthook = handle_exception

# Ensure project directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.gui_app import main

if __name__ == "__main__":
    main()
