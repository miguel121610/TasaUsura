import os
import sys

# Añadir el directorio actual al path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.gui import FormAutomationApp

if __name__ == "__main__":
    app = FormAutomationApp()
    app.run()
