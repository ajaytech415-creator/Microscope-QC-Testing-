import sys
from PyQt6.QtWidgets import QApplication
import os

# Ensure app directory is in system path so relative imports are straightforward
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dashboard import Dashboard
import db_handler

def start_gui():
    app = QApplication(sys.argv)
    
    # Minimal polished styling
    app.setStyleSheet("""
        QMainWindow {
            background-color: #121212;
            color: #E0E0E0;
        }
        QWidget {
            font-family: 'Segoe UI', Inter, Roboto, sans-serif;
        }
        QLabel {
            font-size: 14px;
            color: #E0E0E0;
        }
        QGroupBox {
            font-size: 16px;
            font-weight: 600;
            color: #FFFFFF;
            border: 1px solid #333333;
            border-radius: 8px;
            margin-top: 15px;
            padding-top: 20px;
            background-color: #1E1E1E;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            top: -5px;
            padding: 0 5px;
            color: #4DB8FF;
        }
        QLineEdit, QComboBox, QSpinBox {
            padding: 10px;
            font-size: 14px;
            background-color: #2D2D2D;
            color: #FFFFFF;
            border: 1px solid #444444;
            border-radius: 5px;
        }
        QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
            border: 1px solid #4DB8FF;
            background-color: #333333;
        }
        QPushButton {
            font-family: 'Segoe UI', Inter, Roboto, sans-serif;
            font-weight: bold;
        }
    """)
    
    window = Dashboard()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    # Initialize the database & tables
    db_handler.init_db()
    
    # Start the PyQt6 GUI dashboard
    start_gui()

