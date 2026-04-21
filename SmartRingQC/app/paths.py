import os
import sys
from datetime import datetime

def get_base_dir():
    """
    Returns the absolute directory where the primary persistent data should live.
    When frozen by PyInstaller, sys.executable points to the .exe file natively,
    so we use its directory (e.g., where the user double-clicked it).
    During development, we use this script's directory's parent.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))

def get_asset_dir():
    """
    Returns the directory for read-only bundled assets (like icons, templates).
    In PyInstaller one-file mode, sys._MEIPASS holds the temporary extracted folder.
    In development, it's the standard file path.
    """
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))

def get_db_path():
    return os.path.join(get_base_dir(), 'database', 'qc_data.db')

def get_excel_path():
    return os.path.join(get_base_dir(), 'reports', 'inspection_log.xlsx')

def get_captures_dir():
    return os.path.join(get_base_dir(), 'captures')

def get_reports_dir():
    return os.path.join(get_base_dir(), 'reports')

def get_today_folder_name():
    """
    Returns today's dated folder name, e.g.: 18-April-2026_Friday
    """
    return datetime.now().strftime("%d-%B-%Y_%A")

def get_today_status_dir(status):
    """
    Returns the path to today's subfolder inside the given status folder.
    e.g.: captures/ACCEPTED/18-April-2026_Friday/
    status: 'ACCEPTED', 'REJECTED', or 'REWORK'
    """
    path = os.path.join(get_captures_dir(), status, get_today_folder_name())
    os.makedirs(path, exist_ok=True)
    return path
