import os
import shutil
from datetime import datetime

try:
    from app.paths import get_captures_dir, get_today_status_dir
    from app.db_handler import update_classification, undo_classification, delete_classification as db_delete
    from app.excel_writer import update_classification as xl_update, undo_classification as xl_undo, delete_classification as xl_delete
except ImportError:
    from paths import get_captures_dir, get_today_status_dir
    from db_handler import update_classification, undo_classification, delete_classification as db_delete
    from excel_writer import update_classification as xl_update, undo_classification as xl_undo, delete_classification as xl_delete


class Classifier:
    """
    Manages image classification into daily dated subfolders.

    Folder structure:
        captures/
        ├── WAITING/                         ← flat, temporary holding area
        ├── ACCEPTED/
        │   ├── 18-April-2026_Friday/
        │   ├── 19-April-2026_Saturday/
        ├── REJECTED/
        │   ├── 18-April-2026_Friday/
        └── REWORK/
            ├── 18-April-2026_Friday/
    """

    def __init__(self):
        try:
            from app.paths import get_captures_dir
        except ImportError:
            from paths import get_captures_dir

        self.base_dir = get_captures_dir()

        # WAITING is flat — no daily subfolders
        self._waiting = os.path.join(self.base_dir, 'WAITING')
        os.makedirs(self._waiting, exist_ok=True)

        # Create today's subfolders inside ACCEPTED / REJECTED / REWORK
        for status in ('ACCEPTED', 'REJECTED', 'REWORK'):
            get_today_status_dir(status)   # creates if missing

        # History stack for undo: list of (status, image_name)
        self.action_history = []

    # ----------------------------------------------------------
    # Public property — always returns the flat WAITING folder
    # ----------------------------------------------------------
    @property
    def waiting_dir(self):
        return self._waiting

    # ----------------------------------------------------------
    # Internal helpers
    # ----------------------------------------------------------
    def _move_file(self, filename, src_folder, dest_folder):
        src  = os.path.join(src_folder,  filename)
        dest = os.path.join(dest_folder, filename)
        if os.path.exists(src):
            shutil.move(src, dest)
            return True
        return False

    # ----------------------------------------------------------
    # Classification actions
    # ----------------------------------------------------------
    def accept(self, image_name, operator_id, batch_id):
        dest = get_today_status_dir('ACCEPTED')
        if self._move_file(image_name, self._waiting, dest):
            reviewed_at = update_classification(image_name, 'ACCEPTED', operator_id, batch_id)
            xl_update(image_name, 'ACCEPTED', operator_id, batch_id, '', reviewed_at)
            self.action_history.append(('ACCEPTED', image_name))
            return True
        return False

    def reject(self, image_name, operator_id, batch_id, reason):
        dest = get_today_status_dir('REJECTED')
        if self._move_file(image_name, self._waiting, dest):
            reviewed_at = update_classification(image_name, 'REJECTED', operator_id, batch_id, reason)
            xl_update(image_name, 'REJECTED', operator_id, batch_id, reason, reviewed_at)
            self.action_history.append(('REJECTED', image_name))
            return True
        return False

    def rework(self, image_name, operator_id, batch_id):
        dest = get_today_status_dir('REWORK')
        if self._move_file(image_name, self._waiting, dest):
            reviewed_at = update_classification(image_name, 'REWORK', operator_id, batch_id)
            xl_update(image_name, 'REWORK', operator_id, batch_id, '', reviewed_at)
            self.action_history.append(('REWORK', image_name))
            return True
        return False

    def undo(self):
        if not self.action_history:
            return False

        last_status, image_name = self.action_history.pop()
        # Undo looks in today's dated subfolder for the image
        src = get_today_status_dir(last_status)

        if self._move_file(image_name, src, self._waiting):
            undo_classification(image_name)
            xl_undo(image_name)
            return True
        return False

    def delete(self, image_name):
        src_path = os.path.join(self._waiting, image_name)
        if os.path.exists(src_path):
            os.remove(src_path)
            db_delete(image_name)
            xl_delete(image_name)
            return True
        return False
