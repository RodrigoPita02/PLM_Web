"""Caminhos de persistência centralizados (relativos ao diretório de trabalho, como antes)."""
import os

HISTORY_DIR = os.environ.get("PLM_WEB_HISTORY_DIR", ".history")
REPORTS_DIR = os.environ.get("PLM_WEB_REPORTS_DIR", "reports")

ASSIGNMENTS_DIR = os.path.join(HISTORY_DIR, "assignments")
ASSIGNMENTS_MASTER_PATH = os.path.join(ASSIGNMENTS_DIR, "master.json")
OPEN_ITEMS_PATH = os.path.join(ASSIGNMENTS_DIR, "open_items.json")
ERROR_HISTORY_PATH = os.path.join(ASSIGNMENTS_DIR, "error_history.json")
