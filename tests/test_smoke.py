"""Smoke tests: todos os módulos importam e todas as páginas renderizam sem exceções."""
import importlib
import pkgutil
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]

PAGES = [
    "📈 Evolution Dashboard", "📦 SAP Material Validation", "📦 Indirect Materials Report",
    "🚗 PLM Workflow", "📊 SAP Validation Progress", "🧹 Data Cleanups", "📋 SSA Validation",
]


def test_all_modules_import():
    import engines, services, ui, views
    for pkg in (engines, services, ui, views):
        for m in pkgutil.walk_packages(pkg.__path__, pkg.__name__ + "."):
            importlib.import_module(m.name)


@pytest.mark.parametrize("page", PAGES + ["PLM:Change Cockpit"])
def test_page_renders(page):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60).run()
    if page.startswith("PLM:"):
        at.sidebar.radio[0].set_value("🚗 PLM Workflow").run()
        at.sidebar.radio[1].set_value("Change Cockpit").run()
    else:
        at.sidebar.radio[0].set_value(page).run()
    assert not at.exception, [str(e.value) for e in at.exception]


def test_evolution_dashboard_with_history(tmp_path, monkeypatch):
    """Evolution Dashboard com histórico sintético, em todos os temas."""
    import json
    from datetime import datetime, timedelta
    from streamlit.testing.v1 import AppTest
    from ui.theme import THEMES

    d = tmp_path / ".history" / "sap"
    d.mkdir(parents=True)
    for i in range(8):
        ts = (datetime(2026, 10, 1) + timedelta(days=3 * i)).isoformat(timespec="seconds")
        (d / f"SAP_VALIDATION_{i:03d}.json").write_text(json.dumps({
            "run_type": "SAP_VALIDATION", "timestamp": ts, "user": "tester", "plant_filter": "ALL",
            "metrics": {"total_errors": 500 - i * 20, "fields_with_error": 12, "materials_with_error": 90},
            "top_fields": [{"Field": "MRP Type", "Nº of Errors": 10}],
        }))
    monkeypatch.chdir(tmp_path)
    for theme in THEMES:
        at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=90).run()
        at.sidebar.radio[0].set_value("📈 Evolution Dashboard").run()
        at.selectbox[0].set_value(theme).run()
        assert not at.exception, (theme, [str(e.value) for e in at.exception])


def test_excel_polish_keeps_report_valid():
    import io
    from openpyxl import Workbook, load_workbook
    from engines.common.excel_polish import polish_workbook

    wb = Workbook()
    ws = wb.active; ws.title = "Executive Summary"; ws["A1"] = "Title"
    d = wb.create_sheet("Field With Error")
    d.append(["Material", "Field"]); d.append(["1", "MRP Type"])
    polish_workbook(wb)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    wb2 = load_workbook(buf)
    assert wb2["Field With Error"].auto_filter.ref == "A1:B2"
    assert wb2["Executive Summary"].sheet_view.showGridLines is False
