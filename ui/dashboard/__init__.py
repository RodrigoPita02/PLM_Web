"""Dashboards (KPIs, gráficos e Evolution Dashboard)."""
from ui.dashboard.run_views import (
    render_atl_run_charts, render_atl_run_kpis,
    render_progress_charts, render_progress_kpis,
    render_sap_run_charts, render_sap_run_kpis,
)
from ui.dashboard.history import render_atl_history_charts, render_sap_history_charts
from ui.dashboard.evolution import render_evolution_dashboard

__all__ = [
    "render_sap_run_kpis", "render_sap_run_charts", "render_atl_run_kpis", "render_atl_run_charts",
    "render_progress_kpis", "render_progress_charts", "render_sap_history_charts",
    "render_atl_history_charts", "render_evolution_dashboard",
]
