"""Vista: 📈 Evolution Dashboard."""

from ui.dashboard import render_evolution_dashboard
from ui.layout import render_footer


def render(start_time: float) -> None:
    render_evolution_dashboard()
    render_footer(start_time, version="v1.4")
