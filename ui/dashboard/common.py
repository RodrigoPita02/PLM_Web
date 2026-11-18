"""Utilitários partilhados pelos dashboards: paleta, estilo Plotly, cores."""
from __future__ import annotations



# Plotly é o motor preferido; Matplotlib é o fallback quando o Plotly não está instalado.
try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
    plt = None
except Exception:
    PLOTLY_AVAILABLE = False
    px = None
    import matplotlib.pyplot as plt


# --- COLOR UTILITIES ---
def hex_to_rgba(hex_color: str, alpha: float) -> str:
    hc = hex_color.lstrip("#")
    r = int(hc[0:2], 16)
    g = int(hc[2:4], 16)
    b = int(hc[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


# Palete oficial
VISTEON_COLORS = ["#FF7200", "#6EC6FF", "#00A6A6", "#4D4D4D", "#D9D9D9"]


# central opacity keys if you kept COLOR_OPACITY
COLOR_OPACITY = {
    'spark':       0.12,  # a touch stronger on dark
    'sap_area':    0.14,  # main SAP area visibility in dark
    'anomaly_band':0.12   # uncertainty band
}


def style_plotly(fig, height=450):
    """Aplica estilo padrão Visteon a qualquer gráfico."""
    fig.update_layout(
        template="simple_white",
        height=height,
        margin=dict(l=40, r=40, t=80, b=80),
        title_x=0.5,
        font=dict(size=14),
        xaxis_title_font=dict(size=14),
        yaxis_title_font=dict(size=14),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
        colorway=VISTEON_COLORS
    )
    return fig
