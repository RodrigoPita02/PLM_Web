"""Gráficos de histórico (SAP, ATL)."""
from __future__ import annotations

import streamlit as st
import pandas as pd

from ui.dashboard.common import PLOTLY_AVAILABLE, plt, px, style_plotly

def render_sap_history_charts(logs: list[dict]) -> None:
    """
    Linha temporal do histórico SAP (Dashboard Evolution).
    
    ✅ Corrige mistura de runs (ALL vs PT01 vs PT02)
    ✅ Permite selecionar o scope
    ✅ Remove runs inválidos
    ✅ Adiciona smoothing
    """

    if not logs:
        st.info("No SAP history for charts.")
        return

    df = pd.DataFrame(logs)

    if df.empty:
        st.info("No log data available.")
        return

    # ─────────────────────────────────────────────
    # ✅ NORMALIZAÇÃO
    # ─────────────────────────────────────────────
    df["timestamp"] = pd.to_datetime(df.get("timestamp"), errors="coerce")
    df["total_errors"] = pd.to_numeric(df.get("total_errors"), errors="coerce")

    # ✅ CRUCIAL — evitar mistura de scopes
    df["plant_filter"] = df.get("plant_filter", "ALL").fillna("ALL")

    df = df.dropna(subset=["timestamp", "total_errors"])

    # ─────────────────────────────────────────────
    # ✅ SELECTOR DE SCOPE 🔥
    # ─────────────────────────────────────────────
    scope = st.selectbox(
        "🌍 Scope",
        options=["ALL", "PT01", "PT02"],
        index=0,
        key="sap_dashboard_scope"
    )

    df = df[df["plant_filter"] == scope]

    if df.empty:
        st.warning(f"No data available for scope: {scope}")
        return

    # ─────────────────────────────────────────────
    # ✅ REMOVER RUNS INVÁLIDOS
    # ─────────────────────────────────────────────
    df = df[df["total_errors"] > 100]  # evita runs bugadas

    # ─────────────────────────────────────────────
    # ✅ ORDENAR
    # ─────────────────────────────────────────────
    df = df.sort_values("timestamp")

    # ✅ REMOVER OUTLIERS EXTREMOS (fix principal)
    median = df["total_errors"].median()

    df = df[
        (df["total_errors"] > median * 0.4) & 
        (df["total_errors"] < median * 2.5)
    ]

    # ─────────────────────────────────────────────
    # ✅ SMOOTHING (trend)
    # ─────────────────────────────────────────────
    df["trend"] = df["total_errors"].rolling(3).mean()

    # ─────────────────────────────────────────────
    # ✅ GRÁFICO
    # ─────────────────────────────────────────────
    st.markdown(f"### 📈 SAP Error Trend — {scope}")

    if PLOTLY_AVAILABLE:
        import plotly.graph_objects as go

        fig = go.Figure()

        # 🔥 linha principal
        fig.add_trace(go.Scatter(
            x=df["timestamp"],
            y=df["total_errors"],
            mode="lines+markers",
            name="Errors",
            line=dict(color="#FF7200", width=3),
        ))

        # 🔵 linha trend (suave)
        fig.add_trace(go.Scatter(
            x=df["timestamp"],
            y=df["trend"],
            mode="lines",
            name="Trend",
            line=dict(color="#6EC6FF", dash="dot"),
        ))

        # 🎯 layout limpo
        fig.update_layout(
            template="simple_white",
            height=420,
            margin=dict(l=40, r=40, t=60, b=40),
            xaxis_title="Run",
            yaxis_title="Total Errors",
            showlegend=True
        )

        st.plotly_chart(fig, width='stretch')

    else:
        import matplotlib.pyplot as plt

        plt.figure(figsize=(9, 4))
        plt.plot(df["timestamp"], df["total_errors"], marker="o", label="Errors")
        plt.plot(df["timestamp"], df["trend"], linestyle="--", label="Trend")
        plt.title(f"SAP Error Trend — {scope}")
        plt.xticks(rotation=45)
        plt.legend()
        plt.tight_layout()

        st.pyplot(plt)


def render_atl_history_charts(logs: list[dict]) -> None:
    """
    Evolução temporal do ATL:
      - Curvas PT01 / PT02 com contagem de linhas ao longo do tempo.
      - Curvas de CNs com Timing Letter Number associado.
    """
    if not logs:
        st.info("No ATL history for charts.")
        return

    df = pd.DataFrame([{
        'timestamp':     l.get('timestamp'),
        'user':          l.get('user'),
        'PT01_rows':     l.get('metrics', {}).get('PT01_rows', 0),
        'PT02_rows':     l.get('metrics', {}).get('PT02_rows', 0),
        'PT01_with_TLN': l.get('metrics', {}).get('PT01_with_TLN', 0),
        'PT02_with_TLN': l.get('metrics', {}).get('PT02_with_TLN', 0),
    } for l in logs])

    try:
        df['ts'] = pd.to_datetime(df['timestamp'])
    except Exception:
        df['ts'] = df['timestamp']

    dsorted = df.sort_values('ts')

    st.markdown("### ⏱️ Temporal Evolution (ATL History)")
    if PLOTLY_AVAILABLE:
        fig = px.line(dsorted, x='ts', y=['PT01_rows', 'PT02_rows'],
                      markers=True, title="Lines by Plant over Time")
        fig = style_plotly(fig, height=400)
        st.plotly_chart(fig, width='stretch')
    else:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.plot(dsorted['ts'], dsorted['PT01_rows'], marker='o', label='PT01')
        ax.plot(dsorted['ts'], dsorted['PT02_rows'], marker='o', label='PT02')
        ax.set_title("Lines by Plant over Time")
        ax.set_xlabel("Time"); ax.set_ylabel("Lines")
        ax.legend()
        st.pyplot(fig)

    # Timing Letter Number coverage over time (only if data exists in logs)
    has_tln_history = (df['PT01_with_TLN'].sum() + df['PT02_with_TLN'].sum()) > 0
    if has_tln_history:
        st.markdown("### 🔗 Timing Letter Number Coverage over Time")
        if PLOTLY_AVAILABLE:
            fig = px.line(dsorted, x='ts', y=['PT01_with_TLN', 'PT02_with_TLN'],
                          markers=True, title="CNs with Timing Letter Number — over Time")
            fig = style_plotly(fig, height=380)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.plot(dsorted['ts'], dsorted['PT01_with_TLN'], marker='o', label='PT01 TLN')
            ax.plot(dsorted['ts'], dsorted['PT02_with_TLN'], marker='o', label='PT02 TLN')
            ax.set_title("CNs with TLN over Time")
            ax.set_xlabel("Time"); ax.set_ylabel("Count")
            ax.legend()
            st.pyplot(fig)
