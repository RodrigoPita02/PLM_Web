
# utils/dashboard.py
# -*- coding: utf-8 -*-
"""
Dashboards/KPIs para as três páginas:
 - SAP Material Validation
 - PLM Workflow (ATL)
 - SAP Material Validation Progress (Diff)

Design:
- Tenta usar Plotly; se não estiver instalado, faz fallback para Matplotlib.
- As funções não alteram os DataFrames originais (só leitura/plot).
"""

from __future__ import annotations
import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime
from scipy.stats import linregress
from sap_validation_engine import list_run_logs

# Tenta usar Plotly; se não houver, faz fallback para Matplotlib
try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except Exception:
    PLOTLY_AVAILABLE = False
    # Importa matplotlib apenas se necessário (fallback)
    import matplotlib.pyplot as plt

# Palete oficial
VISTEON_COLORS = ["#FF7200", "#6EC6FF", "#00A6A6", "#4D4D4D", "#D9D9D9"]

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

# =========================
#   SAP MATERIAL VALIDATION
# =========================

def render_sap_run_kpis(errors_df: pd.DataFrame) -> None:
    """
    Mostra KPIs básicos da execução atual de SAP:
      - Total de erros
      - Nº de fields distintos com erro
      - Nº de materiais com erro
      - Nº de erros justificados por View em falta
    """
    if errors_df is None or errors_df.empty:
        st.info("No data for KPIs")
        return

    total_errors = int(errors_df.shape[0])

    total_fields = int(errors_df['Field'].nunique()) if 'Field' in errors_df.columns else 0
    total_materials = int(errors_df['Material'].nunique()) if 'Material' in errors_df.columns else 0
    missing_views_count = int(errors_df['Justified_By_Missing_View'].sum()) \
        if 'Justified_By_Missing_View' in errors_df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Errors", f"{total_errors:,}")
    c2.metric("Fields with errors (distinct)", f"{total_fields:,}")
    c3.metric("Materials with errors (distinct)", f"{total_materials:,}")
    c4.metric("Justified errors (missing View)", f"{missing_views_count:,}")

def render_sap_run_charts(errors_df: pd.DataFrame) -> None:
    """
    Gráficos da execução atual (SAP):
      - Top N Fields com mais erros
      - Erros por Planta
      - Top N Materiais com mais erros
      - Materiais com View em falta (por código de View)
    """
    if errors_df is None or errors_df.empty:
        return

    st.markdown("### 📊 Charts — Current Execution (SAP)")

    # Top-N configurável via coluna auxiliar (se existir), senão 20
    top_n = int(errors_df['_TOP_N'].iloc[0]) if '_TOP_N' in errors_df.columns else 20

    # --- Top N Fields ---
    if 'Field' in errors_df.columns:
        top_fields = (
            errors_df['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Errors')
            .head(top_n)
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(top_fields, x='Field', y='Errors',
                         title=f"Top {top_n} Fields with more errors", text='Errors')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(top_fields['Field'], top_fields['Errors'], color="#1f77b4")
            ax.set_title(f"Top {top_n} Fields with more errors")
            ax.set_xlabel("Field"); ax.set_ylabel("Errors")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

    # --- Erros por Planta ---
    if 'Plant' in errors_df.columns:
        by_plant = (
            errors_df['Plant'].astype(str).str.strip().replace({'': 'Without Plant'})
            .value_counts()
            .rename_axis('Plant')
            .reset_index(name='Errors')
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_plant, x='Plant', y='Errors',
                         title="Errors by Plant", text='Errors', color='Plant')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_plant['Plant'], by_plant['Errors'], color="#ff7f0e")
            ax.set_title("Errors by Plant")
            ax.set_xlabel("Plant"); ax.set_ylabel("Errors")
            st.pyplot(fig)

    # --- Top N Materiais ---
    if 'Material' in errors_df.columns:
        by_mat = (
            errors_df['Material'].astype(str).str.strip()
            .value_counts()
            .rename_axis('Material')
            .reset_index(name='Errors')
            .head(top_n)
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_mat, x='Material', y='Errors',
                         title=f"Top {top_n} Materials with the most errors", text='Errors')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_mat['Material'], by_mat['Errors'], color="#2ca02c")
            ax.set_title(f"Top {top_n} Materials with the most errors")
            ax.set_xlabel("Material"); ax.set_ylabel("Errors")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

    # ============================================================
    # NOVO: Materiais com View em falta por código de View
    # (Usa view_is_missing_df em vez de erros justificados)
    # ============================================================

    view_missing_df = st.session_state.get("sap_view_missing_df")

    if view_missing_df is not None and not view_missing_df.empty:
        by_view = (
            view_missing_df['Missing_View']
            .value_counts()
            .rename_axis('View')
            .reset_index(name='Materials Missing')
        )

        if PLOTLY_AVAILABLE:
            fig = px.bar(by_view, x='View', y='Materials Missing',
                         title="Materials Missing by View Code",
                         text='Materials Missing', color='View')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_view['View'], by_view['Materials Missing'], color="#d62728")
            ax.set_title("Materials Missing by View Code")
            ax.set_xlabel("View"); ax.set_ylabel("Materials Missing")
            st.pyplot(fig)

def render_sap_history_charts(logs: list[dict]) -> None:
    """
    Linha temporal do histórico SAP: total de erros ao longo do tempo,
    segregado por filtro de planta.
    """
    if not logs:
        st.info("No SAP history for charts..")
        return

    df = pd.DataFrame([{
        'timestamp': l.get('timestamp'),
        'user': l.get('user'),
        'plant_filter': l.get('plant_filter', 'ALL'),
        'total_errors': l.get('metrics', {}).get('total_errors', 0)
    } for l in logs])

    # Converter timestamp para datetime
    try:
        df['ts'] = pd.to_datetime(df['timestamp'])
    except Exception:
        df['ts'] = df['timestamp']

    st.markdown("### ⏱️ Temporal Evolution (SAP History)")
    dsorted = df.sort_values('ts')
    if PLOTLY_AVAILABLE:
        fig = px.line(dsorted, x='ts', y='total_errors', color='plant_filter',
                      markers=True, title="Total Errors over time (by plant filter)")
        fig = style_plotly(fig, height=480)
        st.plotly_chart(fig, width='stretch')
    else:
        fig, ax = plt.subplots(figsize=(8, 4))
        for plant, grp in dsorted.groupby('plant_filter'):
            ax.plot(grp['ts'], grp['total_errors'], marker='o', label=plant)
        ax.set_title("Total Errors over time (by plant filter)")
        ax.set_xlabel("Time"); ax.set_ylabel("Errors")
        ax.legend()
        st.pyplot(fig)

# =========================
#   PLM WORKFLOW (ATL)
# =========================

def render_atl_run_kpis(df_pt01: pd.DataFrame, df_pt02: pd.DataFrame) -> None:
    """
    KPIs para ATL:
      - Contagem de linhas PT01/PT02
      - Aging médio (se coluna existir)
    """
    pt01_rows = int(df_pt01.shape[0]) if df_pt01 is not None else 0
    pt02_rows = int(df_pt02.shape[0]) if df_pt02 is not None else 0

    def _avg_aging(df: pd.DataFrame) -> float:
        if df is None or df.empty or 'Aging' not in df.columns:
            return 0.0
        try:
            return round(pd.to_numeric(df['Aging'], errors='coerce').dropna().mean(), 1)
        except Exception:
            return 0.0

    avg_pt01 = _avg_aging(df_pt01)
    avg_pt02 = _avg_aging(df_pt02)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("PT01 — Lines", f"{pt01_rows:,}")
    c2.metric("PT02 — Lines", f"{pt02_rows:,}")
    c3.metric("PT01 — Aging Average (days)", avg_pt01)
    c4.metric("PT02 — Aging Average (days)", avg_pt02)

def render_atl_run_charts(df_pt01: pd.DataFrame, df_pt02: pd.DataFrame) -> None:
    """
    Gráficos ATL:
      - Histograma de Aging por planta
      - Contagem por Approval Status
      - Contagem por Categorias (X) — ex.: Pricing, Purchasing, MPL, etc.
    """
    st.markdown("### 📊 Charts — Current Execution (ATL)")

    # Unir dataframes e marcar a planta
    dfs = []
    if df_pt01 is not None and not df_pt01.empty:
        d = df_pt01.copy(); d['Plant'] = 'PT01'; dfs.append(d)
    if df_pt02 is not None and not df_pt02.empty:
        d = df_pt02.copy(); d['Plant'] = 'PT02'; dfs.append(d)

    if not dfs:
        st.info("No ATL data available for graphs.")
        return

    df_all = pd.concat(dfs, ignore_index=True)

    # Histograma de Aging
    if 'Aging' in df_all.columns:
        if PLOTLY_AVAILABLE:
            fig = px.histogram(df_all, x='Aging', color='Plant', nbins=30, title="Aging Distribution (days)")
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            for plant, grp in df_all.groupby('Plant'):
                ax.hist(pd.to_numeric(grp['Aging'], errors='coerce').dropna(), bins=30, alpha=0.6, label=plant)
            ax.set_title("Aging Distribution (days)")
            ax.set_xlabel("Aging (days)"); ax.set_ylabel("Count")
            ax.legend()
            st.pyplot(fig)

    # Contagem por Approval Status
    if 'Approval Status' in df_all.columns:
        by_status = df_all.groupby(['Plant', 'Approval Status']).size().reset_index(name='count')
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_status, x='Approval Status', y='count', color='Plant', barmode='group',
                         title="Approval Status Count")
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            for plant, grp in by_status.groupby('Plant'):
                ax.bar(grp['Approval Status'], grp['count'], alpha=0.7, label=plant)
            ax.set_title("Approval Status Count")
            ax.set_xlabel("Approval Status"); ax.set_ylabel("Count")
            ax.legend()
            st.pyplot(fig)

    # Categorias marcadas com "X"
    cat_cols = [c for c in ["Pricing", "Purchasing", "MPL", "Customs", "Sales",
                            "Manufacturing", "Costing", "MDS", "MPL - TL", "MDS - TL"]
                if c in df_all.columns]
    if cat_cols:
        melted = df_all.melt(id_vars=['Plant'], value_vars=cat_cols,
                             var_name='Category', value_name='Checked')
        melted['Checked'] = melted['Checked'].astype(str).str.upper().eq('X')
        by_cat = melted.groupby(['Plant', 'Category', 'Checked']).size().reset_index(name='count')
        by_cat = by_cat[by_cat['Checked']]
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_cat, x='Category', y='count', color='Plant', barmode='group',
                         title="Activities Checked by Category (X)")
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            for plant, grp in by_cat.groupby('Plant'):
                ax.bar(grp['Category'], grp['count'], alpha=0.7, label=plant)
            ax.set_title("Activities Checked by Category (X)")
            ax.set_xlabel("Category"); ax.set_ylabel("Count")
            ax.legend()
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

def render_atl_history_charts(logs: list[dict]) -> None:
    """
    Evolução temporal do ATL:
      - Curvas PT01 / PT02 com contagem de linhas ao longo do tempo.
    """
    if not logs:
        st.info("No ATL history for charts.")
        return

    df = pd.DataFrame([{
        'timestamp': l.get('timestamp'),
        'user': l.get('user'),
        'PT01_rows': l.get('metrics', {}).get('PT01_rows', 0),
        'PT02_rows': l.get('metrics', {}).get('PT02_rows', 0),
    } for l in logs])

    try:
        df['ts'] = pd.to_datetime(df['timestamp'])
    except Exception:
        df['ts'] = df['timestamp']

    st.markdown("### ⏱️ Temporal Evolution (ATL History)")
    dsorted = df.sort_values('ts')
    if PLOTLY_AVAILABLE:
        fig = px.line(dsorted, x='ts', y=['PT01_rows', 'PT02_rows'],
                      markers=True, title="Lines by Plant over Time")
        fig = style_plotly(fig, height=480) 
        st.plotly_chart(fig, width='stretch')
    else:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.plot(dsorted['ts'], dsorted['PT01_rows'], marker='o', label='PT01')
        ax.plot(dsorted['ts'], dsorted['PT02_rows'], marker='o', label='PT02')
        ax.set_title("Lines by Plant over Time")
        ax.set_xlabel("Time"); ax.set_ylabel("Lines")
        ax.legend()
        st.pyplot(fig)

# =========================
#   SAP VALIDATION PROGRESS (DIFF)
# =========================

def render_progress_kpis(diff_dfs: dict[str, pd.DataFrame]) -> None:
    """
    KPIs do comparador (Diff):
      - Erros Novos, Resolvidos, Alterados, Novos Materiais, Views Perdidas
    """
    def _n(name: str) -> int:
        df = diff_dfs.get(name, pd.DataFrame())
        return int(df.shape[0])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("New Errors", _n("New Errors"))
    c2.metric("Resolved Errors", _n("Resolved Errors"))
    c3.metric("Changed Errors", _n("Changed Errors"))
    c4.metric("New Materials", _n("New Materials"))
    c5.metric("Views Lost", _n("Views Lost"))

def render_progress_charts(diff_dfs: dict[str, pd.DataFrame]) -> None:
    """
    Gráficos do comparador:
      - Barras New vs Resolved vs Changed
      - Views Perdidas por código
    """
    st.markdown("### 📊 Charts — Reports Comparison")

    # Barras New/Resolved/Changed
    df_bars = pd.DataFrame({
        "Type": ["New Errors", "Resolved Errors", "Changed Errors"],
        "Amount": [
            diff_dfs.get("New Errors", pd.DataFrame()).shape[0],
            diff_dfs.get("Resolved Errors", pd.DataFrame()).shape[0],
            diff_dfs.get("Changed Errors", pd.DataFrame()).shape[0],
        ]
    })

    if PLOTLY_AVAILABLE:
        fig = px.bar(df_bars, x="Type", y="Amount", text='Amount',
                     title="Errors — New Errors vs Resolved vs Changed")
        fig.update_traces(textposition='outside')
        fig = style_plotly(fig, height=480)
        st.plotly_chart(fig, width='stretch')
    else:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.bar(df_bars["Type"], df_bars["Amount"], color="#9467bd")
        ax.set_title("Errors — New Errors vs Resolved vs Changed")
        ax.set_xlabel("Type"); ax.set_ylabel("Amount")
        st.pyplot(fig)

    # Views Perdidas por código
    df_views_lost = diff_dfs.get("Views Lost", pd.DataFrame())
    if not df_views_lost.empty and 'View_Lost' in df_views_lost.columns:
        by_view = (
            df_views_lost['View_Lost']
            .value_counts()
            .rename_axis('View')
            .reset_index(name='Amount')
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_view, x='View', y='Amount', text='Amount',
                         title="Lost Views by Code")
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_view['View'], by_view['Amount'], color="#8c564b")
            ax.set_title("VLost Views by Code")
            ax.set_xlabel("View"); ax.set_ylabel("Amount")
            st.pyplot(fig)

# =========================
# MRP VALIDATION (NOVO)
# =========================
def render_mrp_run_kpis(errors_df: pd.DataFrame) -> None:
    """
    KPIs básicos da execução MRP:
    - Total de erros
    - Fields distintos com erro
    - Materiais distintos com erro
    """
    if errors_df is None or errors_df.empty:
        st.info("No data for KPIs (MRP).")
        return

    total_errors = int(errors_df.shape[0])
    total_fields = int(errors_df['Field'].nunique()) if 'Field' in errors_df.columns else 0
    total_materials = int(errors_df['Material'].nunique()) if 'Material' in errors_df.columns else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Errors (MRP)", f"{total_errors:,}")
    c2.metric("Fields with Error (distinct)", f"{total_fields:,}")
    c3.metric("Materials with Errors (distinct)", f"{total_materials:,}")

def render_mrp_run_charts(errors_df: pd.DataFrame) -> None:
    """
    Gráficos da execução MRP:
    - Top N Fields com mais erros
    - Erros por Planta
    - Top N Materiais com mais erros
    """
    if errors_df is None or errors_df.empty:
        return

    st.markdown("### 📊 Charts — Present Execution (MRP)")
    top_n = int(errors_df['_TOP_N'].iloc[0]) if '_TOP_N' in errors_df.columns else 20

    # Top-N por Field
    if 'Field' in errors_df.columns:
        top_fields = (
            errors_df['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Errors')
            .head(top_n)
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(top_fields, x='Field', y='Errors',
                         title=f"Top {top_n} Fields with more Errors (MRP)", text='Errors')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(top_fields['Field'], top_fields['Errors'], color="#1f77b4")
            ax.set_title(f"Top {top_n} Fields with more Errors (MRP)")
            ax.set_xlabel("Field"); ax.set_ylabel("Errors")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

    # Erros por Plant
    if 'Plant' in errors_df.columns:
        by_plant = (
            errors_df['Plant'].astype(str).str.strip().replace({'': 'No Plant'})
            .value_counts()
            .rename_axis('Plant')
            .reset_index(name='Errors')
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_plant, x='Plant', y='Errors',
                         title="Errors by Plant (MRP)", text='Errors', color='Plant')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_plant['Plant'], by_plant['Errors'], color="#ff7f0e")
            ax.set_title("Errors by Plant (MRP)")
            ax.set_xlabel("Plant"); ax.set_ylabel("Errors")
            st.pyplot(fig)

    # Top-N por Material
    if 'Material' in errors_df.columns:
        by_mat = (
            errors_df['Material'].astype(str).str.strip()
            .value_counts()
            .rename_axis('Material')
            .reset_index(name='Errors')
            .head(top_n)
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_mat, x='Material', y='Errors',
                         title=f"Top {top_n} Materials with more Errors (MRP)", text='Errors')
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=480)
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_mat['Material'], by_mat['Errors'], color="#2ca02c")
            ax.set_title(f"Top {top_n} Materials with more Errors (MRP)")
            ax.set_xlabel("Material"); ax.set_ylabel("Errors")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

def render_mrp_history_charts(logs: list[dict]) -> None:
    """
    Evolução temporal do MRP:
    - Linha do total de erros ao longo do tempo por filtro de planta.
    """
    if not logs:
        st.info("No MRP history for charts.")
        return

    df = pd.DataFrame([{
        'timestamp': l.get('timestamp'),
        'user': l.get('user'),
        'plant_filter': l.get('plant_filter', 'ALL'),
        'total_errors': l.get('metrics', {}).get('total_errors', 0)
    } for l in logs])

    try:
        df['ts'] = pd.to_datetime(df['timestamp'])
    except Exception:
        df['ts'] = df['timestamp']

    st.markdown("### ⏱️ OverTime evolution (MRP History)")
    dsorted = df.sort_values('ts')
    if PLOTLY_AVAILABLE:
        fig = px.line(dsorted, x='ts', y='total_errors', color='plant_filter',
                      markers=True, title="Total Errors over time (MRP)")
        fig = style_plotly(fig, height=480)
        st.plotly_chart(fig, width='stretch')
    else:
        fig, ax = plt.subplots(figsize=(8, 4))
        for plant, grp in dsorted.groupby('plant_filter'):
            ax.plot(grp['ts'], grp['total_errors'], marker='o', label=plant)
        ax.set_title("Total Errors over time (MRP)")
        ax.set_xlabel("Time"); ax.set_ylabel("Errors")
        ax.legend()
        st.pyplot(fig)

# =========================
# Evolution Dashboartd (NOVO)
# =========================
def render_evolution_dashboard():
    """
    Evolution Dashboard (SAP + PLM Workflow + Cleanups) — REV9
    Inclui:
      - Períodos semanais convertidos para intervalos (ex: 05–11 Jan 2026)
      - KPIs inteligentes (Conjunto B)
      - Sparkline
      - Plant filter com ALL = visão global
      - Correções ISO-Week com fromisocalendar
      - Tick spacing inteligente
      - Paleta Visteon
      - Trend Detection
      - Forecast
      - Anomaly Detection
      - Auto-Insights
      - Exportação Excel
    """

    # ------------------------------------------------------------------
    # Função auxiliar: ISO Week → intervalo real (05–11 Jan 2026)
    # ------------------------------------------------------------------
    def iso_week_to_range(iso_str):
        if "-W" not in iso_str:
            return iso_str
        try:
            year, week = iso_str.split("-W")
            year = int(year)
            week = int(week)

            start = datetime.fromisocalendar(year, week, 1)  # Monday
            end = datetime.fromisocalendar(year, week, 7)    # Sunday

            # Mesmos mês → 05–11 Jan 2026
            if start.month == end.month:
                return f"{start.strftime('%d')}–{end.strftime('%d %b %Y')}"
            else:
                # Meses diferentes → 28 Jan–03 Feb 2026
                return f"{start.strftime('%d %b')}–{end.strftime('%d %b %Y')}"
        except:
            return iso_str

    st.subheader("📈 Evolution Dashboard")
    st.caption("Historical evolution of SAP, PLM Workflow and Data Cleanups validations.")

    # ---------------------------------------------------------
    # Load logs
    # ---------------------------------------------------------
    sap_logs = list_run_logs("sap", limit=500) or []
    atl_logs = list_run_logs("atl", limit=500) or []
    dc_logs  = list_run_logs("dataclean", limit=500) or []

    # Determine available sources
    available_sources = []
    if sap_logs: available_sources.append("SAP")
    if atl_logs: available_sources.append("PLM Workflow")
    if dc_logs:  available_sources.append("Cleanups")

    st.sidebar.markdown("### 📌 Data Sources")
    selected_sources = st.sidebar.multiselect(
        "Choose sources:", options=available_sources, default=available_sources
    )

    if not selected_sources:
        st.info("Please select at least one source.")
        return

    # ---------------------------------------------------------
    # Build unified dataset
    # ---------------------------------------------------------
    entries = []

    # SAP
    if "SAP" in selected_sources:
        for log in sap_logs:
            entries.append({
                "ts": pd.to_datetime(log["timestamp"]),
                "plant": log.get("plant_filter", "ALL"),
                "source": "SAP",
                "value": log.get("metrics", {}).get("total_errors", 0)
            })

    # PLM Workflow
    if "PLM Workflow" in selected_sources:
        for log in atl_logs:
            ts = pd.to_datetime(log["timestamp"])
            entries.append({
                "ts": ts, "plant": "PT01", "source": "PLM Workflow PT01",
                "value": log.get("metrics", {}).get("PT01_rows", 0)
            })
            entries.append({
                "ts": ts, "plant": "PT02", "source": "PLM Workflow PT02",
                "value": log.get("metrics", {}).get("PT02_rows", 0)
            })

    # Cleanups
    if "Cleanups" in selected_sources:
        for log in dc_logs:
            entries.append({
                "ts": pd.to_datetime(log["timestamp"]),
                "plant": log.get("plant_filter", "ALL"),
                "source": "Cleanups",
                "value": log.get("metrics", {}).get("outdated", 0)
            })

    if not entries:
        st.info("No data available for selected sources.")
        return

    df = pd.DataFrame(entries)

    # ---------------------------------------------------------
    # Date filter
    # ---------------------------------------------------------
    st.sidebar.markdown("### 📅 Date Range")
    dmin, dmax = df["ts"].min().date(), df["ts"].max().date()
    start_date, end_date = st.sidebar.date_input(
        "Period:", value=(dmin, dmax), min_value=dmin, max_value=dmax
    )
    df = df[(df["ts"].dt.date >= start_date) & (df["ts"].dt.date <= end_date)]

    # ---------------------------------------------------------
    # Group by
    # ---------------------------------------------------------
    mode = st.sidebar.selectbox("Group by:", ["Daily", "Weekly", "Monthly"], index=1)

    if mode == "Daily":
        df["period"] = df["ts"].dt.strftime("%Y-%m-%d")

    elif mode == "Weekly":
        iso = df["ts"].dt.isocalendar()
        df["period"] = iso["year"].astype(str) + "-W" + iso["week"].astype(str)

    else:
        df["period"] = df["ts"].dt.strftime("%Y-%m")

    # ---------------------------------------------------------
    # Plant filter (ALL = show all)
    # ---------------------------------------------------------
    plants = sorted(df["plant"].dropna().unique().tolist())
    real_plants = [p for p in plants if p != "ALL"]

    selected_plants = st.sidebar.multiselect(
        "Plants:", ["ALL"] + real_plants, default=["ALL"]
    )

    if "ALL" in selected_plants:
        df = df[df["plant"].isin(real_plants + ["ALL"])]
    else:
        df = df[df["plant"].isin(selected_plants)]

    # ---------------------------------------------------------
    # Group + robust sorting (ISO‑week fix)
    # ---------------------------------------------------------
    grouped = (
        df.groupby(["period", "source", "plant"], as_index=False)["value"]
          .mean()
    )

    def iso_to_date(iso_str):
        if "-W" not in iso_str:
            return pd.to_datetime(iso_str, errors="coerce")
        try:
            year, week = iso_str.split("-W")
            return datetime.fromisocalendar(int(year), int(week), 1)
        except:
            return pd.NaT

    grouped["sort_col"] = grouped["period"].apply(iso_to_date)
    grouped = grouped.sort_values("sort_col", kind="stable")

    # ---------------------------------------------------------
    # KPIs (Conjunto B) + sparkline
    # ---------------------------------------------------------
    sum_by_period = (
        grouped.groupby(["period", "sort_col"], as_index=False)["value"]
               .sum()
               .sort_values("sort_col", kind="stable")
    )

    # >>>>>>> NOVO: coluna friendly <<<<<<<<<
    sum_by_period["pretty_period"] = sum_by_period["period"].apply(iso_week_to_range)
    grouped["pretty_period"] = grouped["period"].apply(iso_week_to_range)

    # Current
    if len(sum_by_period):
        current_val = float(sum_by_period["value"].iloc[-1])
        current_label = str(sum_by_period["pretty_period"].iloc[-1])
    else:
        current_val, current_label = 0, "—"

    # Δ %
    if len(sum_by_period) >= 2:
        prev_val = float(sum_by_period["value"].iloc[-2])
        delta_pct = None if prev_val == 0 else (current_val - prev_val) / prev_val * 100
    else:
        delta_pct = None

    # Best / Worst
    if len(sum_by_period):
        best_idx = sum_by_period["value"].idxmin()
        worst_idx = sum_by_period["value"].idxmax()
        best_val = float(sum_by_period.loc[best_idx, "value"])
        best_label = str(sum_by_period.loc[best_idx, "pretty_period"])
        worst_val = float(sum_by_period.loc[worst_idx, "value"])
        worst_label = str(sum_by_period.loc[worst_idx, "pretty_period"])
    else:
        best_val = worst_val = 0
        best_label = worst_label = "—"

    # ---------------------------------------------------------
    # Render KPIs
    # ---------------------------------------------------------
    k1, k2, k3 = st.columns(3)

    with k1:
        if delta_pct is None:
            st.metric("Current Value", f"{current_val:,.0f}", delta="—", delta_color="inverse")
        else:
            st.metric("Current Value", f"{current_val:,.0f}",
                      delta=f"{delta_pct:.1f}%", delta_color="inverse")

        if PLOTLY_AVAILABLE and len(sum_by_period) > 1:
            spark = px.line(sum_by_period, x="pretty_period", y="value")
            spark.update_layout(template="simple_white", height=75,
                                margin=dict(l=0,r=0,t=0,b=0), showlegend=False)
            spark.update_xaxes(visible=False)
            spark.update_yaxes(visible=False)
            st.plotly_chart(spark, width='stretch')
        st.caption(f"Last period: {current_label}")

    with k2:
        st.metric("Best Period (min)", f"{best_val:,.0f}", delta=best_label)

    with k3:
        st.metric("Worst Period (max)", f"{worst_val:,.0f}", delta=worst_label)

    # ---------------------------------------------------------
    # Trend Detection
    # ---------------------------------------------------------
    st.markdown("### 📊 Trend Analysis")

    if len(sum_by_period) >= 3:
        y = sum_by_period["value"].to_numpy(dtype=float)
        x = np.arange(len(y))

        lr = linregress(x, y)
        slope = lr.slope
        r2 = lr.rvalue ** 2

        threshold = np.std(y) * 0.15

        if slope > threshold:
            trend_label = "📈 Strong Uptrend"
        elif slope > 0:
            trend_label = "↗ Weak Uptrend"
        elif slope < -threshold:
            trend_label = "📉 Strong Downtrend"
        elif slope < 0:
            trend_label = "↘ Weak Downtrend"
        else:
            trend_label = "➖ Sideways"

        st.info(
            f"**Trend:** {trend_label}\n\n"
            f"**Slope:** `{slope:.3f}`\n"
            f"**R² (confidence):** `{r2:.3f}`\n"
            f"**Data Points:** `{len(y)}`"
        )
    else:
        st.info("Not enough data points to compute trend.")

    # ---------------------------------------------------------
    # Main line chart (pretty_period)
    # ---------------------------------------------------------
    st.markdown(f"### 📉 Evolution Over Time ({mode})")

    if PLOTLY_AVAILABLE:
        fig = px.line(
            grouped, x="pretty_period", y="value",
            color="source", line_dash="plant",
            markers=True, title="Evolution of Validation Metrics"
        )

        fig.update_traces(marker=dict(size=8))

        COLOR_MAP = {
            "SAP": "#FF7200",
            "PLM Workflow PT01": "#6EC6FF",
            "PLM Workflow PT02": "#00A6A6",
            "Cleanups": "#4D4D4D"
        }

        fig.update_layout(
            colorway=[COLOR_MAP[s] for s in grouped["source"].unique() if s in COLOR_MAP]
        )

        unique_sorted = sum_by_period["pretty_period"].tolist()
        n = len(unique_sorted)
        step = 1 if n <= 12 else (2 if n <= 20 else (3 if n <= 35 else 4))

        fig.update_xaxes(tickmode="array", tickvals=unique_sorted[::step])
        fig.update_layout(xaxis_tickangle=-45, xaxis_tickfont=dict(size=11))

        fig = style_plotly(fig, height=480)
        st.plotly_chart(fig, width='stretch')
    else:
        # Fallback Matplotlib
        fig, ax = plt.subplots(figsize=(10, 5))
        for (src, pl), grp in grouped.groupby(["source", "plant"]):
            ax.plot(grp["pretty_period"], grp["value"],
                    marker="o", label=f"{src} ({pl})",
                    color=COLOR_MAP.get(src, "#000000"))
        ax.legend()
        plt.xticks(rotation=45)
        st.pyplot(fig)

    # ---------------------------------------------------------
    # Forecast (Next Period)
    # ---------------------------------------------------------
    st.markdown("### 🔮 Forecast (Next Period)")

    if len(sum_by_period) >= 3:

        y = sum_by_period["value"].to_numpy(dtype=float)
        x = np.arange(len(y))

        slope_f, intercept_f = np.polyfit(x, y, 1)
        next_x = len(y)
        forecast_next = slope_f * next_x + intercept_f

        conf = np.std(y) * 0.20
        lower = forecast_next - conf
        upper = forecast_next + conf

        if slope_f > 0:
            forecast_label = "📈 Expected increase"
        elif slope_f < 0:
            forecast_label = "📉 Expected decrease"
        else:
            forecast_label = "➖ Stable expectation"

        st.success(
            f"**Forecast:** {forecast_label}\n\n"
            f"**Next period estimate:** `{forecast_next:,.0f}`\n"
            f"**Range:** `{lower:,.0f}` → `{upper:,.0f}`"
        )

        if PLOTLY_AVAILABLE:
            fhist = sum_by_period.copy()
            fhist["type"] = "Historical"

            fnext = pd.DataFrame({
                "pretty_period": [f"Forecast +1 {mode}"],
                "value": [forecast_next],
                "type": ["Forecast"]
            })

            fig_f = px.line(fhist, x="pretty_period", y="value")
            fig_f.add_scatter(
                x=fnext["pretty_period"], y=fnext["value"],
                mode="markers+lines",
                name="Forecast", line=dict(dash="dash")
            )
            fig_f.add_scatter(
                x=fnext["pretty_period"], y=[upper],
                mode="markers", marker=dict(color="lightgray"),
                name="Upper Bound"
            )
            fig_f.add_scatter(
                x=fnext["pretty_period"], y=[lower],
                mode="markers", marker=dict(color="lightgray"),
                name="Lower Bound"
            )

            fig_f.update_layout(
                template="simple_white",
                height=300,
                margin=dict(l=10, r=10, t=20, b=10),
                showlegend=True
            )

            st.plotly_chart(fig_f, width='stretch')

    else:
        st.info("Not enough data points to generate forecast.")

    # ---------------------------------------------------------
    # Anomaly Detection
    # ---------------------------------------------------------
    st.markdown("### 🚨 Anomaly Detection")

    if len(sum_by_period) >= 5:

        vals = sum_by_period["value"].to_numpy(dtype=float)

        mean = np.mean(vals)
        std  = np.std(vals)
        median = np.median(vals)
        mad = np.median(np.abs(vals - median)) if np.median(np.abs(vals - median)) != 0 else 1e-6

        anomalies = []

        for i, row in sum_by_period.iterrows():
            v = row["value"]
            period = row["pretty_period"]   # <= now uses pretty
            z = (v - mean) / std if std > 0 else 0
            mod_z = 0.6745 * (v - median) / mad

            if abs(mod_z) > 3.5 or abs(z) > 3.5:
                severity = "🔥 Severe"
            elif abs(mod_z) > 2.5 or abs(z) > 2.5:
                severity = "🟠 Moderate"
            elif abs(mod_z) > 1.8 or abs(z) > 1.8:
                severity = "🟡 Mild"
            else:
                continue

            anomalies.append({
                "period": period,
                "value": v,
                "zscore": float(z),
                "mod_z": float(mod_z),
                "severity": severity
            })

        if anomalies:
            st.warning(f"**Detected {len(anomalies)} anomalies**:")
            for a in anomalies:
                st.markdown(
                    f"- {a['severity']} — **{a['period']}** → `{a['value']}` "
                    f"(Z={a['zscore']:.2f}, MAD={a['mod_z']:.2f})"
                )
        else:
            st.success("No anomalies detected in this period.")

    else:
        st.info("Need at least 5 periods for anomaly detection.")

    # ---------------------------------------------------------
    # Auto-Insights
    # ---------------------------------------------------------
    st.markdown("### 🧠 Auto‑Insights")

    context_lines = []

    if len(sum_by_period) >= 2:
        initial = float(sum_by_period["value"].iloc[0])
        final = float(sum_by_period["value"].iloc[-1])
        total_change = final - initial
        pct_change = (total_change / initial * 100) if initial != 0 else None

        if pct_change is not None:
            if pct_change > 0:
                context_lines.append(
                    f"Across the selected period, total errors increased by **{pct_change:.1f}%**, "
                    f"moving from **{initial:,.0f}** to **{final:,.0f}**."
                )
            elif pct_change < 0:
                context_lines.append(
                    f"Across the selected period, total errors decreased by **{abs(pct_change):.1f}%**, "
                    f"moving from **{initial:,.0f}** to **{final:,.0f}**."
                )
            else:
                context_lines.append(
                    "Across the selected period, total errors remained stable."
                )

    if len(sum_by_period) >= 3:
        context_lines.append(
            f"The regression model indicates a **{trend_label.replace('📈','').replace('📉','').strip()}**, "
            f"with slope **{slope:.3f}** and confidence (R²) **{r2:.3f}**."
        )

    if len(sum_by_period) >= 3:
        if slope > 0:
            context_lines.append(
                f"Forecast suggests an increase next period, near **{forecast_next:,.0f}**."
            )
        elif slope < 0:
            context_lines.append(
                f"Forecast points to a further reduction, around **{forecast_next:,.0f}**."
            )
        else:
            context_lines.append(
                f"Forecast indicates stable results around **{forecast_next:,.0f}**."
            )

    if "anomalies" in locals() and anomalies:
        severe = [a for a in anomalies if "Severe" in a["severity"]]
        if severe:
            context_lines.append(
                f"Detected **{len(severe)} severe anomalies**, indicating significant irregularities."
            )
        else:
            context_lines.append(
                "Detected only mild/moderate anomalies, suggesting normal fluctuations."
            )
    else:
        context_lines.append("No anomalies detected in this timeframe.")

    if len(selected_plants) > 1 or ("ALL" in selected_plants):
        plant_stats = grouped.groupby("plant")["value"].mean().sort_values()
        best_plant = plant_stats.index[0]
        worst_plant = plant_stats.index[-1]
        context_lines.append(
            f"Best performing plant: **{best_plant}**. Worst: **{worst_plant}**."
        )

    final_text = " ".join(context_lines)
    st.write(final_text)

    # ---------------------------------------------------------
    # EXPORTAÇÃO
    # ---------------------------------------------------------
    st.markdown("### 📤 Export Data")

    import io
    export_buffer = io.BytesIO()

    with pd.ExcelWriter(export_buffer, engine="openpyxl") as writer:

        sum_by_period.to_excel(writer, sheet_name="Historical_Summary", index=False)

        if "anomalies" in locals() and anomalies:
            pd.DataFrame(anomalies).to_excel(
                writer, sheet_name="Anomalies", index=False
            )
        else:
            pd.DataFrame(columns=["period","value","zscore","mod_z","severity"]).to_excel(
                writer, sheet_name="Anomalies", index=False
            )

        kpi_df = pd.DataFrame({
            "Metric": ["Current Value", "Delta %", "Best Value", "Best Period", "Worst Value", "Worst Period"],
            "Value": [
                current_val,
                f"{delta_pct:.1f}%" if delta_pct is not None else "—",
                best_val,
                best_label,
                worst_val,
                worst_label
            ]
        })
        kpi_df.to_excel(writer, sheet_name="KPIs", index=False)

        if len(sum_by_period) >= 3:
            trend_df = pd.DataFrame({
                "Slope": [slope],
                "R2": [r2],
                "Trend Classification": [trend_label]
            })
        else:
            trend_df = pd.DataFrame({
                "Slope": ["—"],
                "R2": ["—"],
                "Trend Classification": ["Not enough data"]
            })
        trend_df.to_excel(writer, sheet_name="Trend", index=False)

        if len(sum_by_period) >= 3:
            forecast_df = pd.DataFrame({
                "Forecast Next Period": [forecast_next],
                "Lower Bound": [lower],
                "Upper Bound": [upper],
                "Forecast Classification": [forecast_label]
            })
        else:
            forecast_df = pd.DataFrame({
                "Forecast Next Period": ["—"],
                "Lower Bound": ["—"],
                "Upper Bound": ["—"],
                "Forecast Classification": ["Not enough data"]
            })
        forecast_df.to_excel(writer, sheet_name="Forecast", index=False)

    st.download_button(
        label="⬇️ Download Evolution Report (Excel)",
        data=export_buffer.getvalue(),
        file_name="Evolution_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )