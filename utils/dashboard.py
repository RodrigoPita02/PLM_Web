
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
# Evolution Dashboard (REV9 — fixed)
# =========================
def render_evolution_dashboard():
    """
    Evolution Dashboard (SAP + PLM Workflow + Cleanups) — REV10
    - Agregação: Último registo / Soma / Média
    - Escala: Linear / Log / Eixo Duplo (SAP vs restantes)
    - Normalização opcional (0–100% por fonte)
    - Suavização (rolling window)
    - Exportação Excel
    """
    import io
    import numpy as np
    import pandas as pd
    import matplotlib.pyplot as plt
    from datetime import datetime

    import streamlit as st

    # Plotly (opcional)
    try:
        from plotly.subplots import make_subplots
        import plotly.express as px
        import plotly.graph_objects as go
        PLOTLY_OK = True
    except Exception:
        PLOTLY_OK = False

    # Paleta consistente
    COLOR_MAP = {
        "SAP": "#FF7200",
        "PLM Workflow PT01": "#6EC6FF",
        "PLM Workflow PT02": "#00A6A6",
        "Cleanups": "#4D4D4D"
    }

    # ---------- Aux: formata período ISO week em intervalo legível ----------
    def iso_week_to_range(iso_str: str) -> str:
        if "-W" not in iso_str:
            return iso_str
        try:
            year, week = iso_str.split("-W")
            year = int(year); week = int(week)
            start = datetime.fromisocalendar(year, week, 1)  # segunda
            end   = datetime.fromisocalendar(year, week, 7)  # domingo
            if start.month == end.month:
                return f"{start.strftime('%d')}–{end.strftime('%d %b %Y')}"
            else:
                return f"{start.strftime('%d %b')}–{end.strftime('%d %b %Y')}"
        except Exception:
            return iso_str

    # ---------- Aux: converte ISO week para data (para sort robusto) ----------
    def iso_to_date(iso_str: str):
        if "-W" not in iso_str:
            return pd.to_datetime(iso_str, errors="coerce")
        try:
            y, w = iso_str.split("-W")
            return datetime.fromisocalendar(int(y), int(w), 1)
        except Exception:
            return pd.NaT

    st.subheader("📈 Evolution Dashboard")
    st.caption("Historical evolution of SAP, PLM Workflow and Data Cleanups validations.")

    # ---------------------------------------------------------
    # 1) Carregar logs
    # ---------------------------------------------------------
    # (usa a infraestrutura já existente no projeto)
    from sap_validation_engine import list_run_logs  # já importado no módulo original
    sap_logs = list_run_logs("sap", limit=500) or []
    atl_logs = list_run_logs("atl", limit=500) or []
    dc_logs  = list_run_logs("dataclean", limit=500) or []

    # Fontes disponíveis
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
    # 2) Construir dataset unificado (ts, plant, source, value)
    # ---------------------------------------------------------
    entries = []

    # SAP → usa total_errors da execução
    if "SAP" in selected_sources:
        for log in sap_logs:
            entries.append({
                "ts": pd.to_datetime(log.get("timestamp")),
                "plant": log.get("plant_filter", "ALL"),
                "source": "SAP",
                "value": log.get("metrics", {}).get("total_errors", 0)
            })

    # PLM Workflow → linhas PT01/PT02
    if "PLM Workflow" in selected_sources:
        for log in atl_logs:
            ts = pd.to_datetime(log.get("timestamp"))
            entries.append({
                "ts": ts, "plant": "PT01", "source": "PLM Workflow PT01",
                "value": log.get("metrics", {}).get("PT01_rows", 0)
            })
            entries.append({
                "ts": ts, "plant": "PT02", "source": "PLM Workflow PT02",
                "value": log.get("metrics", {}).get("PT02_rows", 0)
            })

    # Cleanups → métrica "outdated" (confirmado nos logs)
    if "Cleanups" in selected_sources:
        for log in dc_logs:
            entries.append({
                "ts": pd.to_datetime(log.get("timestamp")),
                "plant": log.get("plant_filter", "ALL"),
                "source": "Cleanups",
                "value": log.get("metrics", {}).get("outdated", 0)
            })

    if not entries:
        st.info("No data available for selected sources.")
        return

    df = pd.DataFrame(entries)

    # ---------------------------------------------------------
    # 3) Filtros (Data / Agrupamento / Planta / Normalização / Smoothing)
    # ---------------------------------------------------------
    st.sidebar.markdown("### 📅 Date Range")
    dmin, dmax = df["ts"].min().date(), df["ts"].max().date()
    start_end = st.sidebar.date_input("Period:", value=(dmin, dmax), min_value=dmin, max_value=dmax)
    if isinstance(start_end, (list, tuple)) and len(start_end) == 2:
        start_date, end_date = start_end
    else:
        start_date, end_date = dmin, dmax
    df = df[(df["ts"].dt.date >= start_date) & (df["ts"].dt.date <= end_date)]

    st.sidebar.markdown("### 🧮 Aggregation")
    agg_mode = st.sidebar.selectbox(
        "Aggregate values per period by:",
        ["Latest per period (last log)", "Sum per period", "Average per period"],
        index=0
    )

    st.sidebar.markdown("### ⌛ Group by")
    mode = st.sidebar.selectbox("Periodicity:", ["Daily", "Weekly", "Monthly"], index=1)

    # Construir coluna "period"
    if mode == "Daily":
        df["period"] = df["ts"].dt.strftime("%Y-%m-%d")
    elif mode == "Weekly":
        iso = df["ts"].dt.isocalendar()
        df["period"] = iso["year"].astype(str) + "-W" + iso["week"].astype(str)
    else:
        df["period"] = df["ts"].dt.strftime("%Y-%m")

    # Plant filter
    plants = sorted(df["plant"].dropna().unique().tolist())
    real_plants = [p for p in plants if p != "ALL"]
    selected_plants = st.sidebar.multiselect("Plants:", ["ALL"] + real_plants, default=["ALL"])
    if "ALL" in selected_plants:
        df = df[df["plant"].isin(real_plants + ["ALL"])]
    else:
        df = df[df["plant"].isin(selected_plants)]

    # Normalização por fonte (% do máximo no intervalo filtrado)
    normalize = st.sidebar.checkbox("Normalize each source to % of its own max (0–100)", value=False)

    # Suavização (rolling)
    smoothing_on = st.sidebar.checkbox("Apply smoothing (rolling mean)", value=False)
    window = st.sidebar.slider("Window (periods)", min_value=2, max_value=7, value=3, disabled=not smoothing_on)

    # Escala
    st.sidebar.markdown("### 📏 Scale")
    scale_choice = st.sidebar.selectbox(
        "Y axis scale:",
        ["Linear", "Log", "Dual axis (SAP vs others)"],
        index=2
    )

    # ---------------------------------------------------------
    # 4) Preparar agregação por período/fonte/planta
    # ---------------------------------------------------------
    # Ordenar por timestamp (necessário para "Latest per period")
    df = df.sort_values(["period", "source", "plant", "ts"])

    value_col = "value"
    if normalize:
        # Normaliza por fonte (todas as plantas dessa fonte, no intervalo activo)
        df["_max_by_source"] = df.groupby("source")[value_col].transform(lambda s: s.max() if s.max() > 0 else 1.0)
        df["value_norm"] = (df[value_col] / df["_max_by_source"]) * 100.0
        value_col = "value_norm"

    # Agregação base por (period, source, plant)
    if agg_mode.startswith("Latest"):
        # Último log do período para cada (source, plant)
        grp = df.groupby(["period", "source", "plant"], as_index=False).tail(1)
        grouped = grp.groupby(["period", "source", "plant"], as_index=False)[value_col].sum()
    elif agg_mode.startswith("Sum"):
        grouped = df.groupby(["period", "source", "plant"], as_index=False)[value_col].sum()
    else:  # Average
        grouped = df.groupby(["period", "source", "plant"], as_index=False)[value_col].mean()

    # Colunas auxiliares para ordenação e apresentação
    grouped["sort_col"] = grouped["period"].apply(iso_to_date)
    grouped = grouped.sort_values("sort_col", kind="stable")
    grouped["pretty_period"] = grouped["period"].apply(iso_week_to_range)

    # Suavização por série (source+plant)
    plot_df = grouped.copy()
    if smoothing_on and len(plot_df["period"].unique()) >= window:
        plot_df["value_smooth"] = (
            plot_df.sort_values(["source", "plant", "sort_col"])
                  .groupby(["source", "plant"])[value_col]
                  .transform(lambda s: s.rolling(window=window, min_periods=1).mean())
        )
        y_plot_col = "value_smooth"
    else:
        y_plot_col = value_col

    # ---------------------------------------------------------
    # 5) KPIs (Conjunto B) a partir do total por período (soma das séries)
    # ---------------------------------------------------------
    sum_by_period = (
        plot_df.groupby(["period", "pretty_period", "sort_col"], as_index=False)[y_plot_col]
               .sum()
               .sort_values("sort_col", kind="stable")
               .rename(columns={y_plot_col: "value"})
    )

    if len(sum_by_period):
        current_val = float(sum_by_period["value"].iloc[-1])
        current_label = str(sum_by_period["pretty_period"].iloc[-1])
    else:
        current_val, current_label = 0.0, "—"

    if len(sum_by_period) >= 2:
        prev_val = float(sum_by_period["value"].iloc[-2])
        delta_pct = None if prev_val == 0 else (current_val - prev_val) / prev_val * 100
    else:
        delta_pct = None

    if len(sum_by_period):
        best_idx = sum_by_period["value"].idxmin()
        worst_idx = sum_by_period["value"].idxmax()
        best_val = float(sum_by_period.loc[best_idx, "value"])
        best_label = str(sum_by_period.loc[best_idx, "pretty_period"])
        worst_val = float(sum_by_period.loc[worst_idx, "value"])
        worst_label = str(sum_by_period.loc[worst_idx, "pretty_period"])
    else:
        best_val = worst_val = 0.0
        best_label = worst_label = "—"

    # ---------------------------------------------------------
    # 6) Render KPIs + sparkline
    # ---------------------------------------------------------
    k1, k2, k3 = st.columns(3)
    with k1:
        if delta_pct is None:
            st.metric("Current Value", f"{current_val:,.0f}", delta="—", delta_color="inverse")
        else:
            st.metric("Current Value", f"{current_val:,.0f}", delta=f"{delta_pct:.1f}%", delta_color="inverse")

        if PLOTLY_OK and len(sum_by_period) > 1:
            spark = px.line(sum_by_period, x="pretty_period", y="value")
            spark.update_layout(template="simple_white", height=75,
                                margin=dict(l=0, r=0, t=0, b=0), showlegend=False)
            spark.update_xaxes(visible=False); spark.update_yaxes(visible=False)
            st.plotly_chart(spark, width='stretch')
        st.caption(f"Last period: {current_label}")

    with k2:
        st.metric("Best Period (min)", f"{best_val:,.0f}", delta=best_label)
    with k3:
        st.metric("Worst Period (max)", f"{worst_val:,.0f}", delta=worst_label)

    # ---------------------------------------------------------
    # 7) GRÁFICO PRINCIPAL
    # ---------------------------------------------------------
    st.markdown(f"### 📉 Evolution Over Time ({mode})")

    # Escolha entre Linear / Log / Eixo Duplo
    if PLOTLY_OK:
        # Eixo-duplo com SAP à esquerda e restantes à direita
        use_dual = (scale_choice.startswith("Dual"))
        if use_dual:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            # Título e estilo base
            fig.update_layout(title="Evolution of Validation Metrics")
        else:
            # Um único eixo
            fig = go.Figure()
            fig.update_layout(title="Evolution of Validation Metrics")

        # Adiciona uma trace por série (source, plant)
        for (src, pl), grp in plot_df.groupby(["source", "plant"], sort=False):
            grp_sorted = grp.sort_values("sort_col")
            x_vals = grp_sorted["pretty_period"].tolist()
            y_vals = grp_sorted[y_plot_col].tolist()
            name = f"{src}, {pl}"
            color = COLOR_MAP.get(src, None)

            trace = go.Scatter(
                x=x_vals, y=y_vals, mode="lines+markers",
                name=name, marker=dict(size=8),
                line=dict(shape="linear"),
                marker_color=color
            )

            if use_dual:
                # SAP vai para o eixo esquerdo; outras para o direito
                is_sap = src.startswith("SAP")
                fig.add_trace(trace, secondary_y=not (not is_sap))  # True->right, False->left
            else:
                fig.add_trace(trace)

        # Escalas
        if scale_choice == "Log":
            fig.update_yaxes(type="log")

        if use_dual:
            fig.update_yaxes(title_text="SAP scale", secondary_y=False)
            fig.update_yaxes(title_text="PLM / Cleanups scale", secondary_y=True)

        # Eixo X (ticks espaçados conforme nº de períodos)
        unique_sorted = sum_by_period["pretty_period"].tolist()
        n = len(unique_sorted)
        step = 1 if n <= 12 else (2 if n <= 20 else (3 if n <= 35 else 4))
        fig.update_xaxes(tickmode="array", tickvals=unique_sorted[::step], tickangle=-45)

        # Cores (mantém palette por fonte)
        # (se houver fontes não mapeadas, Plotly atribui automaticamente)
        fig.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5)
        )

        # Altura e margens agradáveis
        fig.update_layout(height=480, margin=dict(l=40, r=40, t=60, b=80))
        st.plotly_chart(fig, width='stretch')

    else:
        # Fallback Matplotlib (inclui dual-axis com twinx)
        fig, ax_left = plt.subplots(figsize=(10, 5))
        if scale_choice.startswith("Dual"):
            ax_right = ax_left.twinx()
        else:
            ax_right = None

        for (src, pl), grp in plot_df.groupby(["source", "plant"], sort=False):
            grp_sorted = grp.sort_values("sort_col")
            x_vals = grp_sorted["pretty_period"].tolist()
            y_vals = grp_sorted[y_plot_col].tolist()
            label = f"{src}, {pl}"
            color = COLOR_MAP.get(src, None)

            if scale_choice.startswith("Dual") and not src.startswith("SAP") and ax_right is not None:
                ax_right.plot(x_vals, y_vals, marker="o", label=label, color=color)
            else:
                ax_left.plot(x_vals, y_vals, marker="o", label=label, color=color)

        if scale_choice == "Log":
            ax_left.set_yscale("log")
        ax_left.set_title("Evolution of Validation Metrics")
        ax_left.set_xlabel("Period"); ax_left.set_ylabel("Value")
        if ax_right is not None:
            ax_right.set_ylabel("PLM / Cleanups")
            # cria legendas agregadas
            lines, labels = ax_left.get_legend_handles_labels()
            lines2, labels2 = ax_right.get_legend_handles_labels()
            ax_left.legend(lines + lines2, labels + labels2, loc="upper left")
        else:
            ax_left.legend(loc="upper left")
        plt.xticks(rotation=45)
        st.pyplot(fig)

    # ---------------------------------------------------------
    # 8) Trend + Anomalias (idêntico ao que já tinhas, mas usando sum_by_period)
    # ---------------------------------------------------------
    st.markdown("### 📊 Trend Analysis")
    if len(sum_by_period) >= 3:
        y = sum_by_period["value"].to_numpy(dtype=float)
        x = np.arange(len(y))
        from scipy.stats import linregress
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

    # Anomalias
    st.markdown("### 🚨 Anomaly Detection")
    if len(sum_by_period) < 5:
        st.info("Need at least 5 periods to run anomaly detection.")
        anomalies_df = pd.DataFrame()
    else:
        # Controlos
        with st.expander("How anomalies are detected", expanded=False):
            st.write(
                "- **Z-score**: distance from the mean in standard deviations.\n"
                "- **Modified Z-score (MAD)**: robust alternative using median and MAD.\n"
                "- We flag periods whose absolute score exceeds the chosen thresholds."
            )
        c1, c2, c3, c4 = st.columns([1,1,1,1])
        with c1:
            metric = st.selectbox("Scoring", ["Modified Z (MAD)", "Z-score"], index=0)
        with c2:
            thr_mild = st.slider("Mild ≥", min_value=1.0, max_value=3.0, value=1.8, step=0.1)
        with c3:
            thr_mod  = st.slider("Moderate ≥", min_value=2.0, max_value=4.0, value=2.5, step=0.1)
        with c4:
            thr_sev  = st.slider("Severe ≥", min_value=3.0, max_value=5.0, value=3.5, step=0.1)

        vals = sum_by_period["value"].to_numpy(dtype=float)
        periods = sum_by_period["pretty_period"].astype(str).tolist()

        mean = float(np.mean(vals)); std = float(np.std(vals)) or 1e-6
        median = float(np.median(vals))
        mad    = float(np.median(np.abs(vals - median))) or 1e-6

        z_scores     = (vals - mean) / std
        mod_z_scores = 0.6745 * (vals - median) / mad
        score = mod_z_scores if metric == "Modified Z (MAD)" else z_scores

        severities = []
        for s in np.abs(score):
            if s >= thr_sev:   severities.append("🔥 Severe")
            elif s >= thr_mod: severities.append("🟠 Moderate")
            elif s >= thr_mild:severities.append("🟡 Mild")
            else:              severities.append("")

        anomalies_df = pd.DataFrame({
            "period": periods,
            "value": vals,
            "zscore": z_scores,
            "mod_z": mod_z_scores,
            "severity": severities
        })
        anomalies_df = anomalies_df[anomalies_df["severity"] != ""].copy()
        key_col = "mod_z" if metric == "Modified Z (MAD)" else "zscore"
        anomalies_df["abs_score"] = anomalies_df[key_col].abs()
        anomalies_df.sort_values("abs_score", ascending=False, inplace=True)

        # Visual
        if PLOTLY_OK:
            fig = px.line(sum_by_period, x="pretty_period", y="value", title="Time Series with Anomalies")
            if not anomalies_df.empty:
                color_map = {"🟡 Mild": "#F4D03F", "🟠 Moderate": "#E67E22", "🔥 Severe": "#E74C3C"}
                anomalies_df["color"] = anomalies_df["severity"].map(color_map)
                fig.add_scatter(
                    x=anomalies_df["period"], y=anomalies_df["value"],
                    mode="markers", marker=dict(size=11, color=anomalies_df["color"]),
                    name="Anomalies",
                    text=[
                        f"{sev}<br>{p}<br>Value: {v:,.0f}"
                        f"<br>Z={z:.2f} | MAD={m:.2f}"
                        for sev, p, v, z, m in zip(
                            anomalies_df["severity"], anomalies_df["period"],
                            anomalies_df["value"], anomalies_df["zscore"], anomalies_df["mod_z"]
                        )
                    ],
                    hovertemplate="%{text}<extra></extra>"
                )
            fig.update_layout(template="simple_white", height=400, margin=dict(l=40, r=40, t=60, b=40))
            fig.update_xaxes(tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            st.line_chart(sum_by_period.set_index("pretty_period")["value"])
            if not anomalies_df.empty:
                st.write("Anomaly markers not available without Plotly.")

        # Painel resumo
        if anomalies_df.empty:
            st.success("No anomalies detected for the current thresholds.")
        else:
            st.warning(f"**Detected {len(anomalies_df)} anomalies** (sorted by |{metric}|):")
            max_n = int(min(10, len(anomalies_df)))
            if max_n <= 1:
                a = anomalies_df.iloc[0]
                st.markdown(
                    f"- {a['severity']} — **{a['period']}** → `{int(a['value'])}` "
                    f"(Z={a['zscore']:.2f}, MAD={a['mod_z']:.2f})"
                )
            else:
                default_n = int(min(5, max_n))
                top_n = st.slider("Show top N anomalies", min_value=1, max_value=max_n, value=default_n)
                for _, a in anomalies_df.head(top_n).iterrows():
                    st.markdown(
                        f"- {a['severity']} — **{a['period']}** → `{int(a['value'])}` "
                        f"(Z={a['zscore']:.2f}, MAD={a['mod_z']:.2f})"
                    )

            # Download CSV
            st.download_button(
                label="⬇️ Download anomalies (CSV)",
                data=anomalies_df.drop(columns=["abs_score"]).to_csv(index=False).encode("utf-8"),
                file_name="anomalies.csv",
                mime="text/csv"
            )

    # ---------------------------------------------------------
    # 9) Auto-Insights (texto)
    # ---------------------------------------------------------
    st.markdown("### 🧠 Auto‑Insights")
    context_lines = []

    if len(sum_by_period) >= 2:
        initial = float(sum_by_period["value"].iloc[0])
        final   = float(sum_by_period["value"].iloc[-1])
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
                context_lines.append("Across the selected period, total errors remained stable.")

    if len(sum_by_period) >= 3:
        # Reutiliza variáveis calculadas acima (trend_label, slope, r2)
        trend_txt = locals().get("trend_label", "—")
        slope_val = locals().get("slope", 0.0)
        r2_val    = locals().get("r2", 0.0)
        trend_clean = (trend_txt.replace('📈', '').replace('📉', '')
                                  .replace('↗', '').replace('↘', '')
                                  .strip())
        context_lines.append(
            f"The regression model indicates a **{trend_clean}**, "
            f"with slope **{slope_val:.3f}** and confidence (R²) **{r2_val:.3f}**."
        )

    if "anomalies_df" in locals() and not anomalies_df.empty:
        severe = anomalies_df[anomalies_df["severity"].str.contains("Severe")]
        if len(severe):
            context_lines.append(f"Detected **{len(severe)} severe anomalies**, indicating significant irregularities.")
        else:
            context_lines.append("Detected only mild/moderate anomalies, suggesting normal fluctuations.")
    else:
        context_lines.append("No anomalies detected in this timeframe.")

    # Insights por planta quando faz sentido
    if len(selected_plants) > 1 or ("ALL" in selected_plants):
        plant_stats = plot_df.groupby("plant")[y_plot_col].mean().sort_values()
        if not plant_stats.empty:
            best_plant = plant_stats.index[0]
            worst_plant = plant_stats.index[-1]
            context_lines.append(f"Best performing plant: **{best_plant}**. Worst: **{worst_plant}**.")

    st.write(" ".join(context_lines))

    # ---------------------------------------------------------
    # 10) Exportação Excel
    # ---------------------------------------------------------
    st.markdown("### 📤 Export Data")
    export_buffer = io.BytesIO()
    with pd.ExcelWriter(export_buffer, engine="openpyxl") as writer:
        sum_by_period.to_excel(writer, sheet_name="Historical_Summary", index=False)

        # KPIs
        kpi_df = pd.DataFrame({
            "Metric": ["Aggregation", "Scale", "Current Value", "Delta %", "Best Value", "Best Period", "Worst Value", "Worst Period"],
            "Value": [
                agg_mode,
                scale_choice,
                current_val,
                f"{delta_pct:.1f}%" if delta_pct is not None else "—",
                best_val,
                best_label,
                worst_val,
                worst_label
            ]
        })
        kpi_df.to_excel(writer, sheet_name="KPIs", index=False)

        # Anomalias (se existirem)
        if "anomalies_df" in locals() and not anomalies_df.empty:
            anomalies_df.drop(columns=["abs_score"], errors="ignore").to_excel(
                writer, sheet_name="Anomalies", index=False
            )
        else:
            pd.DataFrame(columns=["period","value","zscore","mod_z","severity"]).to_excel(
                writer, sheet_name="Anomalies", index=False
            )

        # Trend
        if len(sum_by_period) >= 3:
            trend_df = pd.DataFrame({"Slope": [locals().get("slope", 0.0)],
                                     "R2": [locals().get("r2", 0.0)],
                                     "Trend Classification": [locals().get("trend_label", "—")]})
        else:
            trend_df = pd.DataFrame({"Slope": ["—"], "R2": ["—"], "Trend Classification": ["Not enough data"]})
        trend_df.to_excel(writer, sheet_name="Trend", index=False)

    st.download_button(
        label="⬇️ Download Evolution Report (Excel)",
        data=export_buffer.getvalue(),
        file_name="Evolution_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )