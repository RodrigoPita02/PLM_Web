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
    import matplotlib.pyplot as plt

# --- COLOR UTILITIES ---
def hex_to_rgba(hex_color: str, alpha: float) -> str:
    """Convert #RRGGBB to rgba(r,g,b,a)."""
    hc = hex_color.lstrip("#")
    r = int(hc[0:2], 16)
    g = int(hc[2:4], 16)
    b = int(hc[4:6], 16)
    return f"rgba({r}, {g}, {b}, {alpha})"

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
            # Destaque para campos MVKE (nova regra)
            HIGHLIGHT_FIELDS = {'MVKE Combinations', 'Account Assignment Group',
                                 'Item Category Group', 'Product Hierarchy'}
            colors = [
                "#FF6A13" if f in HIGHLIGHT_FIELDS else "#6EC6FF"
                for f in top_fields['Field']
            ]
            fig = px.bar(top_fields, x='Field', y='Errors',
                         title=f"Top {top_n} Fields with more errors", text='Errors')
            fig.update_traces(textposition='outside', marker_color=colors)
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
      - Aging médio por planta
      - Change Numbers únicos por planta
      - CNs com Timing Letter Number associado
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

    def _unique_cns(df: pd.DataFrame) -> int:
        if df is None or df.empty or 'Trigger ID' not in df.columns:
            return 0
        return int(df['Trigger ID'].astype(str).str.strip().nunique())

    def _cns_with_tln(df: pd.DataFrame) -> int:
        if df is None or df.empty or 'Timing Letter Number' not in df.columns:
            return 0
        return int(
            df['Timing Letter Number'].astype(str).str.strip()
            .replace('', pd.NA).notna().sum()
        )

    avg_pt01 = _avg_aging(df_pt01)
    avg_pt02 = _avg_aging(df_pt02)
    cns_pt01 = _unique_cns(df_pt01)
    cns_pt02 = _unique_cns(df_pt02)
    tln_pt01 = _cns_with_tln(df_pt01)
    tln_pt02 = _cns_with_tln(df_pt02)

    st.markdown("#### PT01")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("PT01 — Lines", f"{pt01_rows:,}")
    c2.metric("PT01 — Unique CNs", f"{cns_pt01:,}")
    c3.metric("PT01 — CNs with TLN", f"{tln_pt01:,}")
    c4.metric("PT01 — Aging Avg (days)", avg_pt01)

    st.markdown("#### PT02")
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("PT02 — Lines", f"{pt02_rows:,}")
    c6.metric("PT02 — Unique CNs", f"{cns_pt02:,}")
    c7.metric("PT02 — CNs with TLN", f"{tln_pt02:,}")
    c8.metric("PT02 — Aging Avg (days)", avg_pt02)

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

    # --- Timing Letter Number coverage por Source ---
    if 'Timing Letter Number' in df_all.columns and 'Source' in df_all.columns:
        tln_summary = []
        for src in df_all['Source'].unique():
            grp = df_all[df_all['Source'] == src]
            total = len(grp)
            with_tln = int(
                grp['Timing Letter Number'].astype(str).str.strip()
                .replace('', pd.NA).notna().sum()
            )
            without_tln = total - with_tln
            tln_summary.append({'Source': str(src), 'With TLN': with_tln, 'Without TLN': without_tln})
        tln_df = pd.DataFrame(tln_summary)

        if PLOTLY_AVAILABLE:
            tln_melted = tln_df.melt(id_vars='Source', value_vars=['With TLN', 'Without TLN'],
                                     var_name='TLN Status', value_name='Count')
            fig = px.bar(tln_melted, x='Source', y='Count', color='TLN Status', barmode='stack',
                         title="Timing Letter Number Coverage by Source",
                         color_discrete_map={'With TLN': '#00A6A6', 'Without TLN': '#D9D9D9'})
            fig.update_traces(texttemplate='%{y}', textposition='inside')
            fig = style_plotly(fig, height=380)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(6, 4))
            for row in tln_df.itertuples():
                ax.bar(row.Source, row._2, label='With TLN', color='#00A6A6')  # With TLN
                ax.bar(row.Source, row._3, bottom=row._2, label='Without TLN', color='#D9D9D9')
            ax.set_title("TLN Coverage by Source")
            ax.set_ylabel("Count")
            st.pyplot(fig)

    # --- Top 15 CNs mais antigos (maior Aging) ---
    if 'Aging' in df_all.columns and 'Trigger ID' in df_all.columns:
        aging_num = pd.to_numeric(df_all['Aging'], errors='coerce')
        top_aging = (
            df_all.assign(_aging=aging_num)
            .dropna(subset=['_aging'])
            .sort_values('_aging', ascending=False)
            .drop_duplicates(subset=['Trigger ID'])
            .head(15)[['Trigger ID', 'Source', 'Plant', 'Timing Letter Number', '_aging']]
            .rename(columns={'_aging': 'Aging (days)'})
            .reset_index(drop=True)
        )
        if not top_aging.empty:
            if PLOTLY_AVAILABLE:
                fig = px.bar(top_aging, x='Trigger ID', y='Aging (days)', color='Plant',
                             title="Top 15 Oldest Change Numbers (by Aging)",
                             text='Aging (days)',
                             hover_data=['Source', 'Timing Letter Number'])
                fig.update_traces(textposition='outside')
                fig = style_plotly(fig, height=420)
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, width='stretch')
            else:
                fig, ax = plt.subplots(figsize=(10, 4))
                ax.bar(top_aging['Trigger ID'], top_aging['Aging (days)'], color='#FF7200')
                ax.set_title("Top 15 Oldest Change Numbers")
                ax.set_xlabel("Change Number"); ax.set_ylabel("Aging (days)")
                plt.xticks(rotation=45, ha='right')
                st.pyplot(fig)

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
            ax.set_title("Lost Views by Code")
            ax.set_xlabel("View"); ax.set_ylabel("Amount")
            st.pyplot(fig)

    # New Errors breakdown by Field
    df_new_errors = diff_dfs.get("New Errors", pd.DataFrame())
    if not df_new_errors.empty and 'Field' in df_new_errors.columns:
        by_field = (
            df_new_errors['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='New Errors')
            .head(20)
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_field, x='Field', y='New Errors', text='New Errors',
                         title="New Errors — Breakdown by Field (Top 20)")
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=440)
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_field['Field'], by_field['New Errors'], color="#5e3a8c")
            ax.set_title("New Errors — Breakdown by Field")
            ax.set_xlabel("Field"); ax.set_ylabel("Count")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

    # Resolved Errors breakdown by Field
    df_resolved = diff_dfs.get("Resolved Errors", pd.DataFrame())
    if not df_resolved.empty and 'Field' in df_resolved.columns:
        by_field_r = (
            df_resolved['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Resolved Errors')
            .head(20)
        )
        if PLOTLY_AVAILABLE:
            fig = px.bar(by_field_r, x='Field', y='Resolved Errors', text='Resolved Errors',
                         title="Resolved Errors — Breakdown by Field (Top 20)",
                         color_discrete_sequence=["#00A6A6"])
            fig.update_traces(textposition='outside')
            fig = style_plotly(fig, height=440)
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, width='stretch')
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_field_r['Field'], by_field_r['Resolved Errors'], color="#00A6A6")
            ax.set_title("Resolved Errors — Breakdown by Field")
            ax.set_xlabel("Field"); ax.set_ylabel("Count")
            plt.xticks(rotation=45, ha='right')
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
# Evolution Dashboard (REV12 — Executive)
# =========================
def render_evolution_dashboard():
    """
    Evolution Dashboard — REV12 (Executive Edition)
    Upgrades vs REV11:
      A) Branded banner header with overall health badge (green/amber/red)
      B) Custom HTML/CSS KPI cards — richer visual than st.metric()
      C) Activity feed always visible (no expander), compact and styled
      D) Per-series mini sparklines above the main chart
      E) Main chart: professional theme, peak/valley auto-annotations,
         reference lines (mean per series), area fill under SAP line
      F) Trend panel: 3-column layout (icon+label | stats | sparkline)
      G) Anomaly chart: shaded band ±1σ around mean, improved markers
      H) Auto-Insights: bullet points replaced by styled alert boxes
      I) Section dividers with consistent visual language
    """
    import io
    import numpy as np
    import pandas as pd
    import matplotlib.pyplot as plt
    from datetime import datetime
    from scipy.stats import linregress

    import streamlit as st
    from sap_validation_engine import list_run_logs

    try:
        from plotly.subplots import make_subplots
        import plotly.express as px
        import plotly.graph_objects as go
        PLOTLY_OK = True
    except Exception:
        PLOTLY_OK = False

    # ── Palette & theme ───────────────────────────────────────────────────────
    # ── Palette & theme ───────────────────────────────────────────────────────
    CLR = {
        # Brand accents (Visteon)
        "sap":        "#FF7200",   # Visteon orange
        "plm01":      "#6EC6FF",   # light blue
        "plm02":      "#00A6A6",   # teal
        "tln01":      "#2176AE",   # blue
        "tln02":      "#1A936F",   # green
        "cleanups":   "#4D4D4D",   # graphite

        # Status
        "green":      "#1A936F",
        "amber":      "#F4A011",
        "red":        "#C00000",

        # DARK UI tokens (new)
        "bg_page":    "#0F141A",   # app background
        "bg_panel":   "#141B22",   # section panels
        "bg_card":    "#161F28",   # cards
        "bg_alt":     "#1A2330",   # alt rows / hover
        "border":     "#253041",
        "grid":       "#2A384D",   # grid lines on plots
        "text_main":  "#E6EDF3",   # main text
        "text_muted": "#9FB0C0",   # secondary text
        "text_invert":"#0E1116",   # for pills/badges on light fills
        "badge_bg":   "#0E1A26",   # subtle badge background
    }

    COLOR_MAP = {
        "SAP":               CLR["sap"],
        "PLM Workflow PT01": CLR["plm01"],
        "PLM Workflow PT02": CLR["plm02"],
        "PLM TLN PT01":      CLR["tln01"],
        "PLM TLN PT02":      CLR["tln02"],
        "Cleanups":          CLR["cleanups"],
    }


    # ── CSS injection ─────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    /* ---------- Page & Typography ---------- */
    :root{
    --bg-page:   #0F141A;
    --bg-panel:  #141B22;
    --bg-card:   #161F28;
    --bg-alt:    #1A2330;
    --border:    #253041;
    --grid:      #2A384D;

    --txt-main:  #E6EDF3;
    --txt-muted: #9FB0C0;
    --txt-inv:   #0E1116;

    --sap:   #FF7200;
    --plm01: #6EC6FF;
    --plm02: #00A6A6;
    --tln01: #2176AE;
    --tln02: #1A936F;
    --clean: #4D4D4D;

    --good:  #1A936F;
    --warn:  #F4A011;
    --bad:   #C00000;

    --badge-bg: #0E1A26;
    }

    html, body, .block-container{
    background: var(--bg-page) !important;
    color: var(--txt-main) !important;
    }
    h1,h2,h3,h4,h5,h6{ color: var(--txt-main); }
    hr{ border-color: var(--border); opacity: 0.6; }

    /* ---------- Exec banner ---------- */
    .exec-banner {
    background: linear-gradient(135deg, #10304C 0%, #1D4F7A 60%, #2A6FA5 100%);
    border-radius: 14px;
    padding: 22px 28px 18px 28px;
    margin: 10px 0 20px 0;
    color: #fff;
    border: 1px solid var(--border);
    box-shadow: 0 10px 30px rgba(0,0,0,.28);
    }
    .exec-banner h1{
    font-size: 1.55rem; font-weight: 800; margin: 0 0 6px 0;
    }
    .exec-banner .subtitle{
    font-size: .9rem; color: #E6EDF3; opacity: .9; margin: 0;
    }
    .exec-banner .badge{
    display: inline-block; padding: 4px 12px; border-radius: 999px;
    font-size: .80rem; font-weight: 700; margin-left: 10px;
    background: var(--badge-bg); border: 1px solid var(--border);
    color: #E6EDF3;
    }
    .badge-green{ background:#0F251E; color:#30C08D; border-color:#1B5E4B; }
    .badge-amber{ background:#281E05; color:#F4A011; border-color:#5A4108; }
    .badge-red{   background:#2B0E0E; color:#E05858; border-color:#5B1A1A; }

    /* ---------- Section header ---------- */
    .section-header{
    color: var(--txt-muted); font-size: .75rem; font-weight: 800;
    letter-spacing: .12em; text-transform: uppercase;
    border-bottom: 1px solid var(--border); padding-bottom: 8px; margin: 24px 0 14px 0;
    }

    /* ---------- KPI cards ---------- */
    .kpi-card{
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 14px 16px 12px 16px;
    margin-bottom: 8px;
    box-shadow: 0 6px 16px rgba(0,0,0,.20);
    }
    .kpi-card .kpi-label{
    color: var(--txt-muted); font-size: .75rem; font-weight: 700;
    letter-spacing: .06em; text-transform: uppercase; margin-bottom: 6px;
    }
    .kpi-card .kpi-value{ font-size: 2.1rem; font-weight: 800; color: var(--txt-main); }
    .kpi-card .kpi-sub{ font-size: .78rem; color: var(--txt-muted); margin-top: 4px; }
    .kpi-delta-good{ color:#30C08D; font-weight: 700; font-size:.85rem; }
    .kpi-delta-bad{  color:#E05858; font-weight: 700; font-size:.85rem; }
    .kpi-delta-neu{  color:var(--txt-muted); font-weight: 700; font-size:.85rem; }

    /* ---------- Activity feed ---------- */
    .activity-row{
    display:flex; align-items:center; gap:10px;
    padding:10px 8px; border-bottom: 1px solid var(--border);
    font-size: .9rem; background: transparent;
    }
    .activity-source{ font-weight: 800; min-width: 160px; color: var(--txt-main); }
    .activity-age{ color: var(--txt-muted); min-width: 80px; }
    .activity-who{ color: #8BB6F5; }
    .activity-row:hover{ background: var(--bg-alt); border-radius: 8px; }

    /* ---------- Trend panels ---------- */
    .trend-panel{
    background: var(--bg-card); border: 1px solid var(--border); border-radius: 12px;
    padding: 14px 16px; margin: 6px 0; box-shadow: 0 6px 16px rgba(0,0,0,.20);
    }
    .trend-icon{ font-size: 2.0rem; line-height:1; }
    .trend-label{ font-size: 1.05rem; font-weight: 800; color: var(--txt-main); }
    .trend-stat{ font-size: .80rem; color: var(--txt-muted); margin-top: 2px; }

    /* ---------- Insight boxes ---------- */
    .insight-box{
    border-radius:10px; padding:10px 14px; margin:6px 0; font-size:.9rem; line-height:1.55;
    border: 1px solid var(--border);
    }
    .insight-neutral{ background:#0F1B28; color:var(--txt-main); }
    .insight-good{    background:#0F251E; color:var(--txt-main); }
    .insight-warn{    background:#2A1F0E; color:var(--txt-main); }
    .insight-bad{     background:#2B0E0E; color:var(--txt-main); }

    /* ---------- Streamlit bits ---------- */
    .block-container{ padding-top: 1.0rem; }
    [data-baseweb="select"] *{ color: var(--txt-main)!important; }
    </style>
    """, unsafe_allow_html=True)


    # ── Helpers ───────────────────────────────────────────────────────────────
    def iso_week_to_range(iso_str: str) -> str:
        if "-W" not in iso_str:
            return iso_str
        try:
            year, week = iso_str.split("-W")
            start = datetime.fromisocalendar(int(year), int(week), 1)
            end   = datetime.fromisocalendar(int(year), int(week), 7)
            return (f"{start.strftime('%d')}–{end.strftime('%d %b %Y')}"
                    if start.month == end.month
                    else f"{start.strftime('%d %b')}–{end.strftime('%d %b %Y')}")
        except Exception:
            return iso_str

    def iso_to_date(iso_str: str):
        if "-W" not in iso_str:
            return pd.to_datetime(iso_str, errors="coerce")
        try:
            y, w = iso_str.split("-W")
            return datetime.fromisocalendar(int(y), int(w), 1)
        except Exception:
            return pd.NaT

    def _age_label(ts_str: str) -> str:
        try:
            ts = pd.to_datetime(ts_str)
            d  = (datetime.now() - ts.to_pydatetime().replace(tzinfo=None)).days
            if d == 0:  return "today"
            if d == 1:  return "yesterday"
            if d < 7:   return f"{d}d ago"
            if d < 31:  return f"{d//7}w ago"
            return f"{d//30}mo ago"
        except Exception:
            return "—"

    def _last_log(logs):
        if not logs: return None
        try:
            return sorted(logs, key=lambda l: l.get("timestamp",""), reverse=True)[0]
        except Exception:
            return logs[0]

    def _prev_log(logs):
        if len(logs) < 2: return None
        try:
            return sorted(logs, key=lambda l: l.get("timestamp",""), reverse=True)[1]
        except Exception:
            return None

    def _html_kpi(label, value, sub, delta=None, good=None, accent="#2E75B6"):
        delta_html = ""
        if delta is not None:
            cls  = "kpi-delta-good" if good else ("kpi-delta-bad" if good is False else "kpi-delta-neu")
            arr  = "▲" if (good is False) else ("▼" if good else "●")
            delta_html = f'''<div class="{cls}">{arr} {delta}</div>'''
        return f'''
        <div class="kpi-card" style="border-left-color:{accent}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            {delta_html}
            <div class="kpi-sub">{sub}</div>
        </div>'''

    # ── Load logs ─────────────────────────────────────────────────────────────
    sap_logs = list_run_logs("sap",       limit=500) or []
    atl_logs = list_run_logs("atl",       limit=500) or []
    dc_logs  = list_run_logs("dataclean", limit=500) or []

    has_any  = bool(sap_logs or atl_logs or dc_logs)

    sap_last = _last_log(sap_logs);  sap_prev = _prev_log(sap_logs)
    atl_last = _last_log(atl_logs);  atl_prev = _prev_log(atl_logs)
    dc_last  = _last_log(dc_logs);   dc_prev  = _prev_log(dc_logs)

    sap_age  = _age_label(sap_last.get("timestamp","")) if sap_last else "never"
    atl_age  = _age_label(atl_last.get("timestamp","")) if atl_last else "never"
    dc_age   = _age_label(dc_last.get("timestamp",""))  if dc_last  else "never"

    # ── A) Overall health badge ───────────────────────────────────────────────
    def _health(logs, key):
        if not logs or len(logs) < 2: return "neutral"
        cur  = _last_log(logs).get("metrics",{}).get(key, 0)
        prev = _prev_log(logs).get("metrics",{}).get(key, 0)
        if prev == 0: return "neutral"
        pct = (cur - prev) / prev * 100
        if pct < -5:   return "good"
        if pct < 10:   return "neutral"
        return "bad"

    sap_health = _health(sap_logs, "total_errors")
    atl_health = _health(atl_logs, "PT01_rows")
    dc_health  = _health(dc_logs,  "outdated")

    statuses = [sap_health, atl_health, dc_health]
    if "bad"     in statuses: overall = ("🔴 Needs Attention", "badge-red",   CLR["red"])
    elif "neutral" in statuses: overall = ("🟡 Monitoring",      "badge-amber", CLR["amber"])
    else:                       overall = ("🟢 On Track",         "badge-green", CLR["green"])

    run_date = datetime.now().strftime("%d %b %Y, %H:%M")

    # ── A) Banner ─────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="exec-banner">
        <h1>📈 Validation Hub — Evolution Dashboard
            <span class="badge {overall[1]}">{overall[0]}</span>
        </h1>
        <p class="subtitle">
            Generated {run_date} &nbsp;·&nbsp;
            SAP last run: <strong>{sap_age}</strong> &nbsp;·&nbsp;
            PLM last run: <strong>{atl_age}</strong> &nbsp;·&nbsp;
            Cleanups last run: <strong>{dc_age}</strong>
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Onboarding ────────────────────────────────────────────────────────────
    if not has_any:
        st.markdown("""
        <div class="insight-box insight-neutral">
            <strong>No execution history found yet.</strong><br>
            Run at least one validation to start seeing data here:
            <ul style="margin:6px 0 0 0">
                <li><strong>📦 SAP Material Validation</strong> → appears as <em>SAP</em> series</li>
                <li><strong>🚗 PLM Workflow</strong> → appears as <em>PLM Workflow</em> series</li>
                <li><strong>🧹 Data Cleanups</strong> → appears as <em>Cleanups</em> series</li>
            </ul>
            Each run is automatically saved to <code>.history/</code> and picked up here.
        </div>
        """, unsafe_allow_html=True)
        return

    # ── B) Custom KPI cards ───────────────────────────────────────────────────
    st.markdown('<div class="section-header">Current Status — Last Run per Source</div>',
                unsafe_allow_html=True)

    n_cols = sum([bool(sap_logs), bool(atl_logs), bool(dc_logs)])
    kpi_cols = st.columns(max(n_cols, 1))
    col_i = 0

    if sap_logs:
        errs  = sap_last.get("metrics",{}).get("total_errors", 0) if sap_last else 0
        mats  = sap_last.get("metrics",{}).get("materials_with_error", 0) if sap_last else 0
        user  = sap_last.get("user", "—") if sap_last else "—"
        prev_e = sap_prev.get("metrics",{}).get("total_errors", 0) if sap_prev else None
        if prev_e is not None:
            diff  = errs - prev_e
            dstr  = f"{abs(diff):,} {'↑' if diff>0 else '↓'} vs prev"
            good  = diff <= 0
        else:
            dstr, good = None, None
        with kpi_cols[col_i]:
            st.markdown(_html_kpi("📦 SAP — Last Run Errors", f"{errs:,}",
                                   f"{mats:,} materials &nbsp;·&nbsp; {user} &nbsp;·&nbsp; {sap_age}",
                                   dstr, good, CLR["sap"]), unsafe_allow_html=True)
        col_i += 1

    if atl_logs:
        pt01  = atl_last.get("metrics",{}).get("PT01_rows", 0) if atl_last else 0
        pt02  = atl_last.get("metrics",{}).get("PT02_rows", 0) if atl_last else 0
        tln   = (atl_last.get("metrics",{}).get("PT01_with_TLN",0) +
                 atl_last.get("metrics",{}).get("PT02_with_TLN",0)) if atl_last else 0
        user  = atl_last.get("user","—") if atl_last else "—"
        total = pt01 + pt02
        prev_t = (atl_prev.get("metrics",{}).get("PT01_rows",0) +
                  atl_prev.get("metrics",{}).get("PT02_rows",0)) if atl_prev else None
        if prev_t is not None:
            diff  = total - prev_t
            dstr  = f"{abs(diff):,} {'↑' if diff>0 else '↓'} vs prev"
            good  = diff <= 0
        else:
            dstr, good = None, None
        with kpi_cols[col_i]:
            st.markdown(_html_kpi("🚗 PLM — Open Change Numbers", f"{total:,}",
                                   f"{tln:,} with TLN &nbsp;·&nbsp; {user} &nbsp;·&nbsp; {atl_age}",
                                   dstr, good, CLR["plm01"]), unsafe_allow_html=True)
        col_i += 1

    if dc_logs:
        out   = dc_last.get("metrics",{}).get("outdated", 0) if dc_last else 0
        mats  = dc_last.get("metrics",{}).get("materials", 0) if dc_last else 0
        user  = dc_last.get("user","—") if dc_last else "—"
        prev_o = dc_prev.get("metrics",{}).get("outdated", 0) if dc_prev else None
        if prev_o is not None:
            diff  = out - prev_o
            dstr  = f"{abs(diff):,} {'↑' if diff>0 else '↓'} vs prev"
            good  = diff <= 0
        else:
            dstr, good = None, None
        with kpi_cols[col_i]:
            st.markdown(_html_kpi("🧹 Cleanups — Outdated", f"{out:,}",
                                   f"{mats:,} materials &nbsp;·&nbsp; {user} &nbsp;·&nbsp; {dc_age}",
                                   dstr, good, CLR["cleanups"]), unsafe_allow_html=True)

    # ── C) Activity feed — always visible, compact ────────────────────────────
    st.markdown('<div class="section-header">Recent Activity</div>', unsafe_allow_html=True)

    icon_map  = {"sap":"📦","atl":"🚗","dataclean":"🧹"}
    lbl_map   = {"sap":"SAP Validation","atl":"PLM Workflow","dataclean":"Data Cleanups"}
    vkey_map  = {"sap":"total_errors","atl":"PT01_rows","dataclean":"outdated"}
    vlbl_map  = {"sap":"errors","atl":"PT01 rows","dataclean":"outdated"}
    acc_map   = {"sap":CLR["sap"],"atl":CLR["plm01"],"dataclean":CLR["cleanups"]}

    feed_html = ''
    for kind, logs in [("sap",sap_logs),("atl",atl_logs),("dataclean",dc_logs)]:
        for log in sorted(logs, key=lambda l: l.get("timestamp",""), reverse=True)[:3]:
            ts  = log.get("timestamp","")
            val = log.get("metrics",{}).get(vkey_map[kind], 0)
            scope = log.get("plant_filter", log.get("plant_scope","ALL"))
            feed_html += f'''
            <div class="activity-row">
                <span class="activity-source" style="border-left:3px solid {acc_map[kind]};padding-left:8px">
                    {icon_map[kind]} {lbl_map[kind]}
                </span>
                <span class="activity-age">{_age_label(ts)}</span>
                <span class="activity-who">{log.get("user","—")}</span>
                <span style="color:{CLR['text_muted']};font-size:0.78rem">{scope}</span>
                <span class="activity-val" style="color:{CLR['text_main']};font-weight:700;">
                    {val:,} {vlbl_map[kind]}
                </span>
            </div>'''

    if feed_html:
        st.markdown(
            f'''
            <div style="
                background:{CLR['bg_panel']};
                border:1px solid {CLR['border']};
                border-radius:12px;
                padding:14px 18px;
                box-shadow:0 6px 16px rgba(0,0,0,.25);
            ">
                {feed_html}
            </div>
            ''',
            unsafe_allow_html=True
        )
    else:
        st.caption("No activity yet.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Sidebar controls ──────────────────────────────────────────────────────
    available_sources = []
    if sap_logs: available_sources.append("SAP")
    if atl_logs: available_sources.append("PLM Workflow")
    _has_tln = any(
        l.get("metrics",{}).get("PT01_with_TLN",0)+l.get("metrics",{}).get("PT02_with_TLN",0)>0
        for l in atl_logs
    )
    if atl_logs and _has_tln: available_sources.append("PLM TLN")
    if dc_logs: available_sources.append("Cleanups")

    st.sidebar.markdown("### 📌 Data Sources")
    selected_sources = st.sidebar.multiselect("Sources:", options=available_sources, default=available_sources)
    if not selected_sources:
        st.info("Please select at least one source.")
        return

    st.sidebar.markdown("### ⌛ Periodicity")
    mode = st.sidebar.selectbox("Group by:", ["Daily","Weekly","Monthly"], index=1)

    st.sidebar.markdown("### 🧮 Aggregation")
    agg_mode = st.sidebar.selectbox("Values per period:",
        ["Latest per period (last log)","Sum per period","Average per period"], index=0)

    st.sidebar.markdown("### 📏 Scale")
    scale_choice = st.sidebar.selectbox("Y axis scale:",
        ["Linear","Log","Dual axis (SAP vs others)"], index=2)

    normalize    = st.sidebar.checkbox("Normalize to % of max (0–100)", value=False)
    smoothing_on = st.sidebar.checkbox("Apply smoothing (rolling mean)", value=False)
    window       = st.sidebar.slider("Window (periods)", 2, 7, 3) if smoothing_on else 3

    # ── Build dataset ─────────────────────────────────────────────────────────
    entries = []
    if "SAP" in selected_sources:
        for log in sap_logs:
            entries.append({"ts":pd.to_datetime(log.get("timestamp")),
                             "plant":log.get("plant_filter","ALL"),
                             "source":"SAP",
                             "value":log.get("metrics",{}).get("total_errors",0)})
    if "PLM Workflow" in selected_sources:
        for log in atl_logs:
            ts = pd.to_datetime(log.get("timestamp")); m = log.get("metrics",{})
            entries.append({"ts":ts,"plant":"PT01","source":"PLM Workflow PT01","value":m.get("PT01_rows",0)})
            entries.append({"ts":ts,"plant":"PT02","source":"PLM Workflow PT02","value":m.get("PT02_rows",0)})
            if "PLM TLN" in selected_sources:
                if m.get("PT01_with_TLN",0)>0:
                    entries.append({"ts":ts,"plant":"PT01","source":"PLM TLN PT01","value":m.get("PT01_with_TLN",0)})
                if m.get("PT02_with_TLN",0)>0:
                    entries.append({"ts":ts,"plant":"PT02","source":"PLM TLN PT02","value":m.get("PT02_with_TLN",0)})
    if "Cleanups" in selected_sources:
        for log in dc_logs:
            entries.append({"ts":pd.to_datetime(log.get("timestamp")),
                             "plant":log.get("plant_filter","ALL"),
                             "source":"Cleanups",
                             "value":log.get("metrics",{}).get("outdated",0)})

    if not entries:
        st.info("No data for the selected sources.")
        return

    df = pd.DataFrame(entries)

    st.sidebar.markdown("### 📅 Date Range")
    dmin, dmax = df["ts"].min().date(), df["ts"].max().date()
    se = st.sidebar.date_input("Period:", value=(dmin,dmax), min_value=dmin, max_value=dmax)
    start_date, end_date = (se[0], se[1]) if (isinstance(se,(list,tuple)) and len(se)==2) else (dmin,dmax)
    df = df[(df["ts"].dt.date >= start_date) & (df["ts"].dt.date <= end_date)]

    plants = sorted(df["plant"].dropna().unique().tolist())
    real_plants = [p for p in plants if p != "ALL"]
    sel_plants = st.sidebar.multiselect("Plants:", ["ALL"]+real_plants, default=["ALL"])
    df = df[df["plant"].isin((real_plants+["ALL"]) if "ALL" in sel_plants else sel_plants)]

    if df.empty:
        st.info("No data for the selected filters.")
        return

    # Period column
    if mode == "Daily":
        df["period"] = df["ts"].dt.strftime("%Y-%m-%d")
    elif mode == "Weekly":
        iso = df["ts"].dt.isocalendar()
        df["period"] = iso["year"].astype(str) + "-W" + iso["week"].astype(str)
    else:
        df["period"] = df["ts"].dt.strftime("%Y-%m")

    value_col = "value"
    if normalize:
        df["_max"] = df.groupby("source")[value_col].transform(lambda s: s.max() if s.max()>0 else 1.0)
        df["value_norm"] = df[value_col] / df["_max"] * 100.0
        value_col = "value_norm"

    df = df.sort_values(["period","source","plant","ts"])
    if agg_mode.startswith("Latest"):
        grp = df.groupby(["period","source","plant"], as_index=False).tail(1)
        grouped = grp.groupby(["period","source","plant"], as_index=False)[value_col].sum()
    elif agg_mode.startswith("Sum"):
        grouped = df.groupby(["period","source","plant"], as_index=False)[value_col].sum()
    else:
        grouped = df.groupby(["period","source","plant"], as_index=False)[value_col].mean()

    grouped["sort_col"]      = grouped["period"].apply(iso_to_date)
    grouped["pretty_period"] = grouped["period"].apply(iso_week_to_range)
    grouped = grouped.sort_values("sort_col", kind="stable")

    plot_df = grouped.copy()
    if smoothing_on and len(plot_df["period"].unique()) >= window:
        plot_df["value_smooth"] = (
            plot_df.sort_values(["source","plant","sort_col"])
                   .groupby(["source","plant"])[value_col]
                   .transform(lambda s: s.rolling(window=window, min_periods=1).mean())
        )
        y_plot_col = "value_smooth"
    else:
        y_plot_col = value_col

    sum_by_period = (
        plot_df.groupby(["period","pretty_period","sort_col"], as_index=False)[y_plot_col]
               .sum().sort_values("sort_col", kind="stable")
               .rename(columns={y_plot_col:"value"})
    )

    # ── D) Per-series sparkline KPIs ──────────────────────────────────────────
    st.markdown('<div class="section-header">Trend per Series — Current Period</div>',
                unsafe_allow_html=True)
    all_srcs = sorted(plot_df["source"].unique())
    if all_srcs and PLOTLY_OK:
        n_spark = min(len(all_srcs), 4)
        spark_cols = st.columns(n_spark)
        for i, src in enumerate(all_srcs[:n_spark]):
            sd = (plot_df[plot_df["source"]==src]
                  .groupby(["period","pretty_period","sort_col"], as_index=False)[y_plot_col]
                  .sum().sort_values("sort_col"))
            if sd.empty: continue
            cur  = float(sd[y_plot_col].iloc[-1])
            prev = float(sd[y_plot_col].iloc[-2]) if len(sd)>=2 else None
            delta_v = cur - prev if prev is not None else None
            delta_str = (f"{'+' if delta_v>=0 else ''}{delta_v:,.0f}") if delta_v is not None else "—"
            good = delta_v <= 0 if delta_v is not None else True
            color = COLOR_MAP.get(src, "#2E75B6")

            with spark_cols[i % n_spark]:
                # mini sparkline
                spark = go.Figure(go.Scatter(
                    x=sd["pretty_period"].tolist(), y=sd[y_plot_col].tolist(),
                    mode="lines", line=dict(color=color, width=2),
                    fill="tozeroy", fillcolor=hex_to_rgba(color, 0.10)
                ))
                spark.update_layout(
                    height=60, margin=dict(l=0,r=0,t=0,b=0),
                    showlegend=False,
                    paper_bgcolor=CLR["bg_card"],
                    plot_bgcolor=CLR["bg_card"],
                    font=dict(color=CLR["text_main"], size=10)
                )
                spark.update_xaxes(visible=False); spark.update_yaxes(visible=False)
                st.plotly_chart(spark, width='stretch', config={"displayModeBar":False})
                delta_color = CLR["green"] if good else CLR["red"]
                st.markdown(
                    f'<div style="text-align:center">' +
                    f'<span style="font-size:1.1rem;font-weight:700;color:{CLR["text_main"]}">{cur:,.0f}</span>' +
                    f'&nbsp;<span style="font-size:0.75rem;color:{delta_color};font-weight:600">{delta_str}</span><br>' +
                    f'<span style="font-size:0.7rem;color:{CLR["text_muted"]}">{src}</span>' +
                    '</div>',
                    unsafe_allow_html=True
                )
    elif all_srcs:
        for src in all_srcs:
            sd = plot_df[plot_df["source"]==src].groupby(["period","sort_col"],as_index=False)[y_plot_col].sum().sort_values("sort_col")
            if not sd.empty:
                st.metric(src, f"{float(sd[y_plot_col].iloc[-1]):,.0f}")

    # ── E) Main chart ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Evolution Over Time</div>', unsafe_allow_html=True)

    use_dual = scale_choice.startswith("Dual")
    unique_sorted = sum_by_period["pretty_period"].tolist()
    n = len(unique_sorted)
    tick_step = 1 if n<=12 else (2 if n<=20 else (3 if n<=35 else 4))

    if PLOTLY_OK:
        fig = make_subplots(specs=[[{"secondary_y": True}]]) if use_dual else go.Figure()

        for (src, pl), grp in plot_df.groupby(["source","plant"], sort=False):
            grp_s  = grp.sort_values("sort_col")
            x_vals = grp_s["pretty_period"].tolist()
            y_vals = grp_s[y_plot_col].tolist()
            color  = COLOR_MAP.get(src, "#888")
            is_sap = src.startswith("SAP")

            # Area fill for SAP; line only for others
            trace = go.Scatter(
                x=x_vals, y=y_vals, mode="lines+markers",
                name=f"{src}",
                marker=dict(size=7, color=color),
                line=dict(color=color, width=2.5),
                fill="tozeroy" if is_sap else "none",
                fillcolor=hex_to_rgba(color, 0.10) if is_sap else None,
                hovertemplate=f"<b>{src}</b><br>Period: %{{x}}<br>Value: %{{y:,.0f}}<extra></extra>"
            )
            if use_dual:
                fig.add_trace(trace, secondary_y=not is_sap)
            else:
                fig.add_trace(trace)

            # Auto-annotate peak and valley for each series
            if len(y_vals) >= 3:
                max_i = int(np.argmax(y_vals))
                min_i = int(np.argmin(y_vals))
                for ann_i, ann_text, ann_ay in [(max_i, "▲", -30), (min_i, "▼", 30)]:
                    ann = dict(
                        x=x_vals[ann_i], y=y_vals[ann_i],
                        text=f"{ann_text} {y_vals[ann_i]:,.0f}",
                        showarrow=True, arrowhead=2, arrowcolor=color,
                        arrowsize=0.8, ax=0, ay=ann_ay,
                        font=dict(size=10, color=color),
                        bgcolor="white", bordercolor=color, borderwidth=1
                    )
                    if use_dual:
                        ann["yref"] = "y2" if not is_sap else "y"
                    fig.add_annotation(ann)

        # Mean reference line for SAP (left axis)
        sap_data = plot_df[plot_df["source"]=="SAP"]
        if not sap_data.empty:
            sap_mean = float(sap_data[y_plot_col].mean())
            fig.add_hline(
                y=sap_mean, line_dash="dot", line_color=CLR["sap"],
                opacity=0.5,
                annotation_text=f"SAP avg: {sap_mean:,.0f}",
                annotation_position="bottom right",
                annotation_font_size=10,
                secondary_y=False
            )

        if scale_choice == "Log":
            fig.update_yaxes(type="log")
        if use_dual:
            fig.update_yaxes(title_text="SAP Errors",          secondary_y=False, showgrid=True)
            fig.update_yaxes(title_text="PLM / Cleanups Count", secondary_y=True,  showgrid=False)

        fig.update_xaxes(
            tickmode="array", tickvals=unique_sorted[::tick_step], tickangle=-45,
            showgrid=True, gridcolor=CLR["grid"], color=CLR["text_muted"]
        )
        fig.update_yaxes(showgrid=True, gridcolor=CLR["grid"], color=CLR["text_muted"])
        fig.update_layout(
            title=dict(text="Validation Metrics — Historical Evolution",
                    font=dict(size=16, color=CLR["text_main"]), x=0.02, xanchor="left"),
            height=520,
            margin=dict(l=40, r=40, t=70, b=80),
            legend=dict(orientation="h", yanchor="bottom", y=-0.28,
                        xanchor="center", x=0.5,
                        bgcolor=CLR["bg_panel"],
                        bordercolor=CLR["border"], borderwidth=1,
                        font=dict(color=CLR["text_main"])),
            paper_bgcolor=CLR["bg_panel"],
            plot_bgcolor=CLR["bg_panel"],
            font=dict(family="Inter, Segoe UI, sans-serif", size=12, color=CLR["text_main"])
        )
        st.plotly_chart(fig, width='stretch')
    else:
        fig, ax_l = plt.subplots(figsize=(10,5))
        ax_r = ax_l.twinx() if use_dual else None
        for (src, pl), grp in plot_df.groupby(["source","plant"], sort=False):
            grp_s = grp.sort_values("sort_col")
            color = COLOR_MAP.get(src, None)
            ax = ax_r if (use_dual and not src.startswith("SAP") and ax_r) else ax_l
            ax.plot(grp_s["pretty_period"].tolist(), grp_s[y_plot_col].tolist(),
                    marker="o", label=f"{src} ({pl})", color=color)
        ax_l.set_title("Validation Metrics — Historical Evolution")
        ax_l.set_xlabel("Period"); ax_l.set_ylabel("Value")
        if ax_r: ax_r.set_ylabel("PLM / Cleanups")
        ax_l.legend(loc="upper left"); plt.xticks(rotation=45)
        st.pyplot(fig)

    # ── Export ────────────────────────────────────────────────────────────────
    with st.expander("📤 Export Data", expanded=False):
        export_buf = io.BytesIO()
        with pd.ExcelWriter(export_buf, engine="openpyxl") as writer:
            sum_by_period.to_excel(writer, sheet_name="Historical_Summary", index=False)
            plot_df.to_excel(writer, sheet_name="All_Series", index=False)
        st.download_button(
            "⬇️ Download Evolution Report (Excel)",
            data=export_buf.getvalue(),
            file_name=f"Evolution_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ── F) Trend panel ────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="section-header">Trend & Anomaly Analysis</div>',
        unsafe_allow_html=True
    )

    import textwrap

    def md_html(s: str):
        """Renderiza HTML sem indentação para evitar code blocks em Markdown."""
        st.markdown(textwrap.dedent(s), unsafe_allow_html=True)

    def split_label(label: str):
        """Devolve (icon, text) de um rótulo como '📈 Strong Uptrend'."""
        label = label or ""
        parts = label.split()
        if len(parts) >= 2:
            return parts[0], " ".join(parts[1:])
        # fallback seguro
        return "➖", (label if label else "Sideways")

    trend_source = st.selectbox(
        "Analyse source:", options=all_srcs, index=0, key="trend_source_select"
    )

    src_series = (
        plot_df[plot_df["source"] == trend_source]  # Fixed: removed space in 'plot_df'
        .groupby(["period", "pretty_period", "sort_col"], as_index=False)[y_plot_col]
        .sum()
        .sort_values("sort_col")
        .rename(columns={y_plot_col: "value"})
    )

    slope = r2 = 0.0
    trend_label = "—"
    y_arr = np.array([])

    if len(src_series) >= 3:
        y_arr = src_series["value"].to_numpy(dtype=float)
        x_arr = np.arange(len(y_arr))
        lr = linregress(x_arr, y_arr)
        slope = lr.slope
        r2 = lr.rvalue ** 2
        thr = np.std(y_arr) * 0.15
        if slope > thr:
            trend_label = "📈 Strong Uptrend"; trend_cls = "insight-bad"
        elif slope > 0:
            trend_label = "↗ Weak Uptrend"; trend_cls = "insight-warn"
        elif slope < -thr:
            trend_label = "📉 Strong Downtrend"; trend_cls = "insight-good"
        elif slope < 0:
            trend_label = "↘ Weak Downtrend"; trend_cls = "insight-good"
        else:
            trend_label = "➖ Sideways"; trend_cls = "insight-neutral"
    else:
        # garantir que y_arr existe para downstream (mesmo que vazio)
        y_arr = src_series["value"].to_numpy(dtype=float) if not src_series.empty else np.array([])
        trend_cls = "insight-neutral"

    tc1, tc2, tc3 = st.columns([1, 1, 2])

    # COLUMN 1 — cartão com ícone + label
    with tc1:
        icon, text = split_label(trend_label)
        md_html(f"""
        <div class="trend-panel" style="
            background:{CLR['bg_card']};
            border:1px solid {CLR['border']};
            border-radius:12px;
            padding:14px 16px;
            box-shadow:0 6px 16px rgba(0,0,0,.22);
        ">
            <div class="trend-icon" style="font-size:2.1rem; line-height:1">{icon}</div>
            <div class="trend-label" style="font-size:1.1rem;font-weight:800;color:{CLR['text_main']}">
                {text}
            </div>
            <div class="trend-stat" style="color:{CLR['text_muted']};font-size:0.8rem;">
                {trend_source}
            </div>
        </div>
        """)

    # COLUMN 2 — slope, R² e contagem de pontos
    with tc2:
        md_html(f"""
        <div class="trend-panel" style="
            background:{CLR['bg_card']};
            border:1px solid {CLR['border']};
            border-radius:12px;
            padding:14px 16px;
            box-shadow:0 6px 16px rgba(0,0,0,.22);
        ">
            <div class="trend-stat" style="color:{CLR['text_muted']}">Slope</div>
            <div class="trend-label" style="font-size:1.5rem;font-weight:800;color:{CLR['text_main']}">
                {slope:+.3f}
            </div>
            <div class="trend-stat" style="color:{CLR['text_muted']}">R² confidence</div>
            <div class="trend-label" style="font-size:1.5rem;font-weight:800;color:{CLR['text_main']}">
                {r2:.3f}
            </div>
            <div class="trend-stat" style="color:{CLR['text_muted']}">
                {len(y_arr)} data points
            </div>
        </div>
        """)

    # COLUMN 3 — mini‑gráfico com (opcional) linha de tendência
    with tc3:
        if PLOTLY_OK and not src_series.empty:
            periods_list = src_series["pretty_period"].tolist()
            values_list  = y_arr.tolist()

            spark_t = go.Figure()
            spark_t.add_trace(go.Scatter(
                x=periods_list,
                y=values_list,
                mode="lines+markers",
                name="Actual",
                marker=dict(size=6, color=COLOR_MAP.get(trend_source, "#6EC6FF"),
                            line=dict(width=1, color="white")),
                line=dict(color=COLOR_MAP.get(trend_source, "#6EC6FF"), width=2)
            ))

            # só desenhar linha de tendência quando houver dados suficientes
            if len(values_list) >= 3:
                x_fit = np.arange(len(values_list))
                y_fit = lr.intercept + lr.slope * x_fit
                spark_t.add_trace(go.Scatter(
                    x=periods_list,
                    y=y_fit.tolist(),
                    mode="lines",
                    name="Trend line",
                    line=dict(color=CLR["red"], width=1.5, dash="dot")
                ))

            spark_t.update_layout(
                height=170,
                margin=dict(l=0, r=0, t=5, b=35),
                showlegend=False,
                paper_bgcolor=CLR["bg_card"],
                plot_bgcolor=CLR["bg_card"],
                font=dict(size=10, color=CLR["text_main"])
            )
            spark_t.update_xaxes(tickangle=-45, showgrid=False, color=CLR["text_muted"])
            spark_t.update_yaxes(showgrid=True, gridcolor=CLR["grid"], color=CLR["text_muted"])
            st.plotly_chart(spark_t, width='stretch', config={"displayModeBar": False})
        else:
            md_html(
                f'<div class="insight-box insight-neutral">'
                f'Not enough data points for <strong>{trend_source}</strong> (need ≥ 3 periods).'
                f'</div>'
            )

    # ── G) Anomaly detection ──────────────────────────────────────────────────
    st.markdown(f"##### 🚨 Anomaly Detection — {trend_source}")
    anomalies_df = pd.DataFrame()

    if len(src_series) < 5:
        st.markdown('<div class="insight-box insight-neutral">Need at least 5 periods to run anomaly detection.</div>',
                    unsafe_allow_html=True)
    else:
        with st.expander("Methodology", expanded=False):
            st.write("**Z-score**: distance from mean in standard deviations. "
                     "**Modified Z-score (MAD)**: robust alternative using median and MAD.")

        c1, c2, c3, c4 = st.columns(4)
        with c1: metric   = st.selectbox("Scoring", ["Modified Z (MAD)","Z-score"], index=0, key="anom_metric")
        with c2: thr_mild = st.slider("Mild ≥",     1.0, 3.0, 1.8, 0.1, key="anom_mild")
        with c3: thr_mod  = st.slider("Moderate ≥", 2.0, 4.0, 2.5, 0.1, key="anom_mod")
        with c4: thr_sev  = st.slider("Severe ≥",   3.0, 5.0, 3.5, 0.1, key="anom_sev")

        vals    = src_series["value"].to_numpy(dtype=float)
        periods = src_series["pretty_period"].astype(str).tolist()
        mean_v  = float(np.mean(vals));   std_v  = float(np.std(vals))   or 1e-6
        median_v= float(np.median(vals)); mad_v  = float(np.median(np.abs(vals-median_v))) or 1e-6
        z_scores     = (vals - mean_v) / std_v
        mod_z_scores = 0.6745 * (vals - median_v) / mad_v
        score        = mod_z_scores if metric=="Modified Z (MAD)" else z_scores

        severities = []
        for s in np.abs(score):
            if   s >= thr_sev:  severities.append("🔥 Severe")
            elif s >= thr_mod:  severities.append("🟠 Moderate")
            elif s >= thr_mild: severities.append("🟡 Mild")
            else:               severities.append("")

        anomalies_df = pd.DataFrame({
            "period":periods, "value":vals,
            "zscore":z_scores, "mod_z":mod_z_scores, "severity":severities
        })
        anomalies_df = anomalies_df[anomalies_df["severity"]!=""].copy()
        key_col = "mod_z" if metric=="Modified Z (MAD)" else "zscore"
        anomalies_df["abs_score"] = anomalies_df[key_col].abs()
        anomalies_df.sort_values("abs_score", ascending=False, inplace=True)

        if PLOTLY_OK:
            # G) Shaded ±1σ band around mean
            fig_a = go.Figure()
            color_src = COLOR_MAP.get(trend_source, "#2E75B6")
            x_periods = src_series["pretty_period"].tolist()

            # ±1σ shaded band
            fig_a.add_trace(go.Scatter(
                x=x_periods + x_periods[::-1],
                y=list(vals + std_v) + list((vals - std_v)[::-1]),
                fill="toself", fillcolor=hex_to_rgba(color_src, 0.08),
                line=dict(color="rgba(0,0,0,0)"),
                name="±1σ band", showlegend=True
            ))
            # Main line
            fig_a.add_trace(go.Scatter(
                x=x_periods, y=vals.tolist(), mode="lines+markers",
                name=trend_source, line=dict(color=color_src, width=2),
                marker=dict(size=6, color=color_src)
            ))
            # Mean reference
            fig_a.add_hline(y=mean_v, line_dash="dot", line_color="#6B7280",
                            annotation_text=f"mean: {mean_v:,.0f}",
                            annotation_font_size=10)
            # Anomaly markers
            if not anomalies_df.empty:
                cmap = {"🟡 Mild":"#F4D03F","🟠 Moderate":"#E67E22","🔥 Severe":"#E74C3C"}
                anomalies_df["color"] = anomalies_df["severity"].map(cmap)
                fig_a.add_trace(go.Scatter(
                    x=anomalies_df["period"].tolist(),
                    y=anomalies_df["value"].tolist(),
                    mode="markers",
                    marker=dict(size=13, color=anomalies_df["color"].tolist(),
                                symbol="diamond", line=dict(width=1, color="white")),
                    name="Anomalies",
                    text=[f"{sev}<br>{p}<br>{v:,.0f}"
                          for sev, p, v in zip(anomalies_df["severity"],
                                               anomalies_df["period"],
                                               anomalies_df["value"])],
                    hovertemplate="%{text}<extra></extra>"
                ))
            
            fig_a.update_layout(
                title=dict(text=f"Anomaly Detection — {trend_source}",
                        font=dict(size=14, color=CLR["text_main"]), x=0.02, xanchor="left"),
                height=360, margin=dict(l=40,r=40,t=50,b=60),
                paper_bgcolor=CLR["bg_panel"], plot_bgcolor=CLR["bg_panel"],
                legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center",
                            bgcolor=CLR["bg_panel"], bordercolor=CLR["border"], borderwidth=1,
                            font=dict(color=CLR["text_main"])),
                font=dict(size=11, color=CLR["text_main"])
            )
            fig_a.update_xaxes(tickangle=-45, showgrid=True, gridcolor=CLR["grid"], color=CLR["text_muted"])
            fig_a.update_yaxes(showgrid=True, gridcolor=CLR["grid"], color=CLR["text_muted"])
            st.plotly_chart(fig_a, width='stretch')
        else:
            st.line_chart(src_series.set_index("pretty_period")["value"])

        if anomalies_df.empty:
            st.markdown(f'<div class="insight-box insight-good">✅ No anomalies detected for <strong>{trend_source}</strong> with current thresholds.</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="insight-box insight-warn">⚠️ <strong>{len(anomalies_df)} anomalies detected</strong> for <strong>{trend_source}</strong>.</div>',
                        unsafe_allow_html=True)
            max_n = int(min(10, len(anomalies_df)))
            show_n = st.slider("Show top N", 1, max_n, min(5,max_n), key="anom_show_n") if max_n>1 else 1
            for _, a in anomalies_df.head(show_n).iterrows():
                cls = "insight-bad" if "Severe" in a["severity"] else ("insight-warn" if "Moderate" in a["severity"] else "insight-neutral")
                st.markdown(
                    f'<div class="insight-box {cls}">{a["severity"]} — <strong>{a["period"]}</strong> → ' +
                    f'<code>{int(a["value"]):,}</code> (Z={a["zscore"]:.2f}, MAD={a["mod_z"]:.2f})</div>',
                    unsafe_allow_html=True
                )
            st.download_button(
                "⬇️ Download anomalies (CSV)",
                data=anomalies_df.drop(columns=["abs_score"]).to_csv(index=False).encode("utf-8"),
                file_name=f"anomalies_{trend_source.replace(' ','_')}.csv",
                mime="text/csv"
            )

    # ── H) Auto-Insights ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Auto-Insights</div>', unsafe_allow_html=True)
    insights = []

    if len(src_series) >= 2:
        initial = float(src_series["value"].iloc[0])
        final   = float(src_series["value"].iloc[-1])
        pct     = (final - initial) / initial * 100 if initial != 0 else None
        if pct is not None:
            direction = "increased" if pct > 0 else "decreased"
            cls = "insight-bad" if pct > 10 else ("insight-good" if pct < -5 else "insight-neutral")
            insights.append((cls,
                f"<strong>{trend_source}</strong> {direction} by <strong>{abs(pct):.1f}%</strong> "
                f"across the selected period (from <strong>{initial:,.0f}</strong> to <strong>{final:,.0f}</strong>)."))
        else:
            insights.append(("insight-neutral",
                f"<strong>{trend_source}</strong> remained stable across the selected period."))

    if len(src_series) >= 3 and trend_label != "—":
        cls = trend_cls if "trend_cls" in dir() else "insight-neutral"
        insights.append((cls,
            f"Regression trend: <strong>{trend_label}</strong> — "
            f"slope <code>{slope:.3f}</code>, confidence R² <code>{r2:.3f}</code>."))

    if not anomalies_df.empty:
        severe = anomalies_df[anomalies_df["severity"].str.contains("Severe")]
        if len(severe):
            insights.append(("insight-bad",
                f"⚠️ <strong>{len(severe)} severe anomalies</strong> detected — worth investigating immediately."))
        else:
            insights.append(("insight-warn",
                "Only mild/moderate anomalies detected — monitor but no immediate action required."))
    elif len(src_series) >= 5:
        insights.append(("insight-good", "✅ No anomalies detected in this timeframe."))

    for cls, text in insights:
        st.markdown(f'<div class="insight-box {cls}">{text}</div>', unsafe_allow_html=True)
    if not insights:
        st.caption("Not enough data to generate insights.")