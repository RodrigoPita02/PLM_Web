
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

# Tenta usar Plotly; se não houver, faz fallback para Matplotlib
try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except Exception:
    PLOTLY_AVAILABLE = False
    # Importa matplotlib apenas se necessário (fallback)
    import matplotlib.pyplot as plt


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
      - Erros (View em falta) por código de View
    Usa padrão robusto para value_counts → (rename_axis + reset_index(name=...)).
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
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
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
            st.plotly_chart(fig, use_container_width=True)
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
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
        else:
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(by_mat['Material'], by_mat['Errors'], color="#2ca02c")
            ax.set_title(f"Top {top_n} Materials with the most errors")
            ax.set_xlabel("Material"); ax.set_ylabel("Errors")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

    # --- Views em falta por código de View ---
    if {'Justified_By_Missing_View', 'View'}.issubset(errors_df.columns):
        mv = errors_df.loc[errors_df['Justified_By_Missing_View']].copy()
        mv = mv[mv['View'].astype(str).str.strip().ne('Not aplicable')]
        if not mv.empty:
            by_view = (
                mv['View']
                .value_counts()
                .rename_axis('View')
                .reset_index(name='Errors')
            )
            if PLOTLY_AVAILABLE:
                fig = px.bar(by_view, x='View', y='Errors',
                             title="Errors (Missing View) due to View code",
                             text='Errors', color='View')
                fig.update_traces(textposition='outside')
                st.plotly_chart(fig, use_container_width=True)
            else:
                fig, ax = plt.subplots(figsize=(8, 4))
                ax.bar(by_view['View'], by_view['Errors'], color="#d62728")
                ax.set_title("Errors (Missing View) due to View code")
                ax.set_xlabel("View"); ax.set_ylabel("Errors")
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
        st.plotly_chart(fig, use_container_width=True)
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
            st.plotly_chart(fig, use_container_width=True)
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
            st.plotly_chart(fig, use_container_width=True)
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
            st.plotly_chart(fig, use_container_width=True)
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
        st.plotly_chart(fig, use_container_width=True)
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
        st.plotly_chart(fig, use_container_width=True)
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
            st.plotly_chart(fig, use_container_width=True)
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
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
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
            st.plotly_chart(fig, use_container_width=True)
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
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
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
        st.plotly_chart(fig, use_container_width=True)
    else:
        fig, ax = plt.subplots(figsize=(8, 4))
        for plant, grp in dsorted.groupby('plant_filter'):
            ax.plot(grp['ts'], grp['total_errors'], marker='o', label=plant)
        ax.set_title("Total Errors over time (MRP)")
        ax.set_xlabel("Time"); ax.set_ylabel("Errors")
        ax.legend()
        st.pyplot(fig)
