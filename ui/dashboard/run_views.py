"""KPIs e gráficos da execução corrente (SAP, ATL, Progress)."""
from __future__ import annotations

import streamlit as st
import pandas as pd

from ui.dashboard.common import PLOTLY_AVAILABLE, plt, px, style_plotly

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

        cols = ['Trigger ID', 'Plant', 'Timing Letter Number', '_aging']
        if 'Source' in df_all.columns:
            cols.insert(1, 'Source')
        cols = [c for c in cols if c in df_all.columns or c == '_aging']

        top_aging = (
            df_all.assign(_aging=aging_num)
            .dropna(subset=['_aging'])
            .sort_values('_aging', ascending=False)
            .drop_duplicates(subset=['Trigger ID'])
            .head(15)
            .loc[:, cols]
            .rename(columns={'_aging': 'Aging (days)'})
            .reset_index(drop=True)
        )

        if not top_aging.empty:
            if PLOTLY_AVAILABLE:
                hover_data = []
                if 'Source' in top_aging.columns:
                    hover_data.append('Source')
                if 'Timing Letter Number' in top_aging.columns:
                    hover_data.append('Timing Letter Number')

                fig = px.bar(top_aging, x='Trigger ID', y='Aging (days)', color='Plant',
                             title="Top 15 Oldest Change Numbers (by Aging)",
                             text='Aging (days)',
                             hover_data=hover_data or None)
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
