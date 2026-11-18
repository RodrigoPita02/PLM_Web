"""Evolution Dashboard (visão executiva com histórico, tendências e atribuições)."""
from __future__ import annotations

import os
from ui.theme import inject_css_file, is_dark

from ui.dashboard.common import hex_to_rgba

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
    from engines.common.run_logs import list_run_logs

    try:
        from plotly.subplots import make_subplots
        import plotly.express as px
        import plotly.graph_objects as go
        PLOTLY_OK = True
    except Exception:
        PLOTLY_OK = False

    # ── Palette & theme ───────────────────────────────────────────────────────
    _is_dark = is_dark()

    CLR = {
        # Brand accents (Visteon) — always the same regardless of theme
        "sap":        "#FF7200",
        "plm01":      "#6EC6FF",
        "plm02":      "#00A6A6",
        "tln01":      "#2176AE",
        "tln02":      "#1A936F",
        "cleanups":   "#4D4D4D",

        # Status
        "green":      "#1A936F",
        "amber":      "#F4A011",
        "red":        "#C00000",

        # UI tokens — switch with theme
        "bg_page":    "#0F141A"  if _is_dark else "#ffffff",
        "bg_panel":   "#141B22"  if _is_dark else "#f5f7fa",
        "bg_card":    "#161F28"  if _is_dark else "#ffffff",
        "bg_alt":     "#1A2330"  if _is_dark else "#f0f2f5",
        "border":     "#253041"  if _is_dark else "#e5e7eb",
        "grid":       "#2A384D"  if _is_dark else "#e5e7eb",
        "text_main":  "#E6EDF3"  if _is_dark else "#1f2937",
        "text_muted": "#9FB0C0"  if _is_dark else "#6b7280",
        "text_invert":"#0E1116",
        "badge_bg":   "#0E1A26"  if _is_dark else "#f3f4f6",
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
    # NOTE: Background/color of html/body is intentionally NOT set here.
    # Streamlit controls that via html[data-theme] in styles.css.
    # These variables bridge the dashboard components to the active theme.
    inject_css_file("dashboard.css")


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

    # ── Assignment overdue badge for banner ───────────────────────────────────
    _overdue_badge = ""
    try:
        import json as _json
        _open_path = os.path.join(".history", "assignments", "open_items.json")
        if os.path.exists(_open_path):
            with open(_open_path, encoding="utf-8") as _f:
                _items = _json.load(_f)
            _today = pd.Timestamp.today().normalize()
            _overdue_n = sum(
                1 for item in _items.values()
                if item.get("due_date") and _today > pd.to_datetime(item["due_date"])
            )
            if _overdue_n > 0:
                _overdue_badge = (
                    f'&nbsp;<span class="badge badge-red">'
                    f'🔴 {_overdue_n} assignment{"s" if _overdue_n > 1 else ""} overdue'
                    f'</span>'
                )
    except Exception:
        pass

    # ── A) Banner ─────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="exec-banner">
        <h1>📈 Validation Hub — Evolution Dashboard
            <span class="badge {overall[1]}">{overall[0]}</span>{_overdue_badge}
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
        _atl_scope = atl_last.get("plant_scope", atl_last.get("plant_filter","ALL")) if atl_last else "ALL"
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
                                   f"{tln:,} with TLN &nbsp;·&nbsp; {user} &nbsp;·&nbsp; {_atl_scope} &nbsp;·&nbsp; {atl_age}",
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
    vlbl_map  = {"sap":"errors","atl":"CNs","dataclean":"outdated"}
    acc_map   = {"sap":CLR["sap"],"atl":CLR["plm01"],"dataclean":CLR["cleanups"]}

    feed_html = ''
    for kind, logs in [("sap",sap_logs),("atl",atl_logs),("dataclean",dc_logs)]:
        for log in sorted(logs, key=lambda l: l.get("timestamp",""), reverse=True)[:3]:
            ts  = log.get("timestamp","")
            m_  = log.get("metrics",{})
            if kind == "atl":
                val = m_.get("PT01_rows", 0) + m_.get("PT02_rows", 0)
            else:
                val = m_.get(vkey_map[kind], 0)
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
            # ATL logs use 'plant_scope', not 'plant_filter'
            _scope = log.get("plant_scope", log.get("plant_filter", "ALL"))
            entries.append({"ts":ts,"plant":_scope if _scope in ("PT01","PT02") else "PT01","source":"PLM Workflow PT01","value":m.get("PT01_rows",0)})
            entries.append({"ts":ts,"plant":_scope if _scope in ("PT01","PT02") else "PT02","source":"PLM Workflow PT02","value":m.get("PT02_rows",0)})
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

    # ── I) Assignment Health ──────────────────────────────────────────────────
    _render_assignment_health(CLR, PLOTLY_OK)


def _render_assignment_health(CLR: dict, PLOTLY_OK: bool) -> None:
    """
    Renders the 📋 Assignment Health section inside the Evolution Dashboard.
    Reads from .history/assignments/ — no additional inputs required.
    """
    import os, json
    import pandas as pd
    import streamlit as st

    if PLOTLY_OK:
        try:
            import plotly.graph_objects as go
        except Exception:
            PLOTLY_OK = False

    open_path   = os.path.join(".history", "assignments", "open_items.json")
    master_path = os.path.join(".history", "assignments", "master.json")

    if not os.path.exists(open_path) and not os.path.exists(master_path):
        return  # No assignment data yet — silently skip

    st.markdown('<div class="section-header">📋 Assignment Health</div>',
                unsafe_allow_html=True)

    # ── Load data ─────────────────────────────────────────────────────────────
    open_items: dict = {}
    if os.path.exists(open_path):
        try:
            with open(open_path, encoding="utf-8") as f:
                open_items = json.load(f)
        except Exception:
            open_items = {}

    master: dict = {"field_assignments": {}, "plant_overrides": {}}
    if os.path.exists(master_path):
        try:
            with open(master_path, encoding="utf-8") as f:
                master = json.load(f)
        except Exception:
            pass

    if not open_items and not master.get("field_assignments"):
        st.markdown(
            '<div class="insight-box insight-neutral">'
            '📋 No assignments yet. Run a SAP validation and assign fields in '
            '<strong>Step 4 → Assign &amp; Track</strong> to start tracking here.'
            '</div>',
            unsafe_allow_html=True
        )
        return

    # ── Build DataFrame ───────────────────────────────────────────────────────
    today = pd.Timestamp.today().normalize()
    rows  = []

    for key, item in open_items.items():
        parts    = key.split("||")
        material = parts[0] if len(parts) > 0 else ""
        plant    = parts[1] if len(parts) > 1 else ""
        field    = parts[2] if len(parts) > 2 else ""
        owner    = item.get("owner", "")
        status   = item.get("status", "Open")
        days_open = int(item.get("days_open", 0))
        due_str  = item.get("due_date", "")

        if due_str:
            try:
                due_dt = pd.to_datetime(due_str)
                status = "Overdue" if today > due_dt else "Open"
            except Exception:
                pass

        rows.append({
            "Field":     field,
            "Owner":     owner if owner else "Unassigned",
            "Team":      item.get("team", ""),
            "Status":    status,
            "Days_Open": days_open,
            "Due_Date":  due_str,
        })

    # Add assigned fields from master that have no open items (clean)
    known_fields = {r["Field"] for r in rows}
    for field, info in master.get("field_assignments", {}).items():
        if field not in known_fields:
            rows.append({
                "Field": field,
                "Owner": info.get("owner", "") or "Unassigned",
                "Team":  info.get("team", ""),
                "Status": "Clean", "Days_Open": 0, "Due_Date": "",
            })

    if not rows:
        st.caption("No assignment data to display.")
        return

    df = pd.DataFrame(rows)

    # ── KPI Cards ─────────────────────────────────────────────────────────────
    total      = len(df)
    assigned   = int((df["Owner"] != "Unassigned").sum())
    unassigned = total - assigned
    overdue    = int((df["Status"] == "Overdue").sum())
    avg_days   = round(df[df["Days_Open"] > 0]["Days_Open"].mean(), 1) \
                 if (df["Days_Open"] > 0).any() else 0

    def _kpi(label, value, sub="", accent="#FF7200"):
        return (f'<div class="kpi-card" style="border-left:3px solid {accent}">'
                f'<div class="kpi-label">{label}</div>'
                f'<div class="kpi-value">{value}</div>'
                f'<div class="kpi-sub">{sub}</div></div>')

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(_kpi("Total Items", f"{total:,}", "in tracking",
                         CLR["sap"]), unsafe_allow_html=True)
    with k2:
        pct_str = f"{assigned/total*100:.0f}% of total" if total else "—"
        st.markdown(_kpi("Assigned", f"{assigned:,}", pct_str,
                         CLR["plm02"]), unsafe_allow_html=True)
    with k3:
        st.markdown(_kpi("Unassigned", f"{unassigned:,}", "need owner",
                         CLR["cleanups"]), unsafe_allow_html=True)
    with k4:
        acc = CLR["red"] if overdue > 0 else CLR["green"]
        st.markdown(_kpi("Overdue", f"{overdue:,}", "past due date",
                         acc), unsafe_allow_html=True)
    with k5:
        st.markdown(_kpi("Avg Days Open", f"{avg_days}", "assigned items",
                         CLR["plm01"]), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Charts ────────────────────────────────────────────────────────────────
    ch1, ch2 = st.columns(2)

    STATUS_COLORS = {
        "Overdue": CLR["red"],
        "Open":    CLR["amber"],
        "New":     CLR["plm01"],
        "Clean":   CLR["green"],
    }

    # Chart A: Errors by Owner stacked by Status
    with ch1:
        _owner_df = (
            df[df["Owner"] != "Unassigned"]
            .groupby(["Owner", "Status"])
            .size()
            .reset_index(name="Count")
        )
        if not _owner_df.empty and PLOTLY_OK:
            fig = go.Figure()
            for status in ["Overdue", "Open", "New", "Clean"]:
                sub = _owner_df[_owner_df["Status"] == status]
                if sub.empty:
                    continue
                fig.add_trace(go.Bar(
                    name=status, x=sub["Owner"], y=sub["Count"],
                    marker_color=STATUS_COLORS.get(status, "#888"),
                ))
            fig.update_layout(
                barmode="stack",
                title=dict(text="Errors by Owner",
                           font=dict(size=13, color=CLR["text_main"]),
                           x=0.02, xanchor="left"),
                height=320,
                margin=dict(l=20, r=20, t=45, b=60),
                paper_bgcolor=CLR["bg_panel"],
                plot_bgcolor=CLR["bg_panel"],
                font=dict(size=11, color=CLR["text_main"]),
                legend=dict(orientation="h", y=-0.28, x=0.5, xanchor="center",
                            bgcolor=CLR["bg_panel"],
                            font=dict(color=CLR["text_main"])),
            )
            fig.update_xaxes(tickangle=-30, color=CLR["text_muted"], showgrid=False)
            fig.update_yaxes(color=CLR["text_muted"], gridcolor=CLR["grid"])
            st.plotly_chart(fig, width='stretch',
                            config={"displayModeBar": False})
        elif not _owner_df.empty:
            st.dataframe(_owner_df, width='stretch')

    # Chart B: Overdue by Field
    with ch2:
        _overdue_df = (
            df[df["Status"] == "Overdue"]
            .groupby("Field").size()
            .reset_index(name="Overdue")
            .sort_values("Overdue", ascending=False)
            .head(12)
        )
        if not _overdue_df.empty and PLOTLY_OK:
            fig2 = go.Figure(go.Bar(
                x=_overdue_df["Field"], y=_overdue_df["Overdue"],
                marker_color=CLR["red"],
                text=_overdue_df["Overdue"], textposition="outside",
            ))
            fig2.update_layout(
                title=dict(text="Overdue Items by Field",
                           font=dict(size=13, color=CLR["text_main"]),
                           x=0.02, xanchor="left"),
                height=320,
                margin=dict(l=20, r=20, t=45, b=80),
                paper_bgcolor=CLR["bg_panel"],
                plot_bgcolor=CLR["bg_panel"],
                font=dict(size=11, color=CLR["text_main"]),
                showlegend=False,
            )
            fig2.update_xaxes(tickangle=-35, color=CLR["text_muted"], showgrid=False)
            fig2.update_yaxes(color=CLR["text_muted"], gridcolor=CLR["grid"])
            st.plotly_chart(fig2, width='stretch',
                            config={"displayModeBar": False})
        else:
            st.markdown(
                '<div class="insight-box insight-good" '
                'style="padding:30px 20px;text-align:center">'
                '✅ No overdue items — all assignments within their deadlines.</div>',
                unsafe_allow_html=True
            )

    # ── Owner detail table ────────────────────────────────────────────────────
    _owner_detail = (
        df[df["Owner"] != "Unassigned"]
        .groupby(["Owner", "Team"])
        .agg(
            Fields       =("Field",     lambda x: len(x.unique())),
            Total_Items  =("Field",     "count"),
            Overdue      =("Status",    lambda x: (x == "Overdue").sum()),
            Avg_Days_Open=("Days_Open", lambda x: round(x[x > 0].mean(), 1)
                                        if (x > 0).any() else 0),
        )
        .reset_index()
        .sort_values(["Overdue", "Total_Items"], ascending=[False, False])
    )

    if not _owner_detail.empty:
        st.markdown('<div class="section-header" style="font-size:0.7rem">Owner Detail</div>',
                    unsafe_allow_html=True)
        try:
            def _style_row(row):
                if row["Overdue"] > 0:
                    return ["background-color: rgba(192,0,0,0.12)"] * len(row)
                return [""] * len(row)
            st.dataframe(
                _owner_detail.style.apply(_style_row, axis=1),
                width='stretch', hide_index=True,
            )
        except Exception:
            st.dataframe(_owner_detail, width='stretch', hide_index=True)

    # ── Unassigned warning ────────────────────────────────────────────────────
    _unassigned = sorted(df[df["Owner"] == "Unassigned"]["Field"].unique().tolist())
    if _unassigned:
        _ulist = ", ".join(f"<strong>{f}</strong>" for f in _unassigned[:8])
        _more  = f" and {len(_unassigned)-8} more" if len(_unassigned) > 8 else ""
        st.markdown(
            f'<div class="insight-box insight-warn">⚠️ <strong>{len(_unassigned)} fields</strong> '
            f'have no owner: {_ulist}{_more}. '
            f'Go to <strong>SAP Validation → Step 4 → Assign &amp; Track</strong>.</div>',
            unsafe_allow_html=True
        )
