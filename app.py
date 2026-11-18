"""Validation Hub — ponto de entrada.

Estrutura:
    engines/   lógica de validação, uma pasta por app (sem Streamlit)
    services/  lógica de aplicação (ficheiros, identidade)
    ui/        componentes visuais, tema e dashboards
    views/     uma página por módulo, cada uma com render(start_time)
"""
import time
from datetime import datetime

import streamlit as st

from ui.components import inject_css
from ui.layout import render_header
from views import (
    atl, change_cockpit, data_cleanups, evolution, indirect_materials,
    progress, sap_material, ssa,
)

# ==========================================================
# CONFIGURAÇÃO INICIAL
# ==========================================================
st.set_page_config(page_title="Validation Hub", layout="wide", page_icon="🧪")
start_time = time.time()

inject_css("assets/styles.css")
theme_name = render_header("🧪 Validation Hub", show_theme_picker=True)

if "executor_name" not in st.session_state:
    st.session_state["executor_name"] = ""

# ==========================================================
# SIDEBAR MENU
# ==========================================================
st.sidebar.header("Report Selection")

main_choice = st.sidebar.radio(
    "Navigation",
    [
        "📈 Evolution Dashboard",
        "📦 SAP Material Validation",
        "📦 Indirect Materials Report",
        "🚗 PLM Workflow",
        "📊 SAP Validation Progress",
        "🧹 Data Cleanups",
        "📋 SSA Validation",
    ],
    label_visibility="collapsed",
)

sub_choice = None
if main_choice == "🚗 PLM Workflow":
    st.sidebar.markdown(
        "<div class='sidebar-submenu-title'>▼ Options</div>",
        unsafe_allow_html=True,
    )
    with st.sidebar.container():
        st.sidebar.markdown("<div class='sidebar-submenu-wrap'>", unsafe_allow_html=True)
        sub_choice = st.sidebar.radio(
            "PLM Submenu",
            ["Automotive + Timing Letter (A & TL)", "Change Cockpit"],
            key="submenu_plm",
            label_visibility="collapsed",
        )
        st.sidebar.markdown("</div>", unsafe_allow_html=True)

# ── About / Changelog ────────────────────────────────────────────────────────
st.sidebar.markdown("---")
with st.sidebar.expander("ℹ️ About Validation Hub", expanded=False):
    st.markdown("""
    **Version:** v1.6  
    **Build:** """ + datetime.now().strftime("%Y-%m-%d") + """

    **Recent changes:**
    - 🔀 Advanced Planning fields merged into SAP Material Validation report
    - 📋 Assignment system — assign fields to owners, track overdue
    - 📊 Assignment Health in Evolution Dashboard
    - ✅ File column validation before running
    - 🔍 Unrecognised file detection with suggestions
    - ⬇️ Re-download previous reports from history
    - 🌙 Full dark mode support (CSS variables)

    **Modules:** SAP Validation (Palmela) · PLM Workflow · Data Cleanups · Progress · Evolution Dashboard
    """)

# ==========================================================
# 📈 EVOLUTION DASHBOARD
# ==========================================================

# ==========================================================
# ROUTER
# ==========================================================
if main_choice == "📈 Evolution Dashboard":
    evolution.render(start_time)
elif main_choice == "📦 SAP Material Validation":
    sap_material.render(start_time)
elif main_choice == "🚗 PLM Workflow" and sub_choice == "Automotive + Timing Letter (A & TL)":
    atl.render(start_time)
elif main_choice == "🚗 PLM Workflow" and sub_choice == "Change Cockpit":
    change_cockpit.render(start_time)
elif main_choice == "📊 SAP Validation Progress":
    progress.render(start_time)
elif main_choice == "🧹 Data Cleanups":
    data_cleanups.render(start_time)
elif main_choice == "📦 Indirect Materials Report":
    indirect_materials.render(start_time)
elif main_choice == "📋 SSA Validation":
    ssa.render(start_time)
