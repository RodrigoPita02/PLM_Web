"""Vista: 🚗 PLM Workflow' and sub_choice."""

import streamlit as st
from ui.components import WIZARD_STEPS_DEFAULT, identity_input, render_wizard
from ui.layout import render_footer


def render(start_time: float) -> None:

    st.subheader("🛠️ Change Cockpit (In Progress...)")
    effective_executor = identity_input("Who performs?", "cc_exec")
    save_history = st.checkbox("Save this execution in history", value=True)

    if "cc_step" not in st.session_state:
        st.session_state.cc_step = 1

    render_wizard(st.session_state.cc_step, WIZARD_STEPS_DEFAULT)

    if st.session_state.cc_step == 1:
        st.markdown("### 📂 Step 1 — Load Files")
        uploads = st.file_uploader("Drop Change Cockpit input(s) here",
                                   type=["xlsx", "csv"], accept_multiple_files=True, key="cc_multi_uploader")
        # TODO: define required CC files and classify
        c1, c2 = st.columns(2)
        with c2:
            if st.button("➡️ Continue to Scope", key="cc_next1", width='stretch'):
                st.session_state.cc_step = 2
                st.rerun()

    elif st.session_state.cc_step == 2:
        st.markdown("### 🌍 Step 2 — Select Scope")
        cc_scope = st.radio("Scope:", ["All", "Only PT01", "Only PT02"], horizontal=True)
        st.session_state.cc_scope = cc_scope
        c1, c2 = st.columns(2)
        with c1:
            if st.button("⬅️ Back", key="cc_back2", width='stretch'):
                st.session_state.cc_step = 1; st.rerun()
        with c2:
            if st.button("➡️ Continue to Validation", key="cc_next2", width='stretch'):
                st.session_state.cc_step = 3; st.rerun()

    elif st.session_state.cc_step == 3:
        st.markdown("### ⚙️ Step 3 — Run Change Cockpit")
        st.info("Define here the processing logic for Change Cockpit.")
        if st.button("▶️ Run Change Cockpit", type="primary", key="cc_run"):
            with st.spinner("Running Change Cockpit..."):
                st.session_state.cc_result_df = None
                st.session_state.cc_excel     = None
            st.session_state.cc_step = 4; st.rerun()

    elif st.session_state.cc_step == 4:
        st.markdown("### 📈 Step 4 — Results")
        st.success("Results placeholder — plug your tables and dashboards here.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 New Run", key="cc_new", width='stretch'):
                st.session_state.cc_step = 1; st.rerun()
        with c2:
            if st.button("🧹 Clear & Start Over", key="cc_clear4", width='stretch'):
                for k in ["cc_result_df", "cc_excel"]:
                    st.session_state.pop(k, None)
                st.session_state.cc_step = 1; st.rerun()

    render_footer(start_time, version="v1.4")
