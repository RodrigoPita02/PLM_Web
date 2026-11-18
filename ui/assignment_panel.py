"""Painel de atribuição de owners (passo de resultados do SAP)."""

import io
import pandas as pd
import streamlit as st
from engines.common.rolesviews import get_rolesviews_bytesio as _safe_rolesviews_bytesio
from engines.assignments.store import (
    build_assignments_sheet,
    get_owner,
    list_all_assignments,
    merge_assignments,
    set_owner,
)
from engines.assignments.history import track_error_count

def render_assignment_panel(errors_df: pd.DataFrame) -> None:
    """
    Renders the 📋 Assign & Track panel in SAP Step 4.
    """

    if errors_df is None or errors_df.empty:
        return

    st.markdown("## 📋 Assign & Track")
    st.markdown(
        """<div style="background:var(--app-bg-secondary);border-left:4px solid var(--primary);
        border-radius:6px;padding:10px 16px;margin-bottom:16px;font-size:0.85rem;color:var(--app-fg)">
        Assign each error field to a responsible person. Assignments are <strong>saved permanently</strong>
        — on the next validation run, owners are pre-filled automatically. Edit any time with the ✏️ button.
        </div>""",
        unsafe_allow_html=True,
    )

    # ─────────────────────────────────────────────
    # FILTER FIELD ERRORS
    # ─────────────────────────────────────────────
    _fe = errors_df[
        errors_df.get("Error_Type", pd.Series("Field Error", index=errors_df.index)) == "Field Error"
    ] if "Error_Type" in errors_df.columns else errors_df

    if _fe.empty:
        st.info("No field errors to assign.")
        return

    field_counts = (
        _fe.groupby("Field")
           .agg(
                Errors=("Field", "count"),
                Plants=("Plant", lambda x: ", ".join(sorted(x.astype(str).str.strip().unique())))
           )
           .reset_index()
           .sort_values("Errors", ascending=False)
    )

    if field_counts.empty:
        return

    # ─────────────────────────────────────────────
    # HEADER
    # ─────────────────────────────────────────────
    hcols = st.columns([4, 2, 3, 2, 1])
    for col, label in zip(hcols, ["Field", "Errors", "Owner", "Team", "Action"]):
        col.markdown(
            f"<div style='font-size:0.75rem;font-weight:700;text-transform:uppercase;"
            f"letter-spacing:.06em;color:var(--app-fg-muted)'>{label}</div>",
            unsafe_allow_html=True,
        )

    st.markdown("<hr style='margin:4px 0 8px 0;border-color:var(--app-border)'>",
                unsafe_allow_html=True)

    # ─────────────────────────────────────────────
    # LOAD MASTER + ROLES (✅ PATCH AQUI)
    # ─────────────────────────────────────────────
    from engines.assignments.store import _load_master, _build_field_to_role_map
    from engines.common.fast_loader import load_xlsx_fast

    _master_cache = _load_master()

    try:
        _rv = load_xlsx_fast(_safe_rolesviews_bytesio(), sheet_name=None)
        _roles_df = _rv.get("Roles")
        _role_map = _build_field_to_role_map(_roles_df)
    except Exception:
        _role_map = {}

    # ─────────────────────────────────────────────
    # LOOP CAMPOS
    # ─────────────────────────────────────────────
    for _, row in field_counts.iterrows():
        field  = row["Field"]
        n_errs = row["Errors"]
        plants = row.get("Plants", "")

        current_owner = get_owner(field, master=_master_cache)

        # ✅ TEAM LOGIC (FINAL)
        _fd = _master_cache.get("field_assignments", {}).get(field, {})
        assigned_team = (_fd.get("team", "") if isinstance(_fd, dict) else "") or ""
        assigned_team = assigned_team.strip()

        role_team = (_role_map.get(field, "") or "").strip()

        if assigned_team:
            current_team = assigned_team
        elif role_team:
            current_team = role_team
        else:
            current_team = ""

        edit_key = f"asgn_edit_{field}"
        if edit_key not in st.session_state:
            st.session_state[edit_key] = False

        cols = st.columns([4, 2, 3, 2, 1])

        with cols[0]:
            st.markdown(f"**{field}**<br>"
                        f"<span style='font-size:0.75rem;color:var(--app-fg-muted)'>{plants}</span>",
                        unsafe_allow_html=True)

        with cols[1]:
            color = "var(--color-error)" if n_errs > 50 else \
                    "var(--color-warn)" if n_errs > 10 else \
                    "var(--app-fg-muted)"
            st.markdown(f"<span style='font-weight:700;color:{color}'>{n_errs:,}</span>",
                        unsafe_allow_html=True)

        with cols[2]:
            if st.session_state[edit_key]:
                st.text_input(
                    "Owner",
                    value=current_owner,
                    key=f"inp_owner_{field}",
                    label_visibility="collapsed"
                )
            else:
                st.markdown(f"👤 **{current_owner}**" if current_owner else "*Unassigned*")

        with cols[3]:
            if st.session_state[edit_key]:
                st.text_input(
                    "Team",
                    value=current_team,
                    key=f"inp_team_{field}",
                    label_visibility="collapsed"
                )
            else:
                st.markdown(current_team if current_team else "-")

        with cols[4]:
            if st.session_state[edit_key]:
                if st.button("💾", key=f"save_{field}"):
                    set_owner(
                        field,
                        st.session_state.get(f"inp_owner_{field}", ""),
                        team=st.session_state.get(f"inp_team_{field}", "")
                    )
                    st.session_state[edit_key] = False
                    st.rerun()
            else:
                if st.button("✏️", key=f"edit_{field}"):
                    st.session_state[edit_key] = True
                    st.rerun()

        st.markdown("<hr style='margin:4px 0;border-color:var(--app-border)'>",
                    unsafe_allow_html=True)

    # ─────────────────────────────────────────────
    # EXPORT (✅ PATCH AQUI)
    # ─────────────────────────────────────────────
    with st.expander("📜 All current assignments", expanded=False):

        _all_asgn = list_all_assignments()
        st.dataframe(_all_asgn, width='stretch')

        try:
            from engines.common.fast_loader import load_xlsx_fast

            _rv = load_xlsx_fast(_safe_rolesviews_bytesio(), sheet_name=None)
            _roles_df = _rv.get("Roles")

            # ✅ IMPORTANTE: usar roles_df
            _enriched_exp = merge_assignments(errors_df, roles_df=_roles_df)

            if not _enriched_exp.empty:
                track_error_count(_enriched_exp)

            _asgn_sheet = build_assignments_sheet(_enriched_exp)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as ew:
                _asgn_sheet.to_excel(ew, sheet_name="Assignments", index=False)

            buf.seek(0)

            st.download_button(
                "⬇️ Export Assignments",
                data=buf.getvalue(),
                file_name="assignments.xlsx"
            )

        except Exception:
            st.caption("Export unavailable")
