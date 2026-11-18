"""Componentes de UI reutilizáveis: wizard, estado dos ficheiros, histórico, CSS."""

import os
import streamlit as st
from services.uploads import LABELS_PT
from pathlib import Path
from engines.common.run_logs import list_run_logs

def inject_css(css_path: str):
    try:
        css_file = Path(__file__).resolve().parents[1] / css_path
        if css_file.exists():
            st.markdown(f"<style>{css_file.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
        else:
            st.warning(f"⚠️ CSS não encontrado: {css_file}")
    except Exception as e:
        st.error(f"Erro ao carregar CSS: {e}")


def render_wizard(current_step: int, steps: list[str]):
    """Wizard visual reutilizável. Recebe o passo atual e a lista de labels."""
    cols = st.columns(len(steps))
    for idx, (col, label) in enumerate(zip(cols, steps), start=1):
        with col:
            if current_step == idx:
                klass = "active"
            elif current_step > idx:
                klass = "done"
            else:
                klass = "pending"
            st.markdown(
                f"""
                <div class="wizard-step">
                    <div class="step-circle {klass}">{idx}</div>
                    <div class="step-label">{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    completion = (current_step - 1) / max(len(steps) - 1, 1)
    st.markdown(
        f'<div class="progress-line" style="width:{completion * 100:.0f}%"></div>',
        unsafe_allow_html=True,
    )


WIZARD_STEPS_DEFAULT  = ["Files", "Scope", "Validate", "Results"]


WIZARD_STEPS_PROGRESS = ["Mode", "Files", "Validate", "Results"]


def show_file_status(required_keys: list[str], files: dict, persisted: dict,
                     optional_keys: list[str] | None = None):
    """
    Mostra um único expander com o status de cada ficheiro obrigatório:
    ✅ identificado no upload atual | 💾 guardado na sessão | ❌ em falta

    Se optional_keys for fornecido, mostra também uma secção de opcionais
    com ⬜ (não carregado) ou ✅ (presente).

    Também mostra ficheiros não reconhecidos com sugestão de renomeação.
    """
    with st.expander("📋 File Status", expanded=True):
        for key in required_keys:
            label = LABELS_PT.get(key, key)
            in_upload    = files.get(key) is not None
            in_persisted = key in persisted

            if in_upload:
                icon = "✅"
                detail = getattr(files[key], "name", "")
            elif in_persisted:
                icon = "💾"
                detail = persisted[key].get("name", "")
            else:
                icon = "❌"
                detail = "Missing"

            st.markdown(f"{icon} **{label}** — `{detail}`")

        if optional_keys:
            st.markdown("---")
            st.caption("🔓 Optional — upload to unlock extra validation rules")
            for key in optional_keys:
                label = LABELS_PT.get(key, key)
                in_upload    = files.get(key) is not None
                in_persisted = key in persisted

                if in_upload:
                    icon   = "✅"
                    detail = getattr(files[key], "name", "")
                elif in_persisted:
                    icon   = "✅"
                    detail = persisted[key].get("name", "")
                else:
                    icon   = "⬜"
                    detail = "Not uploaded (optional)"

                st.markdown(f"{icon} **{label}** — `{detail}`")

        # ── Unrecognised files ────────────────────────────────────────────────
        known_keys  = set(required_keys) | set(optional_keys or [])
        all_matched = {k for k, v in files.items() if v is not None}
        # Files that were uploaded but didn't match any known slot
        # We need to check the raw upload list — files dict has None for unmatched
        # So: any uploaded file whose name doesn't map to a recognised key
        unrecognised = []
        for key, fobj in files.items():
            if fobj is not None and key not in known_keys:
                fname = getattr(fobj, "name", "")
                if fname:
                    unrecognised.append(fname)

        if unrecognised:
            st.markdown("---")
            st.caption("⚠️ Unrecognised files — could not be automatically matched:")
            # Build suggestion map: keyword in filename → what it might be
            _SUGGESTIONS = {
                "mara":       "MARA.xlsx",
                "marc":       "MARC.xlsx",
                "mvke":       "MVKE.xlsx",
                "mbew":       "MBEW.xlsx",
                "ekpo":       "EKPO.xlsx",
                "mard":       "MARD.xlsx",
                "marm":       "MARM.xlsx",
                "zplm":       "ZPLM MARC.xlsx or ZPLM MARA.xlsx",
                "automotive": "PLM Automotive Validation.xlsx",
                "timing":     "Timing Letters.xlsx",
            }
            for fname in unrecognised:
                name_lower = fname.lower()
                suggestion = next(
                    (label for kw, label in _SUGGESTIONS.items() if kw in name_lower),
                    None
                )
                hint = f" → might be **{suggestion}**?" if suggestion else                        " → rename to match expected filename (e.g. MARA.xlsx, MARC.xlsx)"
                st.markdown(f"❓ `{fname}`{hint}")


def identity_input(label: str, key: str) -> str:
    """
    FIX 9: All pages share a single CDS ID stored in session_state["executor_name"].
    The per-page key is kept only for widget uniqueness but the value is always
    read from and written to the shared session key — so navigating between pages
    never clears the CDS ID the user already typed.
    """
    executor_name = st.text_input(
        label,
        value=st.session_state.get("executor_name", ""),
        placeholder="CDS ID...",
        key=key,
    )
    if executor_name.strip():
        st.session_state["executor_name"] = executor_name.strip()
    if not st.session_state.get("executor_name", ""):
        st.warning("⚠ Please enter your CDS ID to continue.")
        st.stop()
    return st.session_state["executor_name"]


def render_history_with_download(run_type: str, limit: int = 10,
                                  col_map: dict | None = None,
                                  label: str = "History") -> None:
    """
    Renders a history expander that shows the last N runs.
    For runs that have a saved report on disk, shows an inline download button.

    col_map: dict of {display_label: log_key_path}
      e.g. {"When": "timestamp", "Who": "user", "Errors": "metrics.total_errors"}
    """
    def _get(log, dotted_key):
        parts = dotted_key.split(".")
        v = log
        for p in parts:
            v = v.get(p, {}) if isinstance(v, dict) else {}
        return v if v != {} else "—"

    with st.expander(f"🕘 {label} (last {limit} runs)", expanded=False):
        logs = list_run_logs(run_type, limit=limit)
        if not logs:
            st.caption("No history yet.")
            return

        if col_map is None:
            col_map = {"When": "timestamp", "Who": "user"}

        for log in logs:
            row_cols = list(col_map.keys())
            # Add a Download column if a report_path exists on disk
            has_file = bool(log.get("report_path") and os.path.exists(log["report_path"]))
            n_cols = len(row_cols) + (1 if has_file else 0)
            cols = st.columns([3] * len(row_cols) + ([2] if has_file else []))

            for i, (display, key) in enumerate(col_map.items()):
                with cols[i]:
                    val = _get(log, key)
                    st.caption(display)
                    st.markdown(f"**{val}**")

            if has_file:
                with cols[-1]:
                    st.caption("Report")
                    try:
                        with open(log["report_path"], "rb") as _f:
                            _data = _f.read()
                        st.download_button(
                            "⬇️",
                            data=_data,
                            file_name=os.path.basename(log["report_path"]),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_{run_type}_{log.get('timestamp','').replace(':','-')}",
                            help=f"Download: {os.path.basename(log['report_path'])}",
                        )
                    except Exception:
                        st.caption("—")
            st.divider()
