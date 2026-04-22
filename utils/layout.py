# utils/layout.py
from __future__ import annotations
import time
from datetime import datetime
import streamlit as st

# =============================================================================
# THEME DEFINITIONS
# Each theme is a complete CSS variable block injected into :root on every rerun.
# This is the only reliable way to change themes in Streamlit at runtime —
# we cannot set html[data-theme] from Python, so we own all the variables.
# =============================================================================

THEMES: dict[str, dict] = {
    "🟠 Visteon Corporate (default)": {
        "--primary":          "#FF6A13",
        "--primary-hover":    "#ff7f38",
        "--primary-dark":     "#d8560f",
        "--accent":           "#2563eb",
        "--glow":             "rgba(255,106,19,.35)",
        "--app-bg":           "#ffffff",
        "--app-bg-secondary": "#f5f7fa",
        "--app-fg":           "#1f2937",
        "--app-fg-muted":     "#6b7280",
        "--app-border":       "rgba(0,0,0,.12)",
        "--menu-bg":          "#f3f4f6",
        "--menu-item-bg":     "#ffffff",
        "--menu-item-hover":  "#ececec",
        "--menu-item-active": "rgba(255,106,19,.18)",
        "--menu-border":      "#e5e7eb",
        "--menu-text":        "#1f2937",
        "--menu-text-muted":  "#6b7280",
        "--menu-radio":       "#FF6A13",
        "--card-bg-1":        "#fafafa",
        "--card-bg-2":        "#ffffff",
    },
    "🌙 Dark Mode": {
        "--primary":          "#FF6A13",
        "--primary-hover":    "#ff7f38",
        "--primary-dark":     "#d8560f",
        "--accent":           "#38bdf8",
        "--glow":             "rgba(255,106,19,.45)",
        "--app-bg":           "#0b0b0b",
        "--app-bg-secondary": "#131313",
        "--app-fg":           "#f3f4f6",
        "--app-fg-muted":     "#9ca3af",
        "--app-border":       "rgba(255,255,255,.12)",
        "--menu-bg":          "#0b0b0b",
        "--menu-item-bg":     "#1a1a1a",
        "--menu-item-hover":  "#252525",
        "--menu-item-active": "rgba(255,106,19,.25)",
        "--menu-border":      "rgba(255,255,255,.12)",
        "--menu-text":        "#e5e5e5",
        "--menu-text-muted":  "#9ca3af",
        "--menu-radio":       "#FF6A13",
        "--card-bg-1":        "#161616",
        "--card-bg-2":        "#1f1f1f",
    },
    "☀️ Light Clean": {
        "--primary":          "#2563eb",
        "--primary-hover":    "#3b82f6",
        "--primary-dark":     "#1d4ed8",
        "--accent":           "#059669",
        "--glow":             "rgba(37,99,235,.30)",
        "--app-bg":           "#ffffff",
        "--app-bg-secondary": "#f8fafc",
        "--app-fg":           "#0f172a",
        "--app-fg-muted":     "#64748b",
        "--app-border":       "rgba(0,0,0,.10)",
        "--menu-bg":          "#f1f5f9",
        "--menu-item-bg":     "#ffffff",
        "--menu-item-hover":  "#e2e8f0",
        "--menu-item-active": "rgba(37,99,235,.15)",
        "--menu-border":      "#e2e8f0",
        "--menu-text":        "#0f172a",
        "--menu-text-muted":  "#64748b",
        "--menu-radio":       "#2563eb",
        "--card-bg-1":        "#f8fafc",
        "--card-bg-2":        "#ffffff",
    },
    "🔲 High Contrast": {
        "--primary":          "#facc15",
        "--primary-hover":    "#fde047",
        "--primary-dark":     "#eab308",
        "--accent":           "#f97316",
        "--glow":             "rgba(250,204,21,.40)",
        "--app-bg":           "#000000",
        "--app-bg-secondary": "#111111",
        "--app-fg":           "#ffffff",
        "--app-fg-muted":     "#d1d5db",
        "--app-border":       "rgba(255,255,255,.20)",
        "--menu-bg":          "#000000",
        "--menu-item-bg":     "#111111",
        "--menu-item-hover":  "#1f1f1f",
        "--menu-item-active": "rgba(250,204,21,.20)",
        "--menu-border":      "rgba(255,255,255,.25)",
        "--menu-text":        "#ffffff",
        "--menu-text-muted":  "#d1d5db",
        "--menu-radio":       "#facc15",
        "--card-bg-1":        "#111111",
        "--card-bg-2":        "#1a1a1a",
    },
}

_DARK_THEMES = {"🌙 Dark Mode", "🔲 High Contrast"}


def inject_theme_overrides(theme_name: str | None) -> None:
    """
    Injects the chosen theme CSS variables into :root on every rerun.

    Rules:
    - Called ONCE per rerun, at the top of render_header (before the selectbox).
    - ALL element overrides (background, color) are always emitted — not just
      for dark themes. This ensures that switching back to a light theme properly
      clears any dark values left over from a previous rerun.
    - dashboard.py must read the active theme from session_state["theme_picker"],
      not from st.get_option("theme.base"), which only reflects config.toml.
    """
    name    = theme_name or "🟠 Visteon Corporate (default)"
    palette = THEMES.get(name, THEMES["🟠 Visteon Corporate (default)"])

    css_vars = "\n    ".join(f"{k}: {v};" for k, v in palette.items())

    # Always inject element overrides regardless of light/dark.
    # Without this, switching from dark -> light leaves the dark !important
    # rules in the DOM because Streamlit never removes injected <style> tags.
    element_overrides = """
        .stApp, .main, section.main, .block-container,
        [data-testid="stAppViewContainer"] {
            background-color: var(--app-bg) !important;
            color: var(--app-fg) !important;
        }
        input, textarea,
        [data-baseweb="input"] input,
        [data-baseweb="textarea"] textarea {
            background-color: var(--menu-item-bg) !important;
            color: var(--app-fg) !important;
            border-color: var(--app-border) !important;
        }
        [data-testid="metric-container"] {
            background-color: var(--card-bg-1) !important;
            border: 1px solid var(--app-border) !important;
            border-radius: 10px;
            padding: 12px;
        }
        [data-testid="metric-container"] label,
        [data-testid="metric-container"] [data-testid="stMetricValue"],
        [data-testid="metric-container"] [data-testid="stMetricDelta"] {
            color: var(--app-fg) !important;
        }
        .stMarkdown, .stText, p, li, td, th, label {
            color: var(--app-fg) !important;
        }
        [data-testid="stExpander"] {
            background-color: var(--app-bg-secondary) !important;
            border-color: var(--app-border) !important;
        }
        [data-testid="stDataFrame"] th {
            background-color: var(--app-bg-secondary) !important;
            color: var(--app-fg) !important;
        }
        [data-baseweb="select"] > div {
            background-color: var(--menu-item-bg) !important;
            border-color: var(--app-border) !important;
        }
        [data-baseweb="select"] span {
            color: var(--app-fg) !important;
        }
        section[data-testid="stSidebar"] {
            background-color: var(--menu-bg) !important;
        }

        /* ---- Streamlit top toolbar ---- */
        [data-testid="stHeader"] {
            background-color: var(--app-bg) !important;
            border-bottom: 1px solid var(--app-border) !important;
        }
        [data-testid="stHeader"] button svg,
        [data-testid="stHeader"] a svg {
            fill: var(--app-fg) !important;
        }
        [data-testid="stToolbar"] {
            background-color: var(--app-bg) !important;
        }
        [data-testid="stDecoration"] {
            background: var(--primary) !important;
        }

        /* ---- BaseWeb dropdowns / selects (full coverage) ---- */
        [data-baseweb="select"] > div:first-child {
            background-color: var(--menu-item-bg) !important;
            border-color: var(--app-border) !important;
            color: var(--app-fg) !important;
        }
        [data-baseweb="select"] [data-baseweb="icon"] svg path {
            fill: var(--app-fg) !important;
        }
        /* Dropdown list popup */
        [data-baseweb="popover"] [data-baseweb="menu"],
        [data-baseweb="menu"] {
            background-color: var(--menu-item-bg) !important;
            border-color: var(--app-border) !important;
        }
        [data-baseweb="menu"] li,
        [data-baseweb="menu"] [role="option"] {
            background-color: var(--menu-item-bg) !important;
            color: var(--app-fg) !important;
        }
        [data-baseweb="menu"] li:hover,
        [data-baseweb="menu"] [role="option"]:hover,
        [data-baseweb="menu"] [aria-selected="true"] {
            background-color: var(--menu-item-hover) !important;
            color: var(--app-fg) !important;
        }
        /* Selected option tag inside multiselect */
        [data-baseweb="tag"] {
            background-color: var(--menu-item-active) !important;
            color: var(--app-fg) !important;
        }
        /* Placeholder text */
        [data-baseweb="select"] input::placeholder {
            color: var(--app-fg-muted, var(--txt-muted)) !important;
        }

        /* ---- st.multiselect / st.text_input backgrounds ---- */
        [data-baseweb="base-input"] {
            background-color: var(--menu-item-bg) !important;
            color: var(--app-fg) !important;
        }

        /* ---- st.tabs ---- */
        [data-baseweb="tab-list"] {
            background-color: var(--app-bg) !important;
        }
        [data-baseweb="tab"] {
            background-color: var(--app-bg) !important;
            color: var(--app-fg-muted, var(--txt-muted)) !important;
        }
        [aria-selected="true"][data-baseweb="tab"] {
            color: var(--primary) !important;
        }
    """

    st.markdown(
        f"<style>:root {{\n    {css_vars}\n}}\n{element_overrides}</style>",
        unsafe_allow_html=True,
    )


def render_header(page_title: str, *, show_theme_picker: bool = True) -> str:
    """
    Draws the app header with title and theme picker.
    Returns the chosen theme name.

    inject_theme_overrides is called exactly ONCE per rerun — before the
    selectbox renders. The selectbox writes to session_state["theme_picker"]
    which triggers a rerun, and the next rerun picks up the new value at the
    top. Calling inject_theme_overrides a second time after the selectbox
    causes duplicate <style> blocks that fight each other in the DOM.
    """
    default_theme = "🟠 Visteon Corporate (default)"

    # Initialise on first load
    if "theme_picker" not in st.session_state:
        st.session_state["theme_picker"] = default_theme

    # Inject ONCE, at the very top of every rerun, before any widgets.
    inject_theme_overrides(st.session_state["theme_picker"])

    with st.container():
        st.markdown('<div class="app-header-wrap">', unsafe_allow_html=True)
        cols = st.columns([4, 3, 2], vertical_alignment="center")

        with cols[0]:
            st.markdown(
                f'<div class="app-page-title">{page_title}</div>',
                unsafe_allow_html=True,
            )

        if show_theme_picker:
            with cols[2]:
                st.selectbox(
                    "Theme",
                    list(THEMES.keys()),
                    index=list(THEMES.keys()).index(
                        st.session_state.get("theme_picker", default_theme)
                    ),
                    key="theme_picker",
                    label_visibility="collapsed",
                    # Changing the selectbox triggers a rerun automatically;
                    # the new theme is injected at the top of that rerun.
                )

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="app-separator"></div>', unsafe_allow_html=True)
    return st.session_state.get("theme_picker", default_theme)


def render_footer(start_time: float, *, version: str = "v1.0") -> None:
    """Draws footer with version and page render time."""
    elapsed = time.time() - start_time
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.markdown(
        f"""
        <div class="app-footer">
          <div>Validation Hub <b>{version}</b></div>
          <div>Render time: {elapsed:.2f}s · {ts}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section(title: str, subtitle: str | None = None) -> None:
    """Opens a standardised section with title/subtitle."""
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="section-subtitle">{subtitle}</div>', unsafe_allow_html=True)