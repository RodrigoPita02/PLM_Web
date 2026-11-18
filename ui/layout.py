"""Layout global: header, footer e secções."""
from __future__ import annotations

import time
from datetime import datetime

import streamlit as st

from ui.theme import DEFAULT_THEME, THEMES, inject_theme_overrides


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
    default_theme = DEFAULT_THEME

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