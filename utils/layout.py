
# utils/layout.py
from __future__ import annotations
import time
from datetime import datetime
import streamlit as st

# ---------- Helpers de estilo dinâmico (Theme Picker) ----------
THEMES = {
    "Corporate Orange (default)": {
        "--primary": "#FF6A13",
        "--primary-hover": "#ff7f38",
        "--primary-dark": "#d8560f",
        "--accent": "#2563eb",
    },
    "Sky Blue": {
        "--primary": "#0ea5e9",
        "--primary-hover": "#38bdf8",
        "--primary-dark": "#0284c7",
        "--accent": "#22c55e",
    },
    "High Contrast": {
        "--primary": "#111827",
        "--primary-hover": "#1f2937",
        "--primary-dark": "#000000",
        "--accent": "#f59e0b",
    },
}

def inject_theme_overrides(theme_name: str | None):
    """Injeta CSS inline a sobrepor variáveis do tema base (styles.css)."""
    name = theme_name or "Corporate Orange (default)"
    palette = THEMES.get(name, THEMES["Corporate Orange (default)"])
    css_vars = "; ".join([f"{k}: {v}" for k, v in palette.items()])
    st.markdown(
        f"""
        <style>
        :root {{
            {css_vars};
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# ---------- HEADER / FOOTER ----------
def render_header(page_title: str, *, show_search: bool = True, show_theme_picker: bool = True):
    """
    Desenha cabeçalho premium com título, pesquisa global (hook) e theme picker.
    Retorna o termo de pesquisa (str) e o nome do tema escolhido.
    """
    # container visual para header (usa classes do teu CSS + suplementares abaixo)
    with st.container():
        st.markdown('<div class="app-header-wrap">', unsafe_allow_html=True)

        c1, c2, c3 = st.columns([4, 3, 2], vertical_alignment="center")

        with c1:
            st.markdown(f'<div class="app-page-title">{page_title}</div>', unsafe_allow_html=True)

        theme_choice = "Corporate Orange (default)"
        if show_theme_picker:
            with c3:
                theme_choice = st.selectbox(
                    "Theme",
                    list(THEMES.keys()),
                    index=0,
                    key="theme_picker",
                )
                inject_theme_overrides(theme_choice)

        st.markdown('</div>', unsafe_allow_html=True)

    # separador fino
    st.markdown('<div class="app-separator"></div>', unsafe_allow_html=True)

    return theme_choice

def render_footer(start_time: float, *, version: str = "v1.0"):
    """Desenha rodapé com versão e tempo de execução da página."""
    elapsed = time.time() - start_time
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.markdown(
        f"""
        <div class="app-footer">
          <div>Validation Hub <b>{version}</b></div>
          <div>Render time: {elapsed:.2f}s · {ts}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------- Section wrappers ----------
def section(title: str, subtitle: str | None = None):
    """Abre uma secção padronizada com título/subtítulo."""
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="section-subtitle">{subtitle}</div>', unsafe_allow_html=True)
