"""Sistema de temas: paleta por tema, tokens CSS e injeção no :root.

Fonte única de verdade para cores da UI. O CSS estático (assets/styles.css) só
define fallbacks; os tokens do tema ativo são injetados aqui em cada rerun.
Os dashboards usam os tokens-alias (--bg-card, --txt-main, ...) definidos abaixo,
por isso acompanham sempre o seletor de tema (e não o tema do browser).
"""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import streamlit as st

DEFAULT_THEME = "🟠 Visteon Corporate (default)"
ASSETS = Path(__file__).resolve().parents[1] / "assets"

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
        "--bg-alt":           "#f0f2f5",
        "--grid":             "#e5e7eb",
        "--badge-bg":         "#f3f4f6",
        "--txt-inv":          "#0E1116",
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
        "--bg-alt":           "#1A2330",
        "--grid":             "#2A384D",
        "--badge-bg":         "#0E1A26",
        "--txt-inv":          "#0E1116",
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
        "--bg-alt":           "#f1f5f9",
        "--grid":             "#e2e8f0",
        "--badge-bg":         "#f1f5f9",
        "--txt-inv":          "#0E1116",
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
        "--bg-alt":           "#1f1f1f",
        "--grid":             "#333333",
        "--badge-bg":         "#1a1a1a",
        "--txt-inv":          "#000000",
    },
}

DARK_THEMES = {"🌙 Dark Mode", "🔲 High Contrast"}

# Aliases semânticos usados por dashboards/componentes (valor = outro token).
ALIASES = {
    "--bg-page":   "var(--app-bg)",
    "--bg-panel":  "var(--app-bg-secondary)",
    "--bg-card":   "var(--card-bg-2)",
    "--border":    "var(--app-border)",
    "--txt-main":  "var(--app-fg)",
    "--txt-muted": "var(--app-fg-muted)",
}


def active_theme() -> str:
    return st.session_state.get("theme_picker", DEFAULT_THEME)


def is_dark(theme_name: str | None = None) -> bool:
    return (theme_name or active_theme()) in DARK_THEMES


@lru_cache(maxsize=None)
def read_asset(name: str) -> str:
    return (ASSETS / name).read_text(encoding="utf-8")


def inject_css_file(name: str) -> None:
    """Injeta um ficheiro de assets/ como <style>."""
    st.markdown(f"<style>{read_asset(name)}</style>", unsafe_allow_html=True)


def inject_theme_overrides(theme_name: str | None) -> None:
    """Injeta os tokens do tema escolhido em :root, mais os overrides de elementos.

    Chamada UMA vez por rerun, antes dos widgets. Os overrides são sempre emitidos
    (mesmo em temas claros) para limpar valores deixados por um tema escuro anterior.
    """
    palette = THEMES.get(theme_name or DEFAULT_THEME, THEMES[DEFAULT_THEME])
    tokens = {**palette, **ALIASES}
    css_vars = "\n    ".join(f"{k}: {v};" for k, v in tokens.items())
    st.markdown(
        f"<style>:root {{\n    {css_vars}\n}}\n{read_asset('theme_overrides.css')}</style>",
        unsafe_allow_html=True,
    )
