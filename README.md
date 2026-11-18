# PLM_Web — Validation Hub

App Streamlit de validação de dados mestres SAP / PLM (Visteon).

## Estrutura

```
app.py                  entrada: tema, sidebar e router (≈100 linhas)
engines/                lógica de validação — sem Streamlit, uma pasta por app
  common/               paths, loaders, rolesviews, relatórios Excel, logs de execução
  material_rules/       regras de material master partilhadas (fields, cross_file, cleanup_rules, ...)
  sap_material/         SAP Material Validation (Palmela): pipeline, postprocess, fp_status
  tunisia/              Tunisia (TN01/TN02): rules + pipeline
  indirect_materials/   Indirect Materials (ERSA/HIBE/VERP)
  data_cleanups/        Data Cleanups
  plm_workflow/         PLM Workflow (ATL / Change Cockpit)
  progress/             SAP Validation Progress (comparação de relatórios)
  ssa/                  SSA Validation
  assignments/          owners, atribuições e histórico de erros
services/               lógica de aplicação sem UI (ficheiros carregados, identidade)
ui/                     tema (theme.py), layout, componentes e dashboards/ (run_views, history, evolution)
views/                  uma página por módulo, cada uma com render(start_time)
assets/                 styles.css, theme_overrides.css, dashboard.css, logo, RolesViews.xlsx
tests/                  smoke tests (imports + render de todas as páginas)
tools/                  scripts de apoio (check_columns.py)
```

Regra de dependência: `views → ui/services → engines`. Os `engines` nunca importam Streamlit.

## Persistência
Caminhos centralizados em `engines/common/paths.py` (`.history/`, `reports/`).
Podem ser redefinidos com `PLM_WEB_HISTORY_DIR` e `PLM_WEB_REPORTS_DIR`.

## Correr
```
pip install -r requirements.txt
streamlit run app.py
```

## Testes
```
pip install -r requirements-dev.txt
pytest
```

## Temas e estilos
`ui/theme.py` é a fonte única das cores: cada tema define tokens CSS (`--app-bg`, `--primary`, ...)
injetados em `:root` a cada rerun, mais aliases (`--bg-card`, `--txt-main`, ...) usados pelos dashboards.
Para adicionar um tema, basta uma nova entrada em `THEMES`.

## Relatórios Excel
`engines/common/excel_report.py` aplica o estilo de cada relatório e `excel_polish.py` acrescenta, no fim,
filtros, gridlines off, data bars, configuração de impressão e um guia de folhas com hiperlinks.
O polish é aditivo e nunca bloqueia a geração do relatório.
