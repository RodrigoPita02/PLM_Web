"""
Corre este script na tua pasta do projecto para listar as colunas 
exactas do teu ficheiro MARA.

Uso:
    python check_columns_mara.py <caminho_para_MARA.xlsx>
"""

import sys
import pandas as pd

if len(sys.argv) < 2:
    print("Uso: python check_columns_mara.py <MARA.xlsx>")
    sys.exit(1)

# Ler apenas header
df = pd.read_excel(sys.argv[1], sheet_name=0, nrows=0)
cols = [str(c).strip() for c in df.columns]

print(f"\n=== {len(cols)} colunas encontradas no MARA ===\n")
for i, c in enumerate(cols):
    print(f"  {i+1:3d}. '{c}'")

# ------------------------------------------------------------
# ✅ COLUNAS QUE O TEU TN ENGINE ESPERA DE MARA
# (baseado no teu build_sap_like_df)
# ------------------------------------------------------------
expected = [
    'Material',
    'Material Type',
    'Description',
    'Base Unit of Measure',
    'Division',
    'Export/Import Group',
    'General Item Category Group',
    'Gross Weight',
    'Net Weight',
    'Unit of Weight (ISO Format)',
    'Material Group',
    'Product Hierarchy',
    'Packaging Mat. Type',
    'Old material number',
]

print(f"\n=== Colunas esperadas pelo TN engine — presença no MARA ===\n")

missing = []

for c in expected:
    found = c in cols
    status = "✅" if found else "❌ MISSING"
    print(f"  {status}  '{c}'")

    if not found:
        missing.append(c)

# ------------------------------------------------------------
# ✅ RESULTADO FINAL
# ------------------------------------------------------------
if missing:
    print(f"\n⚠️  {len(missing)} colunas em falta — o merge MARA vai perder dados!")
else:
    print(f"\n✅ Todas as colunas presentes.")
