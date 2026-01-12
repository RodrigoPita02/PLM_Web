import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from rolesviews_data import get_rolesviews_bytesio  # <-- importa o helper

# ==========================================================
# 1) LOADERS UNIVERSAIS
# ==========================================================

def load_xlsx(file, *, sheet_name=None):
    import pandas as pd, io

    if file is None:
        return None

    # --- Caminho no disco ---
    if isinstance(file, str):
        # Lê diretamente com engine explícito
        return pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)

    # --- UploadedFile (Streamlit) ---
    # Usa getvalue() que é a forma robusta no Streamlit
    if hasattr(file, "getvalue"):
        data = file.getvalue()
        bio = io.BytesIO(data)
        return pd.read_excel(bio, engine="openpyxl", sheet_name=sheet_name)

    # --- Generic file-like (com read/seek) ---
    if hasattr(file, "read"):
        # Se tiver seek, posiciona no início
        try:
            file.seek(0)
        except Exception:
            pass
        data = file.read()
        bio = io.BytesIO(data)
        return pd.read_excel(bio, engine="openpyxl", sheet_name=sheet_name)

    # --- BytesIO ---
    if isinstance(file, io.BytesIO):
        file.seek(0)
        return pd.read_excel(file, engine="openpyxl", sheet_name=sheet_name)

    raise ValueError(f"Formato não suportado: {type(file)}")

def build_material_transfer_map(zplm_marc_df):
    """
    Cria um mapa (Material, Plant) -> bool indicando se 'Material Transfer' está marcado em ZPLM MARC.
    Considera marcados os valores em {'X','1','TRUE','YES','Y'} (case-insensitive).
    """
    import pandas as pd
    if zplm_marc_df is None or zplm_marc_df.empty:
        return {}
    df = zplm_marc_df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    candidates = ['MATERIAL TRANSFER', 'MAT_TRANSFER', 'Z_MATERIAL_TRANSFER']
    mt_col = next((c for c in candidates if c in df.columns), None)
    if mt_col is None or 'MATERIAL' not in df.columns or 'PLANT' not in df.columns:
        return {}

    def is_marked(x):
        s = '' if pd.isna(x) else str(x).strip().upper()
        return s in {'X', '1', 'TRUE', 'YES', 'Y'}

    df['_KEY'] = list(zip(
        df['MATERIAL'].astype(str).str.strip(),
        df['PLANT'].astype(str).str.strip()
    ))
    df['_MARKED'] = df[mt_col].apply(is_marked)
    out = {}
    for key, grp in df.groupby('_KEY'):
        out[key] = grp['_MARKED'].any()
    return out

# ==========================================================
# 2) PLACEHOLDERS PARA AS FUNÇÕES DO TEU SCRIPT
# ==========================================================
# Cola aqui as TUAS FUNÇÕES originais (uma a uma),
# substituindo os "pass" ou a docstring "🔧 ...".


# ------------------------
# Helpers gerais
# ------------------------


def normalize_code(value, *, as_int=False, zero_pad=None):
    """
    Normaliza um valor 'code' vindo do Excel:
    - Remove espaços.
    - Se for numérico (ex.: 2.0, 40.0), converte para int e depois para str.
    - Aplica zero padding (ex.: zero_pad=2 -> '0' -> '00').
    - Retorna string normalizada.
    """
    import pandas as pd
    if pd.isna(value):
        return ''
    s = str(value).strip()

    # Se for algo tipo '2.0', '40.0' ou '2', tenta converter para int
    try:
        num = float(s)
        if as_int or s.endswith('.0') or s.isdigit():
            s = str(int(num))
    except ValueError:
        # não numérico, mantém como está (ex.: '02')
        pass

    # zero padding, se pedido
    if zero_pad:
        try:
            s_int = int(s)
            s = str(s_int).zfill(zero_pad)
        except ValueError:
            # se não der para int (ex.: '00'), pad manual
            if len(s) < zero_pad and s.isdigit():
                s = s.zfill(zero_pad)
    return s


def ensure_basic_data_k(marc_df):
    """
    Garante que 'Maintenance status' contém a letra 'K' (Basic Data) para todos os materiais.
    - Se o campo estiver vazio/NaN, torna-se 'K'.
    - Se já tiver 'K' (mesmo repetida), mantém apenas uma ocorrência.
    - Não altera a ordem das outras letras; apenas acrescenta 'K' no fim quando necessário.
    """
    import pandas as pd
    import re

    col = 'Maintenance status'
    if col not in marc_df.columns:
        print(f"⚠️ MARC não tem a coluna '{col}'.")
        return marc_df

    def add_k(s):
        s = '' if pd.isna(s) else str(s).strip()
        # remove duplicações de K caso existam
        s = re.sub(r'K+', 'K', s)
        if 'K' in s:
            return s
        # se vazio: vira 'K'; se já tiver letras: acrescenta 'K' no fim
        return 'K' if s == '' else (s + 'K')

    marc_df[col] = marc_df[col].apply(add_k)
    return marc_df



# ------------------------
# Validações SAPUI5 Export
# ------------------------


def validar_campo_obrigatorio(df, coluna, resultados):
    """Valida se uma coluna obrigatória tem valores em branco."""
    if coluna in df.columns:
        missing = df[df[coluna].isnull()]
        for idx in missing.index:
            resultados.append({
                'Line': idx + 2,  # +2: índice base + cabeçalho
                'Field': coluna,
                'Error': 'Mandatory field blank',
                'Material': str(df.at[idx, 'Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(df.at[idx, 'Plant']).strip() if 'Plant' in df.columns else ''
            })
    else:
        print(f"⚠️ The column '{coluna}' was not found in the 'SAPUI5 Export' sheet.")

def validar_procurement_type(df, resultados):
    """
    Valida 'Procurement Type' com precedência de Plant Type sobre Material Type.
    Regras:
    - Dominio universal permitido: E, F, X; campo é obrigatório.
    - Plant Type:
      * PP -> permitido: F ou X (precedência)
      * MP/FP -> permitido: E ou X (precedência)
    - Sem PP/MP/FP (ou Plant Type vazio), aplica Material Type:
      * FERT/HALB -> permitido: E ou X
      * ROH -> permitido: F ou X
    """
    import pandas as pd

    # Garantir colunas necessárias
    needed = {'Procurement Type', 'Material Type', 'Plant Type'}
    if not needed.issubset(df.columns):
        # podes trocar por logging se preferires
        # print("⚠️ Columns 'Procurement Type', 'Material Type' or 'Plant Type' not found.")
        return

    universal = {'E', 'F', 'X'}

    for idx, row in df.iterrows():
        erros = []
        material = str(row.get('Material', '')).strip() if 'Material' in df.columns else ''
        plant = str(row.get('Plant', '')).strip() if 'Plant' in df.columns else ''

        raw_pt = row['Procurement Type']
        pt = '' if pd.isna(raw_pt) else str(raw_pt).strip().upper()
        mt = '' if pd.isna(row['Material Type']) else str(row['Material Type']).strip().upper()
        ptype = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip().upper()

        # 1) Obrigatório e domínio
        if pt == '':
            erros.append('Mandatory field blank')
        elif pt not in universal:
            erros.append(f"Invalid value: '{raw_pt}' (allowed: E, F, X)")

        # 2) Precedência do Plant Type (se válido e presente)
        if pt in universal and ptype != '':
            if ptype == 'PP':
                if pt not in {'F', 'X'}:
                    erros.append(f"For Plant Type 'PP', Procurement Type should be 'F' or 'X' (current: '{pt}')")
                apply_mt = False
            elif ptype in {'MP', 'FP'}:
                if pt not in {'E', 'X'}:
                    erros.append(f"For Plant Type '{ptype}', Procurement Type should be 'E' or 'X' (current: '{pt}')")
                apply_mt = False
            else:
                apply_mt = True
        else:
            apply_mt = True

        # 3) Regras por Material Type (apenas quando não houve precedência de Plant)
        if pt in universal and apply_mt:
            if mt in {'FERT', 'HALB'} and pt not in {'E', 'X'}:
                erros.append(f"For Material Type '{mt}', Procurement Type should be 'E' or 'X' (current: '{pt}')")
            elif mt == 'ROH' and pt not in {'F', 'X'}:
                erros.append(f"For Material Type 'ROH', Procurement Type should be 'F' or 'X' (current: '{pt}')")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Procurement Type',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_account_assignment_group(df, resultados, mvke_df):
    """
    Valida 'Account Assignment Group' usando o SAPUI5 Export (Item Browse) + MVKE.xlsx:

    Escopo:
    - Só materiais cujo Material Type ∈ {FERT, HALB, ROH} (do df SAPUI5).

    Regras em MVKE:
    A) Cobertura FERT:
       - Se existir em MVKE (para o material) Sales Organization = 'PT00' e Distribution Channel ∈ {'10', '20'},
         então TEM de existir também um registo com Sales Organization = 'US00' e Distribution Channel = '30'.
    B) Account Assignment (Acct Assmt Grp Mat.):
       - Campos em branco em qualquer linha de MVKE → erro por linha (Mandatory field blank).
       - Por material, cobrar "Missing 'Z1'" apenas se existir pelo menos um AAG preenchido (e nenhum 'Z1').

    Observações:
    - 'Line' é vazio para erros vindos de MVKE (consistente com outras validações cross-file).
    - 'Plant' vazio (MVKE não é por planta).
    """
    import pandas as pd

    # 1) Garantir colunas mínimas no SAPUI5 Export
    for col in ['Material', 'Material Type']:
        if col not in df.columns:
            # print(f"⚠️ Column '{col}' not found in SAPUI5 Export; skipping Account Assignment (MVKE) validation.")
            return

    # 2) Materiais elegíveis (FERT/HALB/ROH)
    valid_types = {'FERT', 'HALB', 'ROH'}
    sap_subset = (
        df.loc[df['Material Type'].astype(str).str.strip().isin(valid_types), ['Material', 'Material Type']]
        .copy()
    )
    sap_subset['Material'] = sap_subset['Material'].astype(str).str.strip()
    mt_map = dict(zip(sap_subset['Material'], sap_subset['Material Type']))
    material_list = set(mt_map.keys())
    if not material_list or mvke_df is None or mvke_df.empty:
        return

    # 3) Preparar MVKE (normalização de cabeçalhos e valores)
    mvke = mvke_df.copy()
    mvke.columns = [c.strip() for c in mvke.columns]

    def pick_col(candidates):
        for c in candidates:
            if c in mvke.columns:
                return c
        return None

    col_material = pick_col(['Material', 'MATNR'])
    col_salesorg = pick_col(['Sales Organization', 'Sales Org', 'VKORG'])
    col_distchan = pick_col(['Distribution Channel', 'Distr. Channel', 'VTWEG'])
    col_aag = pick_col(['Acct Assmt Grp Mat.', 'Account Assignment Group', 'KTGRM'])

    required_cols = {
        'Material': col_material,
        'Sales Organization': col_salesorg,
        'Distribution Channel': col_distchan,
        'Acct Assmt Grp Mat.': col_aag,
    }
    missing = [k for k, v in required_cols.items() if v is None]
    if missing:
        # print(f"⚠️ MVKE is missing required columns: {missing}. Skipping Account Assignment (MVKE) validation.")
        return

    def norm_str(x):
        return '' if pd.isna(x) else str(x).strip()

    def norm_dc(x):
        try:
            return normalize_code(x, as_int=True)  # '10.0' -> '10', etc.
        except Exception:
            return norm_str(x)

    mvke['Material_norm'] = mvke[col_material].astype(str).str.strip()
    mvke['SalesOrg_norm'] = mvke[col_salesorg].astype(str).str.strip().str.upper()
    mvke['DistChan_norm'] = mvke[col_distchan].apply(norm_dc)
    mvke['AAG_norm'] = mvke[col_aag].apply(norm_str)

    # Filtrar MVKE apenas para materiais elegíveis do SAPUI5 Export
    mvke = mvke.loc[mvke['Material_norm'].isin(material_list)].copy()
    if mvke.empty:
        return

    # 4) Loop por material
    for material, grp in mvke.groupby('Material_norm'):
        mt = mt_map.get(material, '')

        # B.1) Reportar todos os AAG vazios por linha
        blank_rows = grp.loc[grp['AAG_norm'] == '']
        for _, r in blank_rows.iterrows():
            so = r['SalesOrg_norm']
            dc = r['DistChan_norm']
            resultados.append({
                'Line': '',
                'Field': 'Account Assignment Group',
                'Error': f"Mandatory field blank (MVKE)\n Sales Org='{so}', Distribution Channel='{dc}'",
                'Material': material,
                'Plant': ''
            })

        # B.2) Cobrar 'Z1' apenas se existir algum AAG preenchido
        non_blank_mask = grp['AAG_norm'].str.len() > 0
        if non_blank_mask.any():
            has_z1 = (grp.loc[non_blank_mask, 'AAG_norm'].str.upper() == 'Z1').any()
            if not has_z1:
                resultados.append({
                    'Line': '',
                    'Field': 'Account Assignment Group',
                    'Error': "Missing 'Z1' in MVKE: at least one Sales Org/Distribution Channel must have 'Z1'",
                    'Material': material,
                    'Plant': ''
                })

        # A) Cobertura FERT: PT00 (10/20) -> requer US00 (30)
        if mt == 'FERT':
            has_pt00_10_20 = grp.loc[
                (grp['SalesOrg_norm'] == 'PT00') &
                (grp['DistChan_norm'].isin({'10', '20'}))
            ].shape[0] > 0
            if has_pt00_10_20:
                has_us00_30 = grp.loc[
                    (grp['SalesOrg_norm'] == 'US00') &
                    (grp['DistChan_norm'] == '30')
                ].shape[0] > 0
                if not has_us00_30:
                    resultados.append({
                        'Line': '',
                        'Field': 'Account Assignment Group',
                        'Error': ("FERT in MVKE with PT00 (DC 10/20) requires US00 (DC 30) "
                                  "for the same Material — missing US00/30"),
                        'Material': material,
                        'Plant': ''
                    })

def validar_automatic_purchase_order(df, resultados):
    if 'Plant Type' not in df.columns or 'Automatic Purchase Order Allowed' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Automatic Purchase Order Allowed' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant_type = str(row['Plant Type']).strip()
        auto_order = str(row['Automatic Purchase Order Allowed']).strip()

        if pd.isna(row['Automatic Purchase Order Allowed']) or auto_order == '':
            erros_linha.append('Mandatory field blank')
        else:
            if plant_type == 'PP' and auto_order != '1':
                erros_linha.append(f"For Plant Type 'PP', should be 'Yes' (Current: '{auto_order}')")
            elif plant_type != 'PP' and auto_order != '0':
                erros_linha.append(f"For Plant Type '{plant_type}', should be 'No' (Current: '{auto_order}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Automatic Purchase Order Allowed',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_batch_management(df, resultados):
    if 'Plant Type' not in df.columns or 'Batch Management Required Ind.' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Batch Management Required Ind.' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant_type = str(row['Plant Type']).strip()
        batch_ind = str(row['Batch Management Required Ind.']).strip()

        if pd.isna(row['Batch Management Required Ind.']) or batch_ind == '':
            erros_linha.append('Mandatory field blank')
        else:
            if plant_type == 'PP' and batch_ind != '1':
                erros_linha.append(f"For Plant Type 'PP', should be 'Yes' (Current: '{batch_ind}')")
            elif plant_type != 'PP' and batch_ind != '0':
                erros_linha.append(f"For Plant Type '{plant_type}', should be 'No' (Current: '{batch_ind}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Batch Management Required Ind.',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_consumption_mode(df, resultados):
    if 'Consumption Mode' not in df.columns:
        print("⚠️ column 'Consumption Mode' not found.")
        return
    for idx, row in df.iterrows():
        erros_linha = []
        raw = row['Consumption Mode']
        val = normalize_code(raw, as_int=True)  # '2.0' -> '2', '02' -> '2'
        if val == '':
            erros_linha.append('Mandatory field blank')
        elif val != '2':
            erros_linha.append(f"Invalid value: '{raw}' (should be '2')")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Consumption Mode',
                'Error': '\n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_division(df, resultados):
    """Valida se o Field 'Division' é sempre '00' e não está vazio (aceitando '0'/'0.0' e normalizando)."""
    if 'Division' not in df.columns:
        print("⚠️ Collumn 'Division' not found.")
        return
    for idx, row in df.iterrows():
        erros_linha = []
        raw = row['Division']
        val = normalize_code(raw, as_int=True, zero_pad=2)  # '0.0' -> '00', '0' -> '00'
        if val == '':
            erros_linha.append('Mandatory field blank')
        elif val != '00':
            erros_linha.append(f"Invalid value: '{raw}' (should be '00')")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Division',
                'Error': '\n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })


def validar_export_import_group(df, resultados):
    """Valida se o Field 'Export/Import Group' é sempre 'ZVIS' e não está vazio."""
    if 'Export/Import Group' not in df.columns:
        print("⚠️ Column 'Export/Import Group' not found.")
        return
    for idx, row in df.iterrows():
        erros_linha = []
        export_import = str(row['Export/Import Group']).strip()
        if pd.isna(row['Export/Import Group']) or export_import == '':
            erros_linha.append('Mandatory field blank')
        elif export_import != 'ZVIS':
            erros_linha.append(f"Invalid value: '{export_import}' (should be 'ZVIS')")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Export/Import Group',
                'Error': '\n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_general_item_category_group(df, resultados):
    """Valida se o Field 'General Item Category Group' é sempre 'NORM' e não está vazio."""
    if 'General Item Category Group' not in df.columns:
        print("⚠️ Column 'General Item Category Group' not found.")
        return
    for idx, row in df.iterrows():
        erros_linha = []
        general_group = str(row['General Item Category Group']).strip()
        if pd.isna(row['General Item Category Group']) or general_group == '':
            erros_linha.append('Mandatory field blank')
        elif general_group != 'NORM':
            erros_linha.append(f"Invalid value: '{general_group}' (should be 'NORM')")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'General Item Category Group',
                'Error': '\n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })


def validar_gross_weight(df, resultados):
    """Valida 'Gross Weight' apenas para Plant Type = FP: obrigatórios (>0), iguais a Net, e limites por unidade."""
    import pandas as pd
    required = {'Gross Weight', 'Net Weight', 'Unit of Weight (ISO Format)', 'Plant Type'}
    if not required.issubset(df.columns):
        print("⚠️ Colunas necessárias não encontradas.")
        return
    for idx, row in df.iterrows():
        # Só FP
        plant_type = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip()
        if plant_type != 'FP':
            continue

        erros_linha = []
        gross = row['Gross Weight']
        net = row['Net Weight']
        unit = str(row['Unit of Weight (ISO Format)']).strip().upper() if not pd.isna(row['Unit of Weight (ISO Format)']) else ''

        if pd.isna(gross) or pd.isna(net):
            erros_linha.append('Mandatory fields blank')
        else:
            try:
                gross_val = float(gross)
                net_val = float(net)
                if gross_val <= 0 or net_val <= 0:
                    erros_linha.append(f"Values must be greater than 0: Gross={gross_val}, Net={net_val}")
                if gross_val != net_val:
                    erros_linha.append(f"Gross Weight ({gross_val}) not equal compared to Net Weight ({net_val})")

                if unit == 'KG':
                    limite = 5
                elif unit == 'G':
                    limite = 5000
                elif unit == 'MG':
                    limite = 5000000
                else:
                    limite = 10
                if gross_val > limite or net_val > limite:
                    erros_linha.append(
                        f"Value too high for unit {unit}: Gross={gross_val}, Net={net_val} (> {limite} {unit})"
                    )
            except ValueError:
                erros_linha.append(f"Invalid value: Gross='{gross}', Net='{net}' (not numeric)")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Gross Weight',
                'Error': ' \n '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_repetitive_manufacturing(df, resultados):
    """Valida regras para 'Indicator: Repetitive Manufa. Allow' considerando Material Type e Plant Type."""
    if 'Material Type' not in df.columns or 'Indicator: Repetitive Manufa. Allow' not in df.columns or 'Plant Type' not in df.columns:
        print("⚠️ Colunas necessárias não encontradas.")
        return
    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        plant_type = str(row['Plant Type']).strip()
        repetitive_allow = str(row['Indicator: Repetitive Manufa. Allow']).strip() if not pd.isna(row['Indicator: Repetitive Manufa. Allow']) else ''

        if repetitive_allow == '':
            erros_linha.append('Mandatory field blank')
        else:
            if material_type == 'ROH':
                if plant_type in ['MP', 'FP']:
                    if repetitive_allow != '1':
                        erros_linha.append(f"For ROH with Plant Type '{plant_type}', should be 'Yes' (Current: '{repetitive_allow}')")
                elif plant_type != '' and repetitive_allow != '0':
                    erros_linha.append(f"For ROH with Plant Type '{plant_type}', should be 'No' (Current: '{repetitive_allow}')")
            elif material_type == 'HALB':
                if plant_type == 'PP':
                    if repetitive_allow == '1':
                        erros_linha.append("For HALB with Plant Type 'PP', shouldn't be 'Yes'")
                elif plant_type != '' and repetitive_allow != '1':
                    erros_linha.append(f"For HALB with Plant Type '{plant_type}', should be 'Yes' (Current: '{repetitive_allow}')")
            elif material_type == 'FERT':
                if plant_type == 'PP':
                    if repetitive_allow == '1':
                        erros_linha.append("For FERT with Plant Type 'PP', shouldn't be 'Yes'")
                elif plant_type != '' and repetitive_allow != '1':
                    erros_linha.append(f"For FERT with Plant Type '{plant_type}', should be 'Yes' (Current: '{repetitive_allow}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Indicator: Repetitive Manufa. Allow',
                'Error': ' \n '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_item_category_group(df, resultados):
    """Valida se o Field 'Item Category Group' não está vazio e só aceita 'NORM' or 'Multiple'."""
    if 'Item Category Group' not in df.columns:
        print("⚠️ Column 'Item Category Group' not found.")
        return
    valores_validos = ['NORM', 'Multiple']
    for idx, row in df.iterrows():
        erros_linha = []
        item_group = str(row['Item Category Group']).strip()
        if pd.isna(row['Item Category Group']) or item_group == '':
            erros_linha.append('Mandatory field blank')
        elif item_group not in valores_validos:
            erros_linha.append(f"Invalid value: '{item_group}' (allowed: NORM or Multiple)")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Item Category Group',
                'Error': ' \n '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_jit_delivery_sched(df, resultados):
    """Valida se 'JIT Delivery Sched.' é '1' para Plant Type = PP e vazio para os restantes."""
    if 'Plant Type' not in df.columns or 'JIT Delivery Sched.' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'JIT Delivery Sched.' not found.")
        return
    for idx, row in df.iterrows():
        erros_linha = []
        plant_type = str(row['Plant Type']).strip()
        jit_sched = str(row['JIT Delivery Sched.']).strip()
        if plant_type == 'PP':
            if pd.isna(row['JIT Delivery Sched.']) or jit_sched == '':
                erros_linha.append("Mandatory field blank For Plant Type 'PP'")
            elif jit_sched != '1.0':
                erros_linha.append(f"For Plant Type 'PP', should be 'Yes' (Current: '{jit_sched}')")
        else:
            if jit_sched != '' and not pd.isna(row['JIT Delivery Sched.']):
                erros_linha.append(f"For Plant Type '{plant_type}', o Field deve estar vazio (Current: '{jit_sched}')")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'JIT Delivery Sched.',
                'Error': ' \n '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_loading_group(df, resultados):
    if 'Loading Group' not in df.columns:
        return
    for idx, row in df.iterrows():
        erros_linha = []
        raw = row['Loading Group']
        val = normalize_code(raw, as_int=True, zero_pad=4)  # '1.0' -> '0001', '1' -> '0001'
        if val == '':
            erros_linha.append('Mandatory field blank')
        elif val != '0001':
            erros_linha.append(f"Invalid value: '{raw}' (should be '0001')")
        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Loading Group',
                'Error': '\n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_lot_sizing_procedure(df, resultados):
    """
    Regras:
    - Campo obrigatório.
    - EX só permitido para HALB ou FERT (regra base).
    - ZK só permitido para ROH (regra base).
    - EXCEÇÃO: para HALB cujo Material contenha o token '14K450' (ex.: 'VPAKLF-14K450-BCD'),
      então Lot Sizing Procedure pode ser 'ZK' ou 'EX'.
    """
    if 'Material Type' not in df.columns or 'Lot Sizing Procedure' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Lot Sizing Procedure' not found.")
        return

    import re
    # token como palavra (delimitada por não-alfanuméricos)
    token_pattern = re.compile(r'(?i)(?:^|[^A-Z0-9])14K450(?:[^A-Z0-9]|$)')

    def material_has_14k450_token(material_value: str) -> bool:
        if material_value is None:
            return False
        s = str(material_value).strip()
        return bool(token_pattern.search(s))

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip() if 'Material Type' in df.columns else ''
        lot_sizing = '' if pd.isna(row['Lot Sizing Procedure']) else str(row['Lot Sizing Procedure']).strip()
        material = str(row['Material']).strip() if 'Material' in df.columns else ''

        # 1) Campo obrigatório
        if lot_sizing == '':
            erros_linha.append('Mandatory field blank')
        else:
            # 2) Exceção HALB + token 14K450 → aceita ZK ou EX
            is_halb_14k450 = (material_type == 'HALB' and material_has_14k450_token(material))
            if is_halb_14k450:
                if lot_sizing not in {'EX', 'ZK'}:
                    erros_linha.append(
                        f"Invalid value: '{lot_sizing}' (for HALB with Material containing token '14K450', allowed: ZK or EX)"
                    )
            else:
                # 3) Regras padrão
                if lot_sizing == 'EX' and material_type not in ['HALB', 'FERT']:
                    erros_linha.append(f"Value 'EX' only allowed for HALB or FERT (current: '{material_type}')")
                if lot_sizing == 'ZK' and material_type != 'ROH':
                    erros_linha.append(f"Value 'ZK' only allowed for ROH (current: '{material_type}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Lot Sizing Procedure',
                'Error': ' \n'.join(erros_linha),
                'Material': material,
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_material_group(df, resultados):
    if 'Material Type' not in df.columns or 'Material Group' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Material Group' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        material_group = str(row['Material Group']).strip()

        # obrigatório
        if pd.isna(row['Material Group']) or material_group == '':
            erros_linha.append('Mandatory field blank')
        else:
            if material_type == 'HALB' and material_group in ['L004', 'APPOINT']:
                erros_linha.append(f"For HALB, can not be '{material_group}'")
            elif material_type == 'FERT' and material_group in ['L003', 'APPOINT']:
                erros_linha.append(f"For FERT, can not be '{material_group}'")
            elif material_type == 'ROH' and material_group in ['L003', 'L004']:
                erros_linha.append(f"For ROH, can not be '{material_group}'")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Material Group',
                'Error': ' \n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_net_weight(df, resultados):
    """Valida 'Net Weight' apenas para Plant Type = FP: obrigatórios (>0), iguais a Gross, e limites por unidade."""
    import pandas as pd
    required = {'Gross Weight', 'Net Weight', 'Unit of Weight (ISO Format)', 'Plant Type'}
    if not required.issubset(df.columns):
        print("⚠️ Colunas necessárias não encontradas.")
        return

    for idx, row in df.iterrows():
        # Só FP
        plant_type = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip()
        if plant_type != 'FP':
            continue

        erros_linha = []
        gross = row['Gross Weight']
        net = row['Net Weight']
        unit = str(row['Unit of Weight (ISO Format)']).strip().upper() if not pd.isna(row['Unit of Weight (ISO Format)']) else ''

        if pd.isna(gross) or pd.isna(net):
            erros_linha.append('Mandatory fields blank')
        else:
            try:
                gross_val = float(gross)
                net_val = float(net)

                if net_val <= 0 or gross_val <= 0:
                    erros_linha.append(f"Values must be greater than 0: Gross={gross_val}, Net={net_val}")

                if gross_val != net_val:
                    erros_linha.append(f"Gross Weight ({gross_val}) not equal compared to Net Weight ({net_val})")

                if unit == 'KG':
                    limite = 5
                elif unit == 'G':
                    limite = 5000
                elif unit == 'MG':
                    limite = 5000000
                else:
                    limite = 10

                if gross_val > limite or net_val > limite:
                    erros_linha.append(
                        f"Value too high for unit {unit}: Gross={gross_val}, Net={net_val} (> {limite} {unit})"
                    )
            except ValueError:
                erros_linha.append(f"Invalid value: Gross='{gross}', Net='{net}' (not numeric)")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Net Weight',
                'Error': ' \n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_period_type(df, resultados):
    if 'Period Type' not in df.columns:
        print("⚠️ Column 'Period Type' not found.")
        return

    valores_validos = ['T', 'P', 'M', 'W']

    for idx, row in df.iterrows():
        erros_linha = []
        period_type = str(row['Period Type']).strip()

        if pd.isna(row['Period Type']) or period_type == '':
            erros_linha.append('Mandatory field blank')
        elif period_type not in valores_validos:
            erros_linha.append(f"Invalid value: '{period_type}' (allowed: T, P, M, W)")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Period Type',
                'Error': ' \n'.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })
def validar_planning_calendar(df, resultados):
    """
    Validação de Planning Calendar quando PSMS ∈ [10, 70], considerando:
    - Plant (PT01 vs PT02)
    - Material Type (FERT, HALB, ROH)
    - Materiais cujo nome contenha '14K450' (em qualquer posição, case-insensitive)

    Regras (PSMS ∈ [10,70]):
      PT01:
        - FERT -> 'PT1'
        - HALB -> 'PT1'
        - ROH  -> não vazio e != 'PT2'
        - Material com '14K450' -> não vazio e != 'PT2'
      PT02:
        - FERT -> 'PT2'
        - HALB -> 'PT2'
        - ROH  -> não vazio (qualquer valor presente)
        - Material com '14K450' -> não vazio (qualquer valor presente)
    """
    import pandas as pd

    required = {'Material', 'Plant', 'Material Type', 'Planning Calendar', 'Plant Specific Material Status'}
    if not required.issubset(df.columns):
        missing = [c for c in required if c not in df.columns]
        print(f"⚠️ Missing columns for Planning Calendar validation: {missing}")
        return

    def parse_status(val):
        """Converte PSMS para inteiro (aceita '10', '10.0'). Retorna None se vazio/NaN ou inválido."""
        if pd.isna(val) or str(val).strip() == '':
            return None
        s = str(val).strip()
        try:
            return int(float(s))
        except Exception:
            return None

    def contains_14k450(material):
        """True se o nome do material contiver '14K450' (case-insensitive), em qualquer posição."""
        if material is None:
            return False
        return '14K450' in str(material).upper()

    for idx, row in df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        mt       = str(row.get('Material Type', '')).strip().upper()
        pc_raw   = row.get('Planning Calendar', None)
        pc       = '' if pd.isna(pc_raw) else str(pc_raw).strip().upper()
        psms     = parse_status(row.get('Plant Specific Material Status', None))
        has14k   = contains_14k450(material)

        # Validar apenas quando PSMS ∈ [10, 70]
        if psms is None or not (10 <= psms <= 70):
            continue

        # Regras por Plant (sem override para PSMS=25)
        if plant == 'PT01':
            if mt == 'FERT':
                if pc != 'PT1':
                    erros.append(f"For Plant PT01 and Material Type 'FERT', Planning Calendar must be 'PT1' (Current: '{pc_raw}')")
            elif mt == 'HALB':
                if pc != 'PT1':
                    erros.append(f"For Plant PT01 and Material Type 'HALB', Planning Calendar must be 'PT1' (Current: '{pc_raw}')")
            elif mt == 'ROH':
                if pc == '':
                    erros.append("For Plant PT01 and Material Type 'ROH', Planning Calendar must not be empty")
                elif pc == 'PT2':
                    erros.append("For Plant PT01 and Material Type 'ROH', Planning Calendar must not be 'PT2'")

            # Materiais com '14K450' em PT01: não vazio e != PT2
            if has14k:
                if pc == '':
                    erros.append("For Plant PT01 and Material containing '14K450', Planning Calendar must not be empty")
                elif pc == 'PT2':
                    erros.append("For Plant PT01 and Material containing '14K450', Planning Calendar must not be 'PT2'")

        elif plant == 'PT02':
            if mt == 'FERT':
                if pc != 'PT2':
                    erros.append(f"For Plant PT02 and Material Type 'FERT', Planning Calendar must be 'PT2' (Current: '{pc_raw}')")
            elif mt == 'HALB':
                if pc != 'PT2':
                    erros.append(f"For Plant PT02 and Material Type 'HALB', Planning Calendar must be 'PT2' (Current: '{pc_raw}')")
            elif mt == 'ROH':
                if pc == '':
                    erros.append("For Plant PT02 and Material Type 'ROH', Planning Calendar must not be empty")

            # Materiais com '14K450' em PT02: não vazio (qualquer valor presente é aceite)
            if has14k and pc == '':
                erros.append("For Plant PT02 and Material containing '14K450', Planning Calendar must not be empty")

        else:
            # Sem regra definida para outros Plants → não validar
            pass

        if erros:
            resultados.append({
                'Line': idx + 2,
                               'Field': 'Planning Calendar',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_planning_time_fence(df, resultados):
    """
    TF semântica:
    - Yes -> TF > 0 (qualquer positivo)
    - No  -> TF = 0 ou vazio

    Regras por Material Type:
    - FERT -> TF > 0
    - ROH  -> TF = 0 ou vazio
    - HALB -> TF pode ser 0/vazio ou >0

    Precedências por PSMS:
    - PSMS = 10 -> ignora a linha (não valida TF).
    - PSMS = 25 -> TF TEM de ser 0 (estritamente 0; vazio não é aceite aqui).
    """
    import pandas as pd

    required = {'Material', 'Plant', 'Material Type', 'Planning Time Fence'}
    if not required.issubset(df.columns):
        print("⚠️ Columns 'Material', 'Plant', 'Material Type' or 'Planning Time Fence' not found.")
        return

    def parse_tf(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return float(str(val).strip())
        except Exception:
            return 'INVALID'

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except Exception:
            return None  # não precisamos de códigos não numéricos aqui

    for idx, row in df.iterrows():
        erros = []
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        mt       = str(row.get('Material Type', '')).strip().upper()
        tf_raw   = row.get('Planning Time Fence', None)
        tf_val   = parse_tf(tf_raw)
        psms_val = parse_status(row.get('Plant Specific Material Status', None))

        # PSMS=10: não validar TF
        if psms_val == 10:
            continue

        # Valor inválido
        if tf_val == 'INVALID':
            erros.append(f"Invalid value: '{tf_raw}' (should be numeric, or empty to represent 0)")
        else:
            has_tf = (tf_val is not None) and (tf_val > 0)
            is_zero_or_empty = (tf_val is None) or (tf_val == 0)

            # PSMS=25: TF TEM de ser 0 estrito (não vazio)
            if psms_val == 25:
                if tf_val is None:
                    erros.append("For Status '25', Planning Time Fence must be 0 (Current: empty)")
                elif tf_val != 0:
                    erros.append(f"For Status '25', Planning Time Fence must be 0 (Current: '{tf_raw}')")
            else:
                # Regras normais por MT
                if mt == 'FERT':
                    if not has_tf:
                        erros.append(f"For FERT, Planning Time Fence should be > 0 (Current: '{tf_raw}')")
                elif mt == 'ROH':
                    if not is_zero_or_empty:
                        erros.append(f"For ROH, Planning Time Fence should be 0 or empty (Current: '{tf_raw}')")
                elif mt == 'HALB':
                    pass  # aceita ambos
                else:
                    pass

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Planning Time Fence',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_mrp_type(df, resultados, zplm_marc_df=None):
    """
    MRP Type considerando:
    - TF com semântica: TF>0 (Yes) ; TF=0 ou vazio (No)
    - HALB transfer (ZPLM MARC)
    - PSMS (Plant Specific Material Status): precedências 10 e 25
    - PLANT (PT01/PT02) para coerência do Planning Calendar quando PSMS=25

    Precedências por Status:
    - PSMS = 10 -> Ignora a linha (não valida nem contabiliza erros).
    - PSMS = 25 -> MRP = ND ; TF = 0 (estritamente 0; vazio não aceite);
                    Planning Calendar:
                        PT01 -> 'PT1'
                        PT02 -> 'PT2'
    Tabela base (aplicada quando não é 10/25):
    - FERT -> MRP = M1, TF > 0
    - HALB transfer -> MRP ∈ {M1, P1, PD}, TF pode ser >0 ou 0/vazio
    - HALB sem transfer -> MRP = PD, TF = 0 ou vazio
    - ROH com TF (>0) -> MRP = P1
    - ROH sem TF (0/vazio) -> MRP = PD
    """
    import pandas as pd

    needed = {'MRP Type', 'Material Type', 'Planning Time Fence', 'Plant'}
    if not needed.issubset(df.columns):
        # print("⚠️ Columns 'MRP Type', 'Material Type', 'Planning Time Fence', 'Plant' not found.")
        return

    has_psms = 'Plant Specific Material Status' in df.columns
    has_pc = 'Planning Calendar' in df.columns

    valores_permitidos = ['M1', 'PD', 'P1', 'ND', 'V1']

    transfer_map = {}
    if zplm_marc_df is not None:
        try:
            transfer_map = build_material_transfer_map(zplm_marc_df)
        except Exception:
            transfer_map = {}

    def parse_tf(val):
        """
        None se vazio/NaN (interpreta como No)
        float se numérico (permite >0 como Yes; 0 como No)
        'INVALID' se não numérico
        """
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return float(str(val).strip())
        except Exception:
            return 'INVALID'

    def parse_status(val):
        """PSMS int (aceita '10', '10.0'); None se vazio/NaN/inválido."""
        if val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except Exception:
            return None

    for idx, row in df.iterrows():
        erros_linha = []

        mrp_type = '' if pd.isna(row['MRP Type']) else str(row['MRP Type']).strip().upper()
        material_type = '' if pd.isna(row['Material Type']) else str(row['Material Type']).strip().upper()
        plant = str(row['Plant']).strip() if not pd.isna(row['Plant']) else ''

        tf_raw = row.get('Planning Time Fence', None)
        tf_val = parse_tf(tf_raw)

        psms_val = parse_status(row.get('Plant Specific Material Status', None)) if has_psms else None
        planning_calendar = '' if (not has_pc or pd.isna(row.get('Planning Calendar', None))) else str(row.get('Planning Calendar')).strip().upper()

        material = str(row['Material']).strip() if 'Material' in df.columns else ''
        is_transfer = transfer_map.get((material, plant), False)

        # ——— Precedências por Status ———
        # PSMS = 10: ignorar completamente (não adicionar erros para esta linha)
        if psms_val == 10:
            continue

        # PSMS = 25: regras obrigatórias
        if psms_val == 25:
            # 1) MRP Type obrigatoriamente 'ND'
            if mrp_type == '':
                erros_linha.append("Mandatory field blank (MRP Type)")
            elif mrp_type != 'ND':
                erros_linha.append(f"For Plant Specific Material Status '25', MRP Type must be 'ND' (Current: '{mrp_type}')")

            # 2) TF obrigatoriamente = 0 (vazio não aceite; tem de ser numericamente 0)
            if tf_val is None:
                erros_linha.append("For Plant Specific Material Status '25', Planning Time Fence must be 0 (Current: empty)")
            elif tf_val == 'INVALID':
                erros_linha.append("Planning Time Fence invalid or not numeric; for Status '25' it must be 0")
            elif tf_val != 0:
                erros_linha.append(f"For Plant Specific Material Status '25', Planning Time Fence must be 0 (Current: '{tf_raw}')")

            # 3) Planning Calendar coerente com a Plant
            if has_pc:
                if plant == 'PT01':
                    if planning_calendar != 'PT1':
                        erros_linha.append(
                            f"For Status '25' and Plant 'PT01', Planning Calendar must be 'PT1' (Current: '{planning_calendar}')"
                        )
                elif plant == 'PT02':
                    if planning_calendar != 'PT2':
                        erros_linha.append(
                            f"For Status '25' and Plant 'PT02', Planning Calendar must be 'PT2' (Current: '{planning_calendar}')"
                        )
                # Outras Plants: não impomos Calendar aqui

            if erros_linha:
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'MRP Type',
                    'Error': ' \n'.join(erros_linha),
                    'Material': material,
                    'Plant': plant
                })
            # Short-circuit: não aplicar mais regras desta linha
            continue

        # ——— Validações base (fora das precedências 10/25) ———
        # Campo obrigatório + domínio
        if mrp_type == '':
            erros_linha.append('Mandatory field blank')
        else:
            if mrp_type not in valores_permitidos:
                erros_linha.append(f"Invalid value: '{mrp_type}' (allowed: {', '.join(valores_permitidos)})")

        # TF inválido (texto não numérico) → semântica não aplicável
        if tf_val == 'INVALID':
            erros_linha.append("Planning Time Fence invalid or not numeric; MRP rule depends on TF > 0 / TF = 0 or empty")
        else:
            # Interpretação do TF (fora das precedências):
            has_tf = (tf_val is not None) and (tf_val > 0)   # Yes
            no_tf = (tf_val is None) or (tf_val == 0)        # No (0 ou vazio)

            # ——— Regras por Material Type ———
            if material_type == 'FERT':
                # M1 e TF > 0
                if mrp_type != 'M1':
                    erros_linha.append(f"For Plant {plant} & Material Type 'FERT', MRP Type should be 'M1' (Current: '{mrp_type}')")
                if not has_tf:
                    erros_linha.append(f"For Plant {plant} & Material Type 'FERT', Planning Time Fence should be > 0 (Current: '{tf_raw}')")

            elif material_type == 'HALB':
                if is_transfer:
                    # M1/P1/PD ; TF livre (>0 ou 0/vazio)
                    if mrp_type not in {'M1', 'P1', 'PD'}:
                        erros_linha.append(
                            f"For Plant {plant} & HALB with Material Transfer, allowed MRP Type: M1, P1 or PD (Current: '{mrp_type}')"
                        )
                else:
                    # Sem transfer: PD ; TF = 0/vazio
                    if mrp_type != 'PD':
                        erros_linha.append(
                            f"For Plant {plant} & HALB without Material Transfer, MRP Type should be 'PD' (Current: '{mrp_type}')"
                        )
                    if not no_tf:
                        erros_linha.append(
                            f"For Plant {plant} & HALB without Material Transfer, Planning Time Fence should be 0 or empty (Current: '{tf_raw}')"
                        )

            elif material_type == 'ROH':
                if has_tf:
                    # Com TF (>0): P1
                    if mrp_type != 'P1':
                        erros_linha.append(
                            f"For Plant {plant} & ROH with Time Fence (>0), MRP Type should be 'P1' (Current: '{mrp_type}')"
                        )
                elif no_tf:
                    # Sem TF (0/vazio): PD
                    if mrp_type != 'PD':
                        erros_linha.append(
                            f"For Plant {plant} & ROH without Time Fence (0 or empty), MRP Type should be 'PD' (Current: '{mrp_type}')"
                        )
                # Outros MTs: sem regra específica

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'MRP Type',
                'Error': ' \n'.join(erros_linha),
                'Material': material,
                'Plant': plant
            })

def validar_price_control(df, resultados):
    if 'Price Control' not in df.columns:
        print("⚠️ Column 'Price Control' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        price_control = str(row['Price Control']).strip()

        # Field obrigatório
        if pd.isna(row['Price Control']) or price_control == '':
            erros_linha.append('Mandatory field blank')
        elif price_control != 'S':
            erros_linha.append(f"Invalid value: '{price_control}' (should be 'S')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Price Control',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })
            
def validar_price_unit(df, resultados):
    if 'Price Unit' not in df.columns:
        print("⚠️ Coluna 'Price Unit' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        price_unit = str(row['Price Unit']).strip()

        # Field obrigatório
        if pd.isna(row['Price Unit']) or price_unit == '':
            erros_linha.append('Mandatory field blank')
        elif price_unit != '10000':
            erros_linha.append(f"Invalid value: '{price_unit}' (should be '10000')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Price Unit',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })


def validar_product_hierarchy(df, resultados):
    if 'Material Type' not in df.columns or 'Product Hierarchy' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Product Hierarchy' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        product_hierarchy = str(row['Product Hierarchy']).strip()

        # Regra: obrigatório apenas For FERT
        if material_type == 'FERT':
            if pd.isna(row['Product Hierarchy']) or product_hierarchy == '':
                erros_linha.append("Mandatory field blank For Material Type 'FERT'")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Product Hierarchy',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_production_scheduling_profile(df, resultados):
    if 'Material Type' not in df.columns or 'Production Scheduling Profile' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Production Scheduling Profile' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        scheduling_profile = str(row['Production Scheduling Profile']).strip()

        # Regra: obrigatório e valor fixo For FERT e HALB
        if material_type in ['FERT', 'HALB']:
            if pd.isna(row['Production Scheduling Profile']) or scheduling_profile == '':
                erros_linha.append("Mandatory field blank For Material Type FERT or HALB")
            elif scheduling_profile != 'ZREM01':
                erros_linha.append(f"Invalid value: '{scheduling_profile}' (should be 'ZREM01' For FERT or HALB)")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Production Scheduling Profile',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_profit_center(df, resultados):
    if 'Plant' not in df.columns or 'Profit Center' not in df.columns:
        print("⚠️ Columns 'Plant' or 'Profit Center' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant = str(row['Plant']).strip()
        profit_center = str(row['Profit Center']).strip()

        # Field obrigatório
        if pd.isna(row['Profit Center']) or profit_center == '':
            erros_linha.append('Mandatory field blank')
        else:
            # Regras específicas
            if plant == 'PT01' and profit_center != 'PT00351':
                erros_linha.append(f"For Plant PT01, should be 'PT00351' (Current: '{profit_center}')")
            elif plant == 'PT02' and profit_center != 'PT00361':
                erros_linha.append(f"For Plant PT02, should be 'PT00361' (Current: '{profit_center}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Profit Center',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_purchasing_group(df, resultados):
    if 'Plant Type' not in df.columns or 'Purchasing Group' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Purchasing Group' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant_type = str(row['Plant Type']).strip()
        purchasing_group = str(row['Purchasing Group']).strip()

        # Field obrigatório apenas For Plant Type = PP
        if plant_type == 'PP':
            if pd.isna(row['Purchasing Group']) or purchasing_group == '':
                erros_linha.append("Mandatory field blank For Plant Type 'PP'")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Purchasing Group',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_repetitive_manufacturing_profile(df, resultados):
    if 'Material Type' not in df.columns or 'Repetitive Manufacturing Profile' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Repetitive Manufacturing Profile' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        repetitive_profile = str(row['Repetitive Manufacturing Profile']).strip()

        # Regra: obrigatório e valor fixo For FERT e HALB
        if material_type in ['FERT', 'HALB']:
            if pd.isna(row['Repetitive Manufacturing Profile']) or repetitive_profile == '':
                erros_linha.append("Mandatory field blank For Material Type FERT or HALB")
            elif repetitive_profile != 'Z002':
                erros_linha.append(f"Invalid value: '{repetitive_profile}' (should be 'Z002' For FERT or HALB)")

        # Nova regra: ROH não pode ter Z002
        if material_type == 'ROH' and repetitive_profile == 'Z002':
            erros_linha.append("Material Type ROH não pode ter valor 'Z002'")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Repetitive Manufacturing Profile',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_source_list_required(df, resultados):
    if 'Plant Type' not in df.columns or 'Source List Required' not in df.columns:
        print("⚠️ Columns 'Plant Type' or 'Source List Required' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant_type = str(row['Plant Type']).strip()
        source_list = str(row['Source List Required']).strip()

        # Field obrigatório
        if pd.isna(row['Source List Required']) or source_list == '':
            erros_linha.append('Mandatory field blank')
        else:
            # Regras específicas
            if plant_type == 'PP' and source_list != '1':
                erros_linha.append(f"For Plant Type PP, should be 'Yes' (Current: '{source_list}')")
            elif plant_type != 'PP' and source_list != '0':
                erros_linha.append(f"For Plant Type {plant_type}, should be 'No' (Current: '{source_list}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Source List Required',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_storage_location(df, resultados):
    if 'Plant' not in df.columns or 'Storage Location' not in df.columns:
        print("⚠️ Columns 'Plant' or 'Storage Location' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant = str(row['Plant']).strip()
        storage_location = str(row['Storage Location']).strip()

        # Field obrigatório
        if pd.isna(row['Storage Location']) or storage_location == '':
            erros_linha.append('Mandatory field blank')
        else:
            # Regras específicas
            if storage_location == 'P105' and plant != 'PT01':
                erros_linha.append(f"Valor 'P105' only allowed for Plant = PT01 (Current: '{plant}')")
            elif storage_location in ['P205', 'P220'] and plant != 'PT02':
                erros_linha.append(f"Valor '{storage_location}' only allowed for Plant = PT02 (Current: '{plant}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Storage Location',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_valuation_class(df, resultados):
    if 'Material Type' not in df.columns or 'Valuation Class' not in df.columns:
        print("⚠️ Columns 'Material Type' or 'Valuation Class' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        valuation_class = str(row['Valuation Class']).strip()

        # Field obrigatório
        if pd.isna(row['Valuation Class']) or valuation_class == '':
            erros_linha.append('Mandatory field blank')
        else:
            # Regras específicas
            if valuation_class == '3000' and material_type != 'ROH':
                erros_linha.append(f"Valor '3000' only allowed for Material Type = ROH (Current: '{material_type}')")
            elif valuation_class == '7900' and material_type != 'HALB':
                erros_linha.append(f"Valor '7900' only allowed for Material Type = HALB (Current: '{material_type}')")
            elif valuation_class == '7920' and material_type != 'FERT':
                erros_linha.append(f"Valor '7920' only allowed for Material Type = FERT (Current: '{material_type}')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Valuation Class',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_with_qty_structure(df, resultados):
    if 'With Qty Structure' not in df.columns:
        print("⚠️ Column 'With Qty Structure' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        qty_structure = str(row['With Qty Structure']).strip()

        # Field obrigatório
        if pd.isna(row['With Qty Structure']) or qty_structure == '':
            erros_linha.append('Mandatory field blank')
        elif qty_structure != '1':
            erros_linha.append(f"Invalid value: '{qty_structure}' (should be 'Yes')")

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'With Qty Structure',
                'Error': ' | '.join(erros_linha),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_enterproj_number(df, resultados):
    if 'Material Type' not in df.columns or 'Enterproj Number' not in df.columns:
        print("⚠️ Column 'Material Type' or 'Enterproj Number' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        material_type = str(row['Material Type']).strip()
        enterproj = str(row['Enterproj Number']).strip() if not pd.isna(row['Enterproj Number']) else ''
        material = str(row['Material']).strip() if 'Material' in df.columns else ''

        # Regras
        if material_type == 'FERT':
            if enterproj == '':
                erros_linha.append("For Material Type 'FERT', the field is required and cannot be empty.")
        elif material_type in ['ROH', 'HALB']:
            if enterproj != '':
                # Converte para inteiro se for numérico (remove .0)
                enterproj_display = enterproj
                if enterproj.replace('.', '', 1).isdigit():
                    try:
                        enterproj_display = str(int(float(enterproj)))
                    except ValueError:
                        enterproj_display = enterproj

                erros_linha.append(f"For Material Type '{material_type}', the field must be empty (atual: '{enterproj_display}')")

        if erros_linha:
            # Verifica se já existe erro para este material (evita duplicação)
            if not any(r['Material'] == material and r['Field'] == 'Enterproj Number' for r in resultados):
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'Enterproj Number',
                    'Error': ' | '.join(erros_linha),
                    'Material': material,
                    'Plant': ''  # Ignora a fábrica
                })

def validar_inspection_type(df, resultados):
    """
    Regras:
      - Campo não pode estar vazio.
      - FERT: NÃO pode ser '08'; permitido '1708'.
      - HALB/ROH: permitido '08' e '1708'.
      - (Outros Material Type, se existirem): permitido '1708' (fallback).
    """
    required_cols = ['Inspection Type', 'Material Type']
    if not set(required_cols).issubset(df.columns):
        print("⚠️ Columns 'Inspection Type' and 'Material Type' not found.")
        return

    def is_code(val, target):
        """Compara códigos aceitando variantes com/sem zero à esquerda (ex.: '8' == '08')."""
        if val is None:
            return False
        v = str(val).strip()
        t = str(target).strip()
        # Se ambos forem numéricos, comparamos como inteiros
        try:
            return int(v) == int(t)
        except ValueError:
            return v == t

    for idx, row in df.iterrows():
        erros = []
        raw_insp = row['Inspection Type']
        # Normaliza: '1708.0' -> '1708'; '08.0'/'8' -> '8'
        insp = normalize_code(raw_insp, as_int=True)
        mt = str(row['Material Type']).strip() if not pd.isna(row['Material Type']) else ''

        # 1) Não pode estar vazio
        if insp == '' or insp is None:
            erros.append('Mandatory field blank')

        else:
            if mt == 'FERT':
                # 2a) FERT não pode ter 08
                if is_code(insp, '8') or is_code(insp, '08'):
                    erros.append(f"Invalid value: '{raw_insp}' (08 is not allowed for FERT)")
                # 2b) FERT deve aceitar 1708
                elif not is_code(insp, '1708'):
                    erros.append(f"Invalid value: '{raw_insp}' (allowed for FERT: 1708)")

            elif mt in ['HALB', 'ROH']:
                # 3) HALB/ROH aceitam 08 e 1708
                if not (is_code(insp, '8') or is_code(insp, '08') or is_code(insp, '1708')):
                    erros.append(f"Invalid value: '{raw_insp}' (allowed for {mt}: 08 or 1708)")

            else:
                # 4) Fallback para outros Material Type (se existirem): só 1708
                if not is_code(insp, '1708'):
                    erros.append(f"Invalid value: '{raw_insp}' (allowed: 1708)")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Inspection Type',
                'Error': ' \n'.join(erros),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_backflush(df, resultados):
    """Backflush: Não pode estar vazio."""
    if 'Backflush' not in df.columns:
        print("⚠️ Column 'Backflush' not found.")
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Backflush']) or str(row['Backflush']).strip() == '':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Backflush',
                'Error': 'Mandatory field blank',
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_strategy_group(df, resultados):
    """Strategy Group: Não pode estar vazio e tem que ser sempre '40'."""
    if 'Strategy Group' not in df.columns:
        print("⚠️ Column 'Strategy Group' not found.")
        return

    for idx, row in df.iterrows():
        erros = []
        raw = row['Strategy Group']
        val = normalize_code(raw, as_int=True)  # '40.0' -> '40', '040' -> '40'

        if val == '':
            erros.append('Mandatory field blank')
        elif val != '40':
            erros.append(f"Invalid value: '{raw}' (should be '40')")

        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Strategy Group',
                'Error': ' \n'.join(erros),
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_product_related_origin(df, resultados):
    """Product-Related Origin: Não pode estar vazio."""
    if 'Product-Related Origin' not in df.columns:
        print("⚠️ Column 'Product-Related Origin' not found.")
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Product-Related Origin']) or str(row['Product-Related Origin']).strip() == '':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Product-Related Origin',
                'Error': 'Mandatory field blank',
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_availability_check(df, resultados):
    """Availability Check: Não pode estar vazio."""
    if 'Availability Check' not in df.columns:
        print("⚠️ Column 'Availability Check' not found.")
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Availability Check']) or str(row['Availability Check']).strip() == '':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Availability Check',
                'Error': 'Mandatory field blank',
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

def validar_mrp_controller(df, resultados):
    """MRP Controller: Não pode estar vazio."""
    if 'MRP Controller' not in df.columns:
        print("⚠️ Column 'MRP Controller' not found.")
        return
    for idx, row in df.iterrows():
        if pd.isna(row['MRP Controller']) or str(row['MRP Controller']).strip() == '':
            resultados.append({
                'Line': idx + 2,
                'Field': 'MRP Controller',
                'Error': 'Mandatory field blank',
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })
    

def validar_plant_specific_material_status(df, resultados):
    """Plant Specific Material Status: Não pode estar vazio."""
    if 'Plant Specific Material Status' not in df.columns:
        print("⚠️ Column 'Plant Specific Material Status' not found.")
        return
    for idx, row in df.iterrows():
        if pd.isna(row['Plant Specific Material Status']) or str(row['Plant Specific Material Status']).strip() == '':
            resultados.append({
                'Line': idx + 2,
                'Field': 'Plant Specific Material Status',
                'Error': 'Mandatory field blank',
                'Material': str(row['Material']).strip() if 'Material' in df.columns else '',
                'Plant': str(row['Plant']).strip() if 'Plant' in df.columns else ''
            })

# ------------------------
# Cruzamentos MARA/MARC/ZPLM
# ------------------------


def add_old_material_number_errors(mara_df, errors_df):
    """
    Adiciona erros à 'errors_df' do relatório anterior para o campo MARA 'Old material number':
    - Não pode estar vazio
    - Deve ser igual ao Material do relatório, removendo os '-' (Material.replace('-', ''))

    Parâmetros:
      mara_df   : DataFrame do MARA (já carregado)
      errors_df : DataFrame lido da aba 'Error Details' do relatório anterior

    Retorna:
      errors_df : DataFrame concatenado com as novas linhas de erro (se houver)
    """
    import pandas as pd

    if 'Material' not in mara_df.columns or 'Old material number' not in mara_df.columns:
        print("⚠️ MARA tem de conter as colunas 'Material' e 'Old material number'.")
        return errors_df

    # Materiais presentes no relatório anterior (já de erro/validação)
    materials = (
        errors_df['Material']
        .dropna()
        .astype(str)
        .str.strip()
        .unique()
    )

    mara_indexed = mara_df.set_index('Material')
    additional_rows = []

    for material in materials:
        expected = material.replace('-', '')

        # Obter o 'Old material number' correspondente em MARA
        if material in mara_indexed.index:
            old_val = mara_indexed.at[material, 'Old material number']
        else:
            old_val = None  # material não encontrado em MARA

        # Normalizar para string
        old_str = '' if pd.isna(old_val) else str(old_val).strip()

        if old_str == '':
            error_msg = 'Mandatory field blank'
        elif old_str != expected:
            error_msg = (
                f"Invalid value: should equal Material without '-' "
                f"(expected: '{expected}', current: '{old_str}')"
            )
        else:
            error_msg = None

        if error_msg:
            additional_rows.append({
                'Line': '',  # não temos linha do SAPUI5; deixamos vazio
                'Field': 'Old material number',
                'Error': error_msg,
                'Material': material,
                'Plant': '',  # MARA é a nível de material
                'View': None,  # será mapeado para 'Not Applicable' mais tarde
                'Justified_By_Missing_View': False,  # estes erros nunca serão justificados por view
                # 'Error_Type' será calculado mais adiante pelo classify_error
            })

    if additional_rows:
        errors_df = pd.concat([errors_df, pd.DataFrame(additional_rows)], ignore_index=True)

    return errors_df


def fill_plant_and_validate_enterproj_from_zplm(
    errors_df,
    latest_report,              # UploadedFile/BytesIO/str path do relatório anterior
    zplm_marc_file=None,        # UploadedFile/BytesIO/str path
    zplm_mara_file=None         # UploadedFile/BytesIO/str path
):
    import pandas as pd

    if errors_df is None or errors_df.empty:
        return errors_df
    for c in ['Field', 'Material', 'Plant', 'Error']:
        if c not in errors_df.columns:
            return errors_df

    # Filtrar linhas-alvo
    mask = (
        errors_df['Field'].astype(str).str.strip() == 'Enterproj Number'
    ) & (
        errors_df['Plant'].astype(str).str.strip().isin(['', 'nan'])
    )
    if not mask.any():
        return errors_df

    subset = errors_df.loc[mask].copy()

    # 👉 Ler ZPLM via load_xlsx (aceita UploadedFile, BytesIO ou path)
    
    zplm_marc_df = load_xlsx(zplm_marc_file, sheet_name=0) if zplm_marc_file else None
    zplm_mara_df = load_xlsx(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

    if zplm_marc_df is None or zplm_mara_df is None:
        # Sem ZPLM, não enriquece
        return errors_df

    zplm_marc_df.columns = [c.strip().upper() for c in zplm_marc_df.columns]
    zplm_mara_df.columns = [c.strip().upper() for c in zplm_mara_df.columns]
    if not {'MATERIAL', 'PLANT'}.issubset(set(zplm_marc_df.columns)):
        return errors_df
    if not {'MATERIAL', 'ENTERPROJ NUMBER'}.issubset(set(zplm_mara_df.columns)):
        return errors_df

    # 👉 Ler “Original Data” do relatório anterior (UploadedFile/BytesIO/path)
    prev_original_df = None
    try:
        prev_original_df = load_xlsx(latest_report, sheet_name="Original Data")
    except Exception:
        prev_original_df = pd.DataFrame(columns=['Material', 'Enterproj Number'])

    def norm_str(x):
        import pandas as pd
        return '' if pd.isna(x) else str(x).strip()

    prev_enterproj_map = dict(zip(
        prev_original_df.get('Material', pd.Series(dtype=str)).astype(str).str.strip(),
        prev_original_df.get('Enterproj Number', pd.Series(dtype=str)).astype(str).str.strip()
    ))
    zplm_enterproj_map = dict(zip(
        zplm_mara_df['MATERIAL'].astype(str).str.strip(),
        zplm_mara_df['ENTERPROJ NUMBER']  # pode ser numérico/vazio
    ))

    # Preencher PLANT a partir do ZPLM MARC (duplica por plant)
    zplm_marc_filtered = zplm_marc_df[['MATERIAL', 'PLANT']].drop_duplicates()
    subset['Material_norm'] = subset['Material'].astype(str).str.strip()
    zplm_marc_filtered['MATERIAL_norm'] = zplm_marc_filtered['MATERIAL'].astype(str).str.strip()

    merged = subset.merge(
        zplm_marc_filtered[['MATERIAL_norm', 'PLANT']],
        left_on='Material_norm',
        right_on='MATERIAL_norm',
        how='left'
    )
    merged['Plant'] = merged['PLANT'].where(merged['PLANT'].notna(), merged['Plant'])
    merged = merged.drop(columns=['MATERIAL_norm', 'Material_norm', 'PLANT', 'MATERIAL'], errors='ignore')

    # Validação do Enterproj vs 1º relatório
    updated_rows = []
    for _, r in merged.iterrows():
        mat = str(r['Material']).strip()
        prev_val = norm_str(prev_enterproj_map.get(mat, ''))
        zplm_val = norm_str(zplm_enterproj_map.get(mat, ''))

        extras = []
        if zplm_val == '':
            extras.append("Mandatory field blank in ZPLM MARA (Enterproj Number)")
        elif prev_val != '' and zplm_val != prev_val:
            extras.append(f"Invalid value: should equal first report Enterproj Number (expected: '{prev_val}', current: '{zplm_val}')")

        base_err = norm_str(r['Error'])
        if extras:
            r['Error'] = (base_err + (' \n' if base_err else '') + ' \n'.join(extras)).strip()
        updated_rows.append(r)

    enriched = pd.DataFrame(updated_rows)

    # Reconciliar com errors_df original
    base = errors_df.loc[~mask].copy()
    out = pd.concat([base, enriched], ignore_index=True)
    return out

def validar_neg_stocks_in_plant(marc_df, mara_df):
    """
    Regras (apenas MARA/MARC) com filtro por Material Type:
    - Só considerar materiais cujo Material Type ∈ {FERT, HALB, ROH}.
    - Para ROH/HALB: campo obrigatório (não pode estar vazio).
    - Valor 'marcado' (X/1/TRUE/YES/Y) apenas permitido quando Plant ∈ {PT01, PT02} e Material Type ∈ {ROH, HALB}.
    Retorna lista de dicts (erros) para serem concatenados ao errors_df.
    """
    import pandas as pd

    # Requisitos mínimos
    if marc_df is None or marc_df.empty or mara_df is None or mara_df.empty:
        return []
    if 'Material' not in marc_df.columns or 'Plant' not in marc_df.columns:
        return []
    if 'Material' not in mara_df.columns or 'Material Type' not in mara_df.columns:
        return []
    if 'Neg. stocks in plant' not in marc_df.columns:
        return []

    # Mapa Material -> Material Type
    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip()
    ))

    consider_types = {'FERT', 'HALB', 'ROH'}

    def _s(x):
        return '' if pd.isna(x) else str(x).strip()

    erros = []
    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant = _s(row.get('Plant'))
        mt = _s(mt_map.get(material, ''))

        # Ignorar fora do conjunto
        if mt not in consider_types:
            continue

        raw = row.get('Neg. stocks in plant')
        val = _s(raw).upper()

        # variantes de "marcado"
        marcado = val in {'X', '1', 'TRUE', 'YES', 'Y'}

        # 1) Obrigatório para ROH/HALB
        if mt in {'ROH', 'HALB'} and val == '':
            erros.append({
                'Line': idx + 2,  # linha real do MARC.xlsx
                'Field': 'Neg. stocks in plant',
                'Error': "Mandatory field blank (required for Material Type ROH/HALB)",
                'Material': material,
                'Plant': plant
            })

        # 2) Se estiver marcado, tem de respeitar (Plant PT01/PT02 e MT ROH/HALB)
        if marcado and (plant not in {'PT01', 'PT02'} or mt not in {'ROH', 'HALB'}):
            # Nota: para FERT ignoramos (não reportar)
            if mt in {'ROH', 'HALB'}:
                erros.append({
                    'Line': idx + 2,
                    'Field': 'Neg. stocks in plant',
                    'Error': (
                        "Invalid value: marked only allowed for Plant PT01/PT02 and "
                        "Material Type ROH or HALB "
                        f"(Current: Plant='{plant}', Material Type='{mt}')"
                    ),
                    'Material': material,
                    'Plant': plant
                })

    return erros



# ------------------------
# Pós-processamento (Views, Roles, Error_Type, Summary)
# ------------------------


def postprocess_with_views_roles_and_classification(
    errors_df: pd.DataFrame,
    mara_df: pd.DataFrame,
    marc_df: pd.DataFrame,
    roles_df: pd.DataFrame,
    views_df: pd.DataFrame,
    plant_filter: str | None = None,
):
    """
    Pós-processamento completo:
      - mapeia Field -> View (Roles)
      - calcula Letters_Present/Missing por Material/Plant (MARC 'Maintenance status')
      - gera views_missing_df (por par Material/Plant)
      - marca Justified_By_Missing_View e classifica Error_Type (Missing View / Field Error)
      - cria coverage_df (quantos materiais sem cada view requerida)
      - cria view_tabs_dict (abas 'Missing_<View>' com detalhes)
      - constrói contagem_final_df (Adjusted Summary) + Role + View_Description + overrides
      - renomeia 'Export/Import Group' -> 'Intrastat Group' e trata 'Not aplicable'

    Retorna:
      errors_df, errors_filtered_df, views_missing_output_df, contagem_final_df, coverage_df, view_tabs_dict
    """

    import pandas as pd

    # -----------------------------------------------
    # 1) Normalizações de inputs (strip e consistência)
    # -----------------------------------------------
    
    if errors_df is None:
        errors_df = pd.DataFrame()
    if roles_df is None or views_df is None:
        empty = pd.DataFrame()
        return errors_df, empty, empty, empty, empty, {}

    roles_df = roles_df.copy()
    views_df = views_df.copy()
    marc_df = marc_df.copy()
    mara_df = mara_df.copy()
    errors_df = errors_df.copy()

    
    # --- NOVO: normalizar e aplicar filtro de planta logo no início ---
    if plant_filter is not None:
        plant_filter = str(plant_filter).strip().upper()

        # Conjunto de materiais que existem na planta selecionada (via MARC)
        mats_in_plant = set()
        if 'Material' in marc_df.columns and 'Plant' in marc_df.columns:
            mats_in_plant = set(
                marc_df.loc[
                    marc_df['Plant'].astype(str).str.strip().str.upper() == plant_filter,
                    'Material'
                ].astype(str).str.strip()
            )

        # 1) Filtrar MARC (para views missing / coverage / tabs)
        if 'Plant' in marc_df.columns:
            marc_df = marc_df[marc_df['Plant'].astype(str).str.strip().str.upper() == plant_filter].copy()

        # 2) Filtrar errors_df:
        #    - manter linhas com Plant == plant_filter
        #    - manter linhas com Plant vazio SE o Material existir na planta (mats_in_plant)
        if 'Plant' in errors_df.columns and 'Material' in errors_df.columns:
            plant_col = errors_df['Plant'].astype(str).str.strip().str.upper()
            mat_col   = errors_df['Material'].astype(str).str.strip()

            mask_keep = (
                (plant_col == plant_filter) |
                ((plant_col == '') & (mat_col.isin(mats_in_plant)))
            )
            errors_df = errors_df[mask_keep].copy()

    # Strip de cabeçalhos críticos
    roles_df.columns = roles_df.columns.astype(str).str.strip()
    views_df.columns = views_df.columns.astype(str).str.strip()
    marc_df.columns = marc_df.columns.astype(str).str.strip()
    mara_df.columns = mara_df.columns.astype(str).str.strip()
    errors_df.columns = errors_df.columns.astype(str).str.strip()

    # -----------------------------------------------
    # 2) Mapeamentos base: Field -> View ; View -> Descrição
    # -----------------------------------------------
    if 'Field' in roles_df.columns and 'View' in roles_df.columns:
        roles_df['Field'] = roles_df['Field'].astype(str).str.strip()
        roles_df['View']  = roles_df['View'].astype(str).str.strip()
        field_to_view = roles_df.set_index('Field')['View'].to_dict()
    else:
        field_to_view = {}

    if 'Codigo' in views_df.columns and 'Descricao' in views_df.columns:
        views_df['Codigo']    = views_df['Codigo'].astype(str).str.strip()
        views_df['Descricao'] = views_df['Descricao'].astype(str).str.strip()
        view_to_desc = views_df.set_index('Codigo')['Descricao'].to_dict()
        expected_letters = set(views_df['Codigo'].dropna().unique())
    else:
        view_to_desc = {}
        expected_letters = set()

    # Adicionar coluna 'View' aos erros a partir do Field
    if 'Field' in errors_df.columns:
        errors_df['View'] = errors_df['Field'].map(field_to_view).fillna('Not aplicable')
    else:
        errors_df['View'] = 'Not aplicable'

    # -----------------------------------------------
    # 3) Extrair Letters_Present por Material/Plant
    # -----------------------------------------------
    def extract_letters(status):
        if pd.isna(status):
            return []
        return list(str(status).strip())

    # garantir a coluna 'Maintenance status' (já deve vir do ensure_basic_data_k)
    if 'Maintenance status' not in marc_df.columns:
        marc_df['Maintenance status'] = ''

    marc_df['Letters_Present'] = marc_df['Maintenance status'].apply(extract_letters)

    # Mapear MT e PT por Material/Plant
    # MARA: Material -> Material Type
    mt_map = dict(zip(
        mara_df.get('Material', pd.Series(dtype=str)).astype(str).str.strip(),
        mara_df.get('Material Type', pd.Series(dtype=str)).astype(str).str.strip()
    ))
    # MARC: (Material, Plant) -> Plant Type (se existir)
    if 'Plant Type' in marc_df.columns:
        marc_df['_KEY'] = list(zip(
            marc_df['Material'].astype(str).str.strip(),
            marc_df['Plant'].astype(str).str.strip()
        ))
        pt_map = dict(zip(marc_df['_KEY'], marc_df['Plant Type'].astype(str).str.strip()))
    else:
        pt_map = {}

    # -----------------------------------------------
    # 4) Views Missing por (Material, Plant)
    # -----------------------------------------------
    valid_types = {'FERT', 'HALB', 'ROH'}
    # Filtrar materiais com MT válido
    marc_filtered = marc_df[
        marc_df['Material'].astype(str).str.strip().isin(
            mara_df.loc[mara_df['Material Type'].astype(str).str.strip().isin(valid_types),
                       'Material'].astype(str).str.strip().unique()
        )
    ].copy()

    view_to_fields = roles_df.groupby('View')['Field'].apply(list).to_dict()

    missing_rows = []
    for _, row in marc_filtered.iterrows():
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        present  = set(row.get('Letters_Present', []) or [])
        missing  = expected_letters - present

        fields_missing = []
        descriptions_missing = []
        for letter in missing:
            fields_missing.extend(view_to_fields.get(letter, []))
            descriptions_missing.append(view_to_desc.get(letter, ''))

        missing_rows.append({
            'Material': material,
            'Plant': plant,
            'Letters_Present': ','.join(sorted(present)),
            'Letters_Missing': ','.join(sorted(missing)),
            'Fields_Missing': ','.join(sorted(fields_missing)),
            'Descriptions_Missing': ','.join(sorted(descriptions_missing))
        })

    views_missing_df = pd.DataFrame(missing_rows)

    # -----------------------------------------------
    # 5) Classificação por Missing View e Justification
    # -----------------------------------------------
    # helper para obter conjunto de views em falta para (Material, Plant)
    def missing_set_for(mat, plt):
        s = views_missing_df.loc[
            (views_missing_df['Material'] == mat) &
            (views_missing_df['Plant'] == plt),
            'Letters_Missing'
        ]
        if s.empty:
            return set()
        raw = s.iloc[0]
        return set(raw.split(',')) if isinstance(raw, str) and raw else set()

    # Classificar Error_Type
    def classify_error(row):
        # Se a View do erro está em falta para (Material, Plant) → Missing View
        mat = str(row.get('Material', '')).strip()
        plt = str(row.get('Plant', '')).strip()
        view = str(row.get('View', '')).strip()
        miss = missing_set_for(mat, plt)
        if view in miss and view != '' and view != 'Not aplicable':
            return 'Missing View'
        return 'Field Error'

    if not errors_df.empty:
        errors_df['Error_Type'] = errors_df.apply(classify_error, axis=1)
        # Justified_By_Missing_View = Missing View
        errors_df['Justified_By_Missing_View'] = (errors_df['Error_Type'] == 'Missing View')
    else:
        errors_df['Error_Type'] = []
        errors_df['Justified_By_Missing_View'] = []

    # -----------------------------------------------
    # 6) Separar erros por tipo
    # -----------------------------------------------
    errors_filtered_df = errors_df[errors_df['Justified_By_Missing_View']].copy()        # View is Missing
    views_missing_output_df = errors_df[~errors_df['Justified_By_Missing_View']].copy()  # Field With Error

    # -----------------------------------------------
    # 7) Adjusted Summary (contagem apenas de Field Errors)
    # -----------------------------------------------
    if not views_missing_output_df.empty and 'Field' in views_missing_output_df.columns:
        contagem_final_df = (
            views_missing_output_df['Field'].value_counts()
            .rename_axis('Field')
            .reset_index(name='Adjusted Nº of Errors')
        )
    else:
        contagem_final_df = pd.DataFrame(columns=['Field', 'Adjusted Nº of Errors'])

    # Mapear View e Descrição, com fallback
    if not contagem_final_df.empty:
        contagem_final_df['View'] = contagem_final_df['Field'].map(field_to_view).fillna('Not aplicable')
        contagem_final_df['View_Description'] = contagem_final_df['View'].map(view_to_desc).fillna('Not aplicable')
    else:
        contagem_final_df['View'] = []
        contagem_final_df['View_Description'] = []

    # Role por Field (um-para-um)
    if 'Role' in roles_df.columns:
        roles_df['Role']  = roles_df['Role'].astype(str).str.strip()
        roles_df['Field'] = roles_df['Field'].astype(str).str.strip()
        field_to_role = roles_df.set_index('Field')['Role'].to_dict()
    else:
        field_to_role = {}

    if not contagem_final_df.empty:
        contagem_final_df['Role'] = contagem_final_df['Field'].map(field_to_role).fillna('Not aplicable')

    # Override específico da View_Description para Enterproj Number e Plant Type
    if not contagem_final_df.empty:
        mask_override = contagem_final_df['Field'].isin(['Enterproj Number', 'Plant Type'])
        contagem_final_df.loc[mask_override, 'View_Description'] = 'Product Master Visteon Data'

    # Renomeação de rótulos de Fields para os relatórios
    field_renames = {
        'Export/Import Group': 'Intrastat Group'
    }

    def apply_field_renames(df_):
        df_ = df_.copy()
        if 'Field' in df_.columns:
            df_['Field'] = df_['Field'].replace(field_renames)
        return df_

    errors_df              = apply_field_renames(errors_df)
    errors_filtered_df     = apply_field_renames(errors_filtered_df)
    views_missing_output_df= apply_field_renames(views_missing_output_df)
    contagem_final_df      = apply_field_renames(contagem_final_df)

    # Limpeza de textos literais: Not aplicable
    for col in ['View', 'View_Description', 'Role']:
        if col in contagem_final_df.columns:
            contagem_final_df[col] = contagem_final_df[col].astype(str).replace(
                {'nan': 'Not aplicable', 'None': 'Not aplicable', '': 'Not aplicable'}
            )

    # -----------------------------------------------
    # 8) Coverage por View (quantos materiais sem cada view requerida)
    #    → usa field_rules para decidir se uma view é requerida para MT/PT
    # -----------------------------------------------
    field_rules = {
        'Production Scheduling Profile': lambda mt, pt: mt in ['FERT', 'HALB'],
        'Repetitive Manufacturing Profile': lambda mt, pt: mt in ['FERT', 'HALB'],
        'Enterproj Number': lambda mt, pt: mt == 'FERT',
        'Procurement Type': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'] and pt in ['PP', 'MP'],
        'Account Assignment Group': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Automatic Purchase Order Allowed': lambda mt, pt: pt == 'PP',
        'Batch Management Required Ind.': lambda mt, pt: True,
        'Source List Required': lambda mt, pt: pt == 'PP',
        'Purchasing Group': lambda mt, pt: pt == 'PP',
        'MRP Type': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Planning Time Fence': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Lot Sizing Procedure': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'With Qty Structure': lambda mt, pt: mt in ['FERT', 'HALB'],
        'Item Category Group': lambda mt, pt: True,
        'General Item Category Group': lambda mt, pt: True,
        'Loading Group': lambda mt, pt: True,
        'Division': lambda mt, pt: True,
        'Export/Import Group': lambda mt, pt: True,
        'Product Hierarchy': lambda mt, pt: mt == 'FERT',
        'Price Control': lambda mt, pt: True,
        'Price Unit': lambda mt, pt: True,
        'Valuation Class': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'Gross Weight': lambda mt, pt: pt == 'FP',
        'Net Weight': lambda mt, pt: pt == 'FP',
        'Planning Calendar': lambda mt, pt: True,
        'Period Type': lambda mt, pt: True,
        'Strategy Group': lambda mt, pt: True,
        'Inspection Type': lambda mt, pt: mt in ['FERT', 'HALB', 'ROH'],
        'MRP Controller': lambda mt, pt: True,
        'Plant Specific Material Status': lambda mt, pt: True,
        'Backflush': lambda mt, pt: True,
        'Product-Related Origin': lambda mt, pt: True,
        'Availability Check': lambda mt, pt: True,
        'Neg. stocks in plant': lambda mt, pt: mt in ['ROH', 'HALB'],
    }

    def is_view_required(material_type, plant_type, view_code):
        fields = view_to_fields.get(view_code, [])
        for field in fields:
            rule = field_rules.get(field)
            if rule and rule(material_type, plant_type):
                return True
        return False

    # Contagem por view
    view_missing_count = {letter: 0 for letter in expected_letters}
    for _, r in marc_filtered.iterrows():
        mt  = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
        pt  = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper() if pt_map else ''
        present = set(r.get('Letters_Present', []) or [])
        for v in expected_letters:
            if v not in present and is_view_required(mt, pt, v):
                view_missing_count[v] += 1

    coverage_data = []
    for letter, count in view_missing_count.items():
        coverage_data.append({
            'View': letter,
            'View_Description': view_to_desc.get(letter, ''),
            'Materiais without View': count,
            'Link': f'=HYPERLINK("#Missing_{letter}!A1","Check Details")'
        })
    coverage_df = pd.DataFrame(coverage_data)

    # -----------------------------------------------
    # 9) Abas Missing_<View> com detalhe por material/planta
    # -----------------------------------------------
    view_tabs_dict = {}
    for v in expected_letters:
        rows = []
        for _, r in marc_filtered.iterrows():
            mt  = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
            pt  = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper() if pt_map else ''
            present = set(r.get('Letters_Present', []) or [])
            if v not in present and is_view_required(mt, pt, v):
                rows.append({
                    'Material': str(r['Material']).strip(),
                    'Plant': str(r['Plant']).strip(),
                    'Missing_View': v,
                    'View_Description': view_to_desc.get(v, ''),
                    'Fields_Expected': ','.join(view_to_fields.get(v, []))
                })
        view_tabs_dict[v] = pd.DataFrame(rows)

    # -----------------------------------------------
    # 10) Saída
    # -----------------------------------------------
    return errors_df, errors_filtered_df, views_missing_output_df, contagem_final_df, coverage_df, view_tabs_dict

# ==========================================================
# 3) FUNÇÃO PRINCIPAL — ORQUESTRAÇÃO COMPLETA
# ==========================================================

def run_full_validation(
    sap_file,
    mara_file,
    marc_file,
    mvke_file,
    roles_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter: str | None = None,
):
    """
    Executa o processo completo:
      - Carregamento
      - Normalização
      - Validações
      - Pós-processamento Views/Role
      - Geração do Excel final em memória
      - Devolve (excel_bytes, errors_df, summary_df)
    """

    # ------------------
    # Carregar DataFrames
    # ------------------

    import os

    sap_df = load_xlsx(sap_file, sheet_name="SAPUI5 Export")
    mara_df = load_xlsx(mara_file, sheet_name=0)
    marc_df = load_xlsx(marc_file, sheet_name=0)
    mvke_df = load_xlsx(mvke_file, sheet_name=0)

    # Fallback: se não houver roles_file, usa o Excel embutido via Base64
    if roles_file is None:
        roles_file = get_rolesviews_bytesio()  # BytesIO

    rolesviews_dict = load_xlsx(roles_file, sheet_name=None)

    zplm_marc_df = load_xlsx(zplm_marc_file, sheet_name=0) if zplm_marc_file else None
    zplm_mara_df = load_xlsx(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

    # ------------------
    # Verificações base
    # ------------------
    if sap_df is None:
        raise ValueError("SAPUI5 Export.xlsx não carregado ou sem folha 'SAPUI5 Export'")
    if mara_df is None:
        raise ValueError("MARA.xlsx não carregado.")
    if marc_df is None:
        raise ValueError("MARC.xlsx não carregado.")
    if mvke_df is None:
        raise ValueError("MVKE.xlsx não carregado.")
    
    # Garantir que o RolesViews embutido tem as duas folhas
    if "Roles" not in rolesviews_dict or "Views" not in rolesviews_dict:
        raise ValueError("RolesViews.xlsx (embutido) deve conter as folhas 'Roles' e 'Views'.")

    roles_df = rolesviews_dict["Roles"]
    views_df = rolesviews_dict["Views"]

    # --- NOVO: aplicar filtro de planta ao SAPUI5 Export (Original Data) ---
    if plant_filter is not None and 'Plant' in sap_df.columns:
        pf = str(plant_filter).strip().upper()
        sap_df = sap_df[sap_df['Plant'].astype(str).str.strip().str.upper() == pf].copy()


    # ------------------
    # Pré-normalização
    # ------------------
    marc_df = ensure_basic_data_k(marc_df)

    # ------------------
    # Execução das Validações
    # ------------------
    resultados = []

    # Campos obrigatórios
    for coluna in ['Material', 'Material Type', 'Description', 'Base UOM', 'Plant', 'Plant Type']:
        validar_campo_obrigatorio(sap_df, coluna, resultados)

    # Chamadas em sequência
    validar_procurement_type(sap_df, resultados)
    validar_account_assignment_group(sap_df, resultados, mvke_df)
    validar_automatic_purchase_order(sap_df, resultados)
    validar_batch_management(sap_df, resultados)
    validar_consumption_mode(sap_df, resultados)
    validar_division(sap_df, resultados)
    validar_export_import_group(sap_df, resultados)
    validar_general_item_category_group(sap_df, resultados)
    validar_gross_weight(sap_df, resultados)
    validar_repetitive_manufacturing(sap_df, resultados)
    validar_item_category_group(sap_df, resultados)
    validar_jit_delivery_sched(sap_df, resultados)
    validar_loading_group(sap_df, resultados)
    validar_lot_sizing_procedure(sap_df, resultados)
    validar_material_group(sap_df, resultados)
    validar_net_weight(sap_df, resultados)
    validar_period_type(sap_df, resultados)
    validar_planning_calendar(sap_df, resultados)
    validar_planning_time_fence(sap_df, resultados)
    validar_mrp_type(sap_df, resultados, zplm_marc_df=zplm_marc_df)
    validar_price_control(sap_df, resultados)
    validar_price_unit(sap_df, resultados)
    validar_product_hierarchy(sap_df, resultados)
    validar_production_scheduling_profile(sap_df, resultados)
    validar_profit_center(sap_df, resultados)
    validar_purchasing_group(sap_df, resultados)
    validar_repetitive_manufacturing_profile(sap_df, resultados)
    validar_source_list_required(sap_df, resultados)
    validar_storage_location(sap_df, resultados)
    validar_valuation_class(sap_df, resultados)
    validar_with_qty_structure(sap_df, resultados)
    validar_enterproj_number(sap_df, resultados)

    validar_inspection_type(sap_df, resultados)
    validar_mrp_controller(sap_df, resultados)
    validar_plant_specific_material_status(sap_df, resultados)
    validar_backflush(sap_df, resultados)
    validar_strategy_group(sap_df, resultados)
    validar_product_related_origin(sap_df, resultados)
    validar_availability_check(sap_df, resultados)

    # ------------------
    # Criar DataFrame de erros
    # ------------------
    errors_df = pd.DataFrame(resultados)

    # Enriquecimento opcional (se tiveres previous_report)
    
    if previous_report:
        errors_df = fill_plant_and_validate_enterproj_from_zplm(
            errors_df,
            latest_report=previous_report,     # UploadedFile/BytesIO/path do relatório anterior
            zplm_marc_file=zplm_marc_file,     # passa o próprio upload
            zplm_mara_file=zplm_mara_file      # idem
        )

    # ------------------
    # Enriquecimentos MARA / MARC / ZPLM
    # ------------------
    if not errors_df.empty:
        errors_df = add_old_material_number_errors(mara_df, errors_df)
        neg_errors = validar_neg_stocks_in_plant(marc_df, mara_df)
        if neg_errors:
            errors_df = pd.concat([errors_df, pd.DataFrame(neg_errors)], ignore_index=True)

    # ------------------
    # Pós-processamento: Views / Roles / Summary ajustado
    # ------------------
    (
        errors_df,
        errors_filtered_df,
        views_missing_output_df,
        contagem_final_df,
        coverage_df,
        view_tabs_dict
    ) = postprocess_with_views_roles_and_classification(
        errors_df,
        mara_df,
        marc_df,
        roles_df,
        views_df,
        plant_filter=plant_filter,
    )

    # ------------------
    # Geração do Excel final em memória
    # ------------------
    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M")
    file_bytes = io.BytesIO()

    with pd.ExcelWriter(file_bytes, engine="openpyxl") as writer:
        sap_df.to_excel(writer, sheet_name="Original Data", index=False)
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        errors_filtered_df.to_excel(writer, sheet_name="View is Missing", index=False)
        views_missing_output_df.to_excel(writer, sheet_name="Field With Error", index=False)

        # Summary + Coverage
        base_row = 1
        if not contagem_final_df.empty:
            contagem_final_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
            base_row = len(contagem_final_df) + 5

        if not coverage_df.empty:
            coverage_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=base_row)

        # Abas Missing_<View>
        for view_code, df_view in view_tabs_dict.items():
            sheet = f"Missing_{view_code}"[:31]  # Excel limita a 31 chars
            df_view.to_excel(writer, sheet_name=sheet, index=False)

    file_bytes.seek(0)

    # ------------------
    # Summary para exibir no Streamlit
    # ------------------
    if 'Field' in errors_df.columns and not errors_df.empty:
        summary_df = (
            errors_df['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Adjusted Nº of Errors')
        )
    else:
        summary_df = pd.DataFrame({"Field": [], "Adjusted Nº of Errors": []})

    # ------------------
    # Retorno final
    # ------------------
    return file_bytes, errors_df, summary_df

# ==========================================================
# ATL / PLM WORKFLOW COMBINED (a partir do SAP_ATL_Validation.py)
# ==========================================================

def _apply_conditional_formatting_in_memory(xlsx_bytes: io.BytesIO, sheet_names=("PT01", "PT02")) -> io.BytesIO:
    """
    Abre o Excel em memória e aplica formatação condicional na coluna 'Aging'
    das folhas indicadas. Devolve novamente os bytes prontos a download.
    """
    xlsx_bytes.seek(0)
    wb = load_workbook(xlsx_bytes)
    for sheet in sheet_names:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        # Encontrar a coluna 'Aging'
        aging_col_idx = None
        for idx, cell in enumerate(ws[1], start=1):
            if cell.value == "Aging":
                aging_col_idx = idx
                break
        if aging_col_idx is None:
            continue
        green  = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        red    = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        for row in ws.iter_rows(min_row=2, min_col=aging_col_idx, max_col=aging_col_idx):
            for cell in row:
                try:
                    val = int(cell.value)
                except (TypeError, ValueError):
                    val = None
                if val is None:
                    continue
                if val <= 30:
                    cell.fill = green
                elif 31 <= val <= 90:
                    cell.fill = yellow
                elif 91 <= val <= 180:
                    cell.fill = orange
                else:
                    cell.fill = red
    # Regravar para bytes
    out_bytes = io.BytesIO()
    wb.save(out_bytes)
    out_bytes.seek(0)
    return out_bytes

def run_plm_workflow_combined(auto_file, plm_file, timing_file):
    """
    Implementa o ATL (Automotive + TimingLetter + Timing Letters mapping),
    gera o Excel em memória com abas PT01/PT02 e devolve:
      - excel_bytes (BytesIO)
      - df_pt01 (DataFrame)
      - df_pt02 (DataFrame)
    """
    # --- Ler fontes com o loader universal já existente no projeto ---
    df_auto   = load_xlsx(auto_file, sheet_name=0)      # "PLM Automotive Validation.xlsx"
    df_plm    = load_xlsx(plm_file, sheet_name=0)       # "PLM TimingLetter Validation.xlsx"
    df_timing = load_xlsx(timing_file, sheet_name=0)    # "Timing Letters.xlsx"

    if df_auto is None or df_plm is None or df_timing is None:
        raise ValueError("Ficheiros ATL em falta: Automotive/TimingLetter/Timing Letters")

    # 2) Validar colunas obrigatórias (igual ao script)
    required_base_cols = ['Plant', 'Approval Status', 'Trigger ID', 'Creation Date']
    for col in required_base_cols:
        if col not in df_auto.columns:
            raise ValueError(f"Coluna '{col}' não existe no ficheiro Automotive.")
        if col not in df_plm.columns:
            raise ValueError(f"Coluna '{col}' não existe no ficheiro TimingLetter PLM.")
    if 'Timing Letter Number' not in df_timing.columns or 'Change Number' not in df_timing.columns:
        raise ValueError("Ficheiro Timing Letters deve conter 'Timing Letter Number' e 'Change Number'.")

    # 3) MAPEAR Trigger ID -> Change Number (TimingLetter)
    map_trigger_to_change = dict(zip(df_timing['Timing Letter Number'], df_timing['Change Number']))
    df_plm['Trigger ID'] = df_plm['Trigger ID'].apply(lambda x: map_trigger_to_change.get(x, x))

    # 4) Função 'Pending(Claimed)' — igual à do script ATL
    def add_pending_claimed(df):
        actual_end_cols = [col for col in df.columns if col.startswith("Actual End Date of Work Item")]
        df['Pending(Claimed)'] = ""
        for idx, row in df.iterrows():
            approvers_pending = []
            for col in actual_end_cols:
                date_val = row[col] if col in df.columns else None
                if pd.notna(date_val):
                    s = str(date_val)
                    if s.startswith("9999-12-31") or s.startswith("12/31/9999"):
                        col_idx = df.columns.get_loc(col)
                        if col_idx > 0:
                            approver_name = row[df.columns[col_idx - 1]]
                            if pd.notna(approver_name):
                                approvers_pending.append(str(approver_name))
            if approvers_pending:
                df.at[idx, 'Pending(Claimed)'] = ", ".join(approvers_pending)
        return df

    df_auto = add_pending_claimed(df_auto)
    df_plm  = add_pending_claimed(df_plm)

    # 5) Preparar Automotive
    approver_cols_auto = ['Approver1'] + [f'approver{i}' for i in range(2, 8)] + ['Approver8']
    categories_auto = ["Pricing", "Purchasing", "MPL", "Customs", "Sales", "Manufacturing", "Costing", "MDS"]
    df_auto_f = df_auto[df_auto['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    df_auto_f['Creation Date'] = pd.to_datetime(df_auto_f['Creation Date'], errors='coerce')
    df_auto_f.sort_values('Creation Date', ascending=False, inplace=True)
    df_auto_f.drop_duplicates(subset=['Trigger ID'], keep='first', inplace=True)
    for cat in categories_auto:
        df_auto_f[cat] = ""
    for idx, row in df_auto_f.iterrows():
        for col in approver_cols_auto:
            if col in df_auto_f.columns:
                value = str(row[col]) if pd.notna(row[col]) else ""
                clean = value.split("(")[0].strip()
                if clean in categories_auto:
                    df_auto_f.at[idx, clean] = "X"
    df_auto_f['Creation Date'] = df_auto_f['Creation Date'].dt.strftime('%Y-%m-%d')
    df_auto_f['Source'] = "Automotive"

    # 6) Preparar TimingLetter (MPL - TL, MDS - TL)
    df_tl_f = df_plm[df_plm['Approval Status'].isin(['INAPPROVAL', 'REJECTED'])].copy()
    df_tl_f['Creation Date'] = pd.to_datetime(df_tl_f['Creation Date'], errors='coerce')
    df_tl_f.sort_values('Creation Date', ascending=False, inplace=True)
    df_tl_f.drop_duplicates(subset=['Trigger ID'], keep='first', inplace=True)
    df_tl_f['MPL - TL'] = ""
    df_tl_f['MDS - TL'] = ""
    approver_cols_tl = [c for c in df_plm.columns if "approver" in c.lower()]
    for idx, row in df_tl_f.iterrows():
        for col in approver_cols_tl:
            if col in df_tl_f.columns:
                val = str(row[col]) if pd.notna(row[col]) else ""
                if val.startswith('MPL'):
                    df_tl_f.at[idx, 'MPL - TL'] = "X"
                if val.startswith('MDS'):
                    df_tl_f.at[idx, 'MDS - TL'] = "X"
    df_tl_f['Creation Date'] = df_tl_f['Creation Date'].dt.strftime('%Y-%m-%d')
    df_tl_f['Source'] = "TimingLetter"

    base_cols  = ['Source', 'Trigger ID', 'Plant', 'Creation Date', 'Approval Status', 'Pending(Claimed)']
    auto_cols  = base_cols + ["MPL - TL", "MDS - TL", "Pricing", "Purchasing", "MPL", "Customs",
                              "Sales", "Manufacturing", "Costing", "MDS"]
    # Preencher colunas vazias para alinhamento
    for col in ["Pricing", "Purchasing", "MPL", "Customs", "Sales", "Manufacturing", "Costing", "MDS"]:
        if col not in df_tl_f.columns:
            df_tl_f[col] = ""
    for col in ["MPL - TL", "MDS - TL"]:
        if col not in df_auto_f.columns:
            df_auto_f[col] = ""

    df_auto_f = df_auto_f[auto_cols]
    df_tl_f   = df_tl_f[auto_cols]

    # 7) Concatenar e ordenar
    df_final = pd.concat([df_auto_f, df_tl_f], ignore_index=True)
    df_final['Creation Date'] = pd.to_datetime(df_final['Creation Date'], errors='coerce')
    df_final['Source'] = pd.Categorical(df_final['Source'], categories=["Automotive", "TimingLetter"], ordered=True)
    df_final.sort_values(by=["Source", "Creation Date"], ascending=[True, False], inplace=True)

    # 8) Aging + formatar datas
    today = datetime.now()
    df_final['Aging'] = (today - df_final['Creation Date']).dt.days
    df_final['Creation Date'] = df_final['Creation Date'].dt.strftime('%Y-%m-%d')

    # 9) Ordem final das colunas
    final_cols = [
        "Source",
        "Trigger ID", "Plant", "Creation Date", "Approval Status", "Pending(Claimed)",
        "MPL - TL", "MDS - TL",
        "Pricing", "Purchasing", "MPL", "Customs", "Sales", "Manufacturing", "Costing", "MDS",
        "Aging"
    ]
    df_final = df_final[final_cols]

    # 10) Filtrar PT01/PT02
    df_pt01 = df_final[df_final['Plant'] == 'PT01'].copy()
    df_pt02 = df_final[df_final['Plant'] == 'PT02'].copy()

    # 11) Escrever para Excel (memória) e aplicar formatação condicional
    xlsx_bytes = io.BytesIO()
    with pd.ExcelWriter(xlsx_bytes, engine="openpyxl") as writer:
        df_pt01.to_excel(writer, sheet_name="PT01", index=False)
        df_pt02.to_excel(writer, sheet_name="PT02", index=False)
    xlsx_bytes.seek(0)
    xlsx_bytes = _apply_conditional_formatting_in_memory(xlsx_bytes, sheet_names=("PT01", "PT02"))

    return xlsx_bytes, df_pt01, df_pt02


# =============== DIFF ENGINE (Comparador) ===============
import pandas as pd
import io

def _norm_str(x) -> str:
    """Normaliza para string sem espaços; NaN -> ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ''
    return str(x).strip()

def _load_rolesviews_maps() -> tuple[dict, dict]:
    """
    Carrega o Excel embutido RolesViews.xlsx (via Base64) e devolve:
      - field_to_view: Field -> View (letra)
      - view_to_desc:  View (letra) -> Descrição
    """
    try:
        rolesviews_dict = load_xlsx(get_rolesviews_bytesio(), sheet_name=None)
        roles_df = rolesviews_dict.get("Roles", pd.DataFrame())
        views_df = rolesviews_dict.get("Views", pd.DataFrame())
        if not roles_df.empty:
            roles_df['Field'] = roles_df['Field'].astype(str).str.strip()
            roles_df['View']  = roles_df['View'].astype(str).str.strip()
            field_to_view = roles_df.set_index('Field')['View'].to_dict()
        else:
            field_to_view = {}
        if not views_df.empty:
            views_df['Codigo']    = views_df['Codigo'].astype(str).str.strip()
            views_df['Descricao'] = views_df['Descricao'].astype(str).str.strip()
            view_to_desc = views_df.set_index('Codigo')['Descricao'].to_dict()
        else:
            view_to_desc = {}
        return field_to_view, view_to_desc
    except Exception:
        return {}, {}

def _extract_missing_views_from_workbook(book_dict: dict[str, pd.DataFrame]) -> dict[tuple[str,str], set[str]]:
    """
    Lê todas as folhas 'Missing_<View>' e devolve um mapa:
        (Material, Plant) -> {views_missing}
    Se não existir nenhuma 'Missing_*', tenta inferir a partir de 'All Errors'
    usando a coluna 'Justified_By_Missing_View' = True e 'View' (letra).
    """
    missing_map: dict[tuple[str, str], set[str]] = {}
    # 1) Preferência: folhas Missing_<View>
    missing_sheets = [name for name in book_dict.keys() if name.startswith("Missing_")]
    if missing_sheets:
        for sheet in missing_sheets:
            df = book_dict.get(sheet, pd.DataFrame())
            if df.empty:
                continue
            view_code = sheet.replace("Missing_", "").strip()
            for _, r in df.iterrows():
                mat = _norm_str(r.get('Material'))
                plant = _norm_str(r.get('Plant'))
                if not mat:
                    continue
                key = (mat, plant)
                missing_map.setdefault(key, set()).add(view_code)
        return missing_map

    # 2) Fallback: Inferir a partir de 'All Errors'
    df_err = book_dict.get("All Errors", pd.DataFrame())
    if not df_err.empty and {'Material', 'Plant', 'View', 'Justified_By_Missing_View'}.issubset(df_err.columns):
        for _, r in df_err.iterrows():
            if bool(r.get('Justified_By_Missing_View', False)):
                mat = _norm_str(r.get('Material'))
                plant = _norm_str(r.get('Plant'))
                view_code = _norm_str(r.get('View'))
                if not mat or not view_code or view_code in ('', 'Not aplicable'):
                    continue
                key = (mat, plant)
                missing_map.setdefault(key, set()).add(view_code)
    return missing_map


def compare_adjusted_reports(current_file, previous_file):
    """
    Compara dois relatórios ajustados gerados por run_full_validation e devolve:
      - excel_bytes_diff: BytesIO com folhas:
            'New Errors', 'Resolved Errors', 'Changed Errors',
            'New Materials', 'Views Lost'
      - diff_dfs: dict com os DataFrames acima (para mostrar no Streamlit)

    Keys usadas:
      - 'New/Resolved Errors': (Material, Plant, Field, Error) exata
      - 'Changed Errors':      (Material, Plant, Field) igual, Error diferente
      - 'New Materials':       materiais em 'Original Data' atuais que não existiam no anterior
      - 'Views Lost':          views que passaram a estar em falta no atual, mas não estavam no anterior
    """
    # 1) Carregar ambos os workbooks para dict de DataFrames
    current_dict  = load_xlsx(current_file, sheet_name=None)
    previous_dict = load_xlsx(previous_file, sheet_name=None)

    # 2) Obter DataFrames base
    curr_err = current_dict.get("All Errors", pd.DataFrame()).copy()
    prev_err = previous_dict.get("All Errors", pd.DataFrame()).copy()

    curr_orig = current_dict.get("Original Data", pd.DataFrame()).copy()
    prev_orig = previous_dict.get("Original Data", pd.DataFrame()).copy()

    # Normalização light
    for df in [curr_err, prev_err, curr_orig, prev_orig]:
        if not df.empty:
            for col in df.columns:
                # Convert to string for key columns used later (avoid NaNs mismatch)
                if col in {'Material', 'Plant', 'Field', 'Error'}:
                    df[col] = df[col].astype(str).str.strip()

    # 3) NEW / RESOLVED ERRORS pela chave completa (Material, Plant, Field, Error)
    def _key4(df: pd.DataFrame):
        if df.empty:
            return set()
        cols = ['Material', 'Plant', 'Field', 'Error']
        for c in cols:
            if c not in df.columns:
                df[c] = ''
        return set(tuple(row[c] for c in cols) for _, row in df[cols].iterrows())

    curr_set4 = _key4(curr_err)
    prev_set4 = _key4(prev_err)

    new_errors_tuples     = curr_set4 - prev_set4
    resolved_errors_tuples= prev_set4 - curr_set4

    df_new_errors = pd.DataFrame(list(new_errors_tuples), columns=['Material', 'Plant', 'Field', 'Error'])
    df_resolved_errors = pd.DataFrame(list(resolved_errors_tuples), columns=['Material', 'Plant', 'Field', 'Error'])

    # 4) CHANGED ERRORS: mesmo (Material, Plant, Field), mas Error mudou
    def _key3_map(df: pd.DataFrame) -> dict[tuple[str,str,str], set[str]]:
        if df.empty:
            return {}
        cols = ['Material', 'Plant', 'Field', 'Error']
        for c in cols:
            if c not in df.columns:
                df[c] = ''
        m: dict[tuple[str,str,str], set[str]] = {}
        for _, row in df[cols].iterrows():
            key3 = (row['Material'], row['Plant'], row['Field'])
            m.setdefault(key3, set()).add(row['Error'])
        return m

    curr_map3 = _key3_map(curr_err)
    prev_map3 = _key3_map(prev_err)

    changed_rows = []
    for key3 in set(curr_map3.keys()).intersection(set(prev_map3.keys())):
        if curr_map3[key3] != prev_map3[key3]:
            # houve alteração de texto de erro para o mesmo Material/Plant/Field
            mat, plant, field = key3
            changed_rows.append({
                'Material': mat,
                'Plant': plant,
                'Field': field,
                'Prev_Errors': "; ".join(sorted(prev_map3[key3])),
                'Curr_Errors': "; ".join(sorted(curr_map3[key3])),
            })
    df_changed_errors = pd.DataFrame(changed_rows)

    # 5) NEW MATERIALS
    def _materials(df: pd.DataFrame) -> set[str]:
        if df.empty or 'Material' not in df.columns:
            return set()
        return set(df['Material'].astype(str).str.strip().unique())

    curr_mats = _materials(curr_orig)
    prev_mats = _materials(prev_orig)
    new_mats  = sorted(curr_mats - prev_mats)
    df_new_materials = pd.DataFrame({'Material': new_mats})

    # 6) VIEWS LOST (passaram a estar em falta no atual)
    curr_missing_map = _extract_missing_views_from_workbook(current_dict)
    prev_missing_map = _extract_missing_views_from_workbook(previous_dict)

    # Mapa para fácil lookup
    all_keys = set(curr_missing_map.keys()).union(set(prev_missing_map.keys()))
    field_to_view, view_to_desc = _load_rolesviews_maps()

    views_lost_rows = []
    for key in all_keys:
        curr_missing = curr_missing_map.get(key, set())
        prev_missing = prev_missing_map.get(key, set())
        # Views novas em falta agora (presentes no curr_missing mas NÃO no prev_missing)
        lost = sorted(curr_missing - prev_missing)
        if not lost:
            continue
        mat, plant = key
        for v in lost:
            views_lost_rows.append({
                'Material': mat,
                'Plant': plant,
                'View_Lost': v,
                'View_Description': view_to_desc.get(v, ''),
            })
    df_views_lost = pd.DataFrame(views_lost_rows)

    # 7) Construir Excel com estes DataFrames
    out_bytes = io.BytesIO()
    with pd.ExcelWriter(out_bytes, engine="openpyxl") as writer:
        # Ordem de abas
        df_new_errors.to_excel(writer,      sheet_name="New Errors",      index=False)
        df_resolved_errors.to_excel(writer, sheet_name="Resolved Errors", index=False)
        df_changed_errors.to_excel(writer,  sheet_name="Changed Errors",  index=False)
        df_new_materials.to_excel(writer,   sheet_name="New Materials",   index=False)
        df_views_lost.to_excel(writer,      sheet_name="Views Lost",      index=False)
    out_bytes.seek(0)

    # 8) Devolver ambos (bytes + dict com dfs)
    diff_dfs = {
        "New Errors": df_new_errors,
        "Resolved Errors": df_resolved_errors,
        "Changed Errors": df_changed_errors,
        "New Materials": df_new_materials,
        "Views Lost": df_views_lost,
    }
    return out_bytes, diff_dfs


# ==========================
# HISTÓRICO DE EXECUÇÕES / LOGS
# ==========================
import os, json
from datetime import datetime

def _ensure_dir(path: str):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass

def _safe_name(uploaded, default_name: str):
    """Extrai o nome do ficheiro de UploadedFile/BytesIO/path; fallback para default_name."""
    try:
        if hasattr(uploaded, "name") and uploaded.name:
            return str(uploaded.name)
        if isinstance(uploaded, str):
            return os.path.basename(uploaded)
    except Exception:
        pass
    return default_name

def _ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# --------- Construção de payloads por tipo de run ---------
def build_sap_log_payload(*, user: str, plant_filter: str | None,
                          sap_file, mara_file, marc_file, mvke_file,
                          errors_df, summary_df):
    return {
        "run_type": "SAP_VALIDATION",
        "timestamp": _ts(),
        "user": user,
        "plant_filter": plant_filter or "ALL",
        "files": {
            "SAP": _safe_name(sap_file, "SAPUI5 Export.xlsx"),
            "MARA": _safe_name(mara_file, "MARA.xlsx"),
            "MARC": _safe_name(marc_file, "MARC.xlsx"),
            "MVKE": _safe_name(mvke_file, "MVKE.xlsx"),
        },
        "metrics": {
            "total_errors": int(errors_df.shape[0]) if errors_df is not None else 0,
            "top_fields": (
                summary_df.head(5).to_dict(orient="records") if summary_df is not None else []
            ),
        },
    }

def build_atl_log_payload(*, user: str, plant_scope: str,
                          df_pt01: pd.DataFrame, df_pt02: pd.DataFrame):
    return {
        "run_type": "PLM_WORKFLOW_ATL",
        "timestamp": _ts(),
        "user": user,
        "plant_scope": plant_scope,
        "metrics": {
            "PT01_rows": int(df_pt01.shape[0]) if df_pt01 is not None else 0,
            "PT02_rows": int(df_pt02.shape[0]) if df_pt02 is not None else 0,
        },
    }

def build_progress_log_payload(*, user: str,
                               current_name: str, previous_name: str,
                               diff_dfs: dict):
    def _n(name): 
        df = diff_dfs.get(name)
        return int(df.shape[0]) if df is not None else 0
    return {
        "run_type": "SAP_VALIDATION_PROGRESS",
        "timestamp": _ts(),
        "user": user,
        "files": {
            "current": current_name,
            "previous": previous_name,
        },
        "metrics": {
            "new_errors": _n("New Errors"),
            "resolved_errors": _n("Resolved Errors"),
            "changed_errors": _n("Changed Errors"),
            "new_materials": _n("New Materials"),
            "views_lost": _n("Views Lost"),
        },
    }

def build_mrp_log_payload(
    *, user: str, plant_filter: str | None, boolean_col: str | None,
    errors_df: pd.DataFrame, marc_rows: int
):
    return {
        "run_type": "MRP_VALIDATION",
        "timestamp": _ts(),
        "user": user,
        "plant_filter": plant_filter or "ALL",
        "boolean_col": boolean_col or "",
        "metrics": {
            "total_errors": int(errors_df.shape[0]) if errors_df is not None else 0,
            "rows_validated": int(marc_rows),
        },
    }

# --------- Persistência ---------
def save_run_log(payload: dict, *, base_dir: str = ".history"):
    """Guarda o payload JSON numa subpasta por tipo de run."""
    run_type = payload.get("run_type", "UNKNOWN")
    subdir = {
        "SAP_VALIDATION": "sap",
        "PLM_WORKFLOW_ATL": "atl",
        "SAP_VALIDATION_PROGRESS": "progress",
        "MRP_VALIDATION": "mrp", 
    }.get(run_type, "misc")

    out_dir = os.path.join(base_dir, subdir)
    _ensure_dir(out_dir)
    ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(out_dir, f"{run_type}_{ts_file}.json")

    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return out_path
    except Exception:
        return None

def list_run_logs(run_type: str, *, base_dir: str = ".history", limit: int = 25):
    """Lista os últimos 'limit' logs para um run_type: 'sap' | 'atl' | 'progress'."""
    subdir = {"sap": "sap", "atl": "atl", "progress": "progress"}.get(run_type, "")
    if not subdir:
        return []
    folder = os.path.join(base_dir, subdir)
    if not os.path.isdir(folder):
        return []

    files = sorted(
        (os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".json")),
        reverse=True
    )
    out = []
    for p in files[:limit]:
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
                data["_path"] = p  # útil para debugging
                out.append(data)
        except Exception:
            continue
    return out

# ==========================================================
#  MRP VALIDATION (MARC + MARA + ZPLM MARC)
# ==========================================================
def run_mrp_validation(
    mara_file,
    marc_file,
    zplm_marc_file=None,
    *,
    plant_filter: str | None = None,
):
    """
    Constrói a página 'MRP Validation' com regras:
    - Advanced Planning:
        * HALB com Transfer e FERT  -> 'true'
        * HALB sem Transfer (interno) e ROH -> 'false'
    - Heuristic / PP Plng Procedure / Requirement Strategy:
        * ROH e HALB sem Transfer (interno) -> VAZIOS
        * HALB com Transfer e FERT -> 'SAP_PP_002' / 4 / 20

    Gera:
    - Excel (Original Data + All Errors + Adjusted Summary)
    - errors_df, summary_df
    """
    import pandas as pd
    import io

    # 1) Carregar dataframes
    mara_df = load_xlsx(mara_file, sheet_name=0)
    marc_df = load_xlsx(marc_file, sheet_name=0)
    zplm_marc_df = load_xlsx(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    if mara_df is None or marc_df is None:
        raise ValueError("MARA.xlsx/MARC.xlsx não carregados.")

    # 2) Normalizações mínimas + filtro por Plant (visualização/exportação)
    for df in (mara_df, marc_df):
        df.columns = [str(c).strip() for c in df.columns]
    if plant_filter is not None and 'Plant' in marc_df.columns:
        pf = str(plant_filter).strip().upper()
        marc_df = marc_df[marc_df['Plant'].astype(str).str.strip().str.upper() == pf].copy()

    # 3) Mapa Material -> Material Type (MARA)
    if not {'Material', 'Material Type'}.issubset(mara_df.columns):
        raise ValueError("MARA.xlsx tem de conter 'Material' e 'Material Type'.")
    mt_map = dict(zip(
        mara_df['Material'].astype(str).str.strip(),
        mara_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    # 4) Mapa (Material, Plant) -> Transfer (ZPLM MARC)
    transfer_map = {}
    try:
        transfer_map = build_material_transfer_map(zplm_marc_df)
    except Exception:
        transfer_map = {}

    # 5) Helper p/ normalizar valores e apanhar colunas-alvo
    def _s(x):  # string normalizada
        import pandas as pd
        return '' if pd.isna(x) else str(x).strip()

    def pick_col(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    col_adv = pick_col(marc_df, ['Advanced Planning'])
    col_heur = pick_col(marc_df, ['Heuristic'])
    col_pp = pick_col(marc_df, ['PP Plng Procedure'])
    col_req = pick_col(marc_df, ['Requirement Strategy'])
    for needed, name in [(col_adv, 'Advanced Planning'),
                         (col_heur, 'Heuristic'),
                         (col_pp, 'PP Plng Procedure'),
                         (col_req, 'Requirement Strategy')]:
        if needed is None:
            raise ValueError(f"MARC.xlsx tem de conter a coluna '{name}'.")

    # 6) Loop MARC e aplicar regras
    resultados = []
    consider_types = {'FERT', 'HALB', 'ROH'}

    for idx, row in marc_df.iterrows():
        material = _s(row.get('Material'))
        plant = _s(row.get('Plant'))
        mt = mt_map.get(material, '').upper()

        if mt not in consider_types:
            continue  # fora do escopo

        # Detectar HALB com/sem Transfer
        is_transfer = transfer_map.get((material, plant), False)
        # Valores
        adv = _s(row.get(col_adv)).lower()  # esperamos 'true'/'false'
        heur = _s(row.get(col_heur))
        pp = normalize_code(row.get(col_pp), as_int=True) if col_pp else ''
        req = normalize_code(row.get(col_req), as_int=True) if col_req else ''

        # --- Regras Advanced Planning ---
        # HALB com Transfer  OU FERT  -> "true"
        # HALB sem Transfer (interno) OU ROH -> "false"
        if (mt == 'HALB' and is_transfer) or (mt == 'FERT'):
            if adv == '' or adv != 'true':
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'Advanced Planning',
                    'Error': "Invalid value: deve ser 'true' (HALB com Transfer ou FERT)",
                    'Material': material,
                    'Plant': plant
                })
        elif (mt == 'HALB' and not is_transfer) or (mt == 'ROH'):
            if adv == '' or adv != 'false':
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'Advanced Planning',
                    'Error': "Invalid value: deve ser 'false' (HALB interno ou ROH)",
                    'Material': material,
                    'Plant': plant
                })

        # --- Regras Heuristic / PP / Requirement ---
        if (mt == 'ROH') or (mt == 'HALB' and not is_transfer):
            # Devem estar VAZIOS
            if heur != '':
                resultados.append({
                    'Line': idx + 2, 'Field': 'Heuristic',
                    'Error': "Should be empty (ROH ou HALB interno)",
                    'Material': material, 'Plant': plant
                })
            if _s(row.get(col_pp)) != '':
                resultados.append({
                    'Line': idx + 2, 'Field': 'PP Plng Procedure',
                    'Error': "Should be empty (ROH ou HALB interno)",
                    'Material': material, 'Plant': plant
                })
            if _s(row.get(col_req)) != '':
                resultados.append({
                    'Line': idx + 2, 'Field': 'Requirement Strategy',
                    'Error': "Should be empty (ROH ou HALB interno)",
                    'Material': material, 'Plant': plant
                })
        else:
            # HALB com Transfer OU FERT -> valores fixos
            if heur != 'SAP_PP_002':
                resultados.append({
                    'Line': idx + 2, 'Field': 'Heuristic',
                    'Error': "Invalid value: deve ser 'SAP_PP_002' (HALB com Transfer ou FERT)",
                    'Material': material, 'Plant': plant
                })
            if pp != '4':
                resultados.append({
                    'Line': idx + 2, 'Field': 'PP Plng Procedure',
                    'Error': "Invalid value: deve ser 4 (HALB com Transfer ou FERT)",
                    'Material': material, 'Plant': plant
                })
            if req != '20':
                resultados.append({
                    'Line': idx + 2, 'Field': 'Requirement Strategy',
                    'Error': "Invalid value: deve ser 20 (HALB com Transfer ou FERT)",
                    'Material': material, 'Plant': plant
                })

    errors_df = pd.DataFrame(resultados)

    # 7) Original Data enriquecido (útil para análise)
    out_cols = ['Material', 'Plant', 'Material Type',
                'Advanced Planning', 'Heuristic', 'PP Plng Procedure', 'Requirement Strategy']
    marc_enriched = marc_df.copy()
    marc_enriched['Material Type'] = marc_enriched['Material'].astype(str).str.strip().map(mt_map).fillna('')
    # Marca HALB Transfer (para referência)
    marc_enriched['HALB Transfer'] = [
        bool(transfer_map.get((str(r['Material']).strip(), str(r.get('Plant', '')).strip()), False))
        for _, r in marc_enriched.iterrows()
    ]
    original_df = marc_enriched[[c for c in out_cols if c in marc_enriched.columns]].copy()
    if 'HALB Transfer' in marc_enriched.columns:
        original_df['HALB Transfer'] = marc_enriched['HALB Transfer']

    # 8) Pós-processamento (Views/Roles/Summary) usando o RolesViews embutido
    rolesviews_dict = load_xlsx(get_rolesviews_bytesio(), sheet_name=None)
    roles_df = rolesviews_dict.get("Roles")
    views_df = rolesviews_dict.get("Views")
    (
        errors_df,
        errors_filtered_df,
        views_missing_output_df,
        contagem_final_df,
        coverage_df,
        view_tabs_dict
    ) = postprocess_with_views_roles_and_classification(
        errors_df=errors_df,
        mara_df=mara_df,
        marc_df=marc_df,
        roles_df=roles_df,
        views_df=views_df,
        plant_filter=plant_filter,
    )

    # 9) Escrever Excel em memória (Original Data + erros + summary)
    file_bytes = io.BytesIO()
    with pd.ExcelWriter(file_bytes, engine="openpyxl") as writer:
        original_df.to_excel(writer, sheet_name="Original Data", index=False)
        errors_df.to_excel(writer, sheet_name="All Errors", index=False)
        contagem_final_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=1)
        # (opcional) incluir as folhas de views/coverage como nas outras páginas
        base_row = len(contagem_final_df) + 5
        if not coverage_df.empty:
            coverage_df.to_excel(writer, sheet_name="Adjusted Summary", index=False, startrow=base_row)
        for view_code, df_view in view_tabs_dict.items():
            sheet = f"Missing_{view_code}"[:31]
            df_view.to_excel(writer, sheet_name=sheet, index=False)

    file_bytes.seek(0)

    # 10) Summary para interface
    if 'Field' in errors_df.columns and not errors_df.empty:
        summary_df = (
            errors_df['Field']
            .value_counts()
            .rename_axis('Field')
            .reset_index(name='Adjusted Nº of Errors')
        )
    else:
        import pandas as pd
        summary_df = pd.DataFrame({"Field": [], "Adjusted Nº of Errors": []})

    return file_bytes, errors_df, summary_df


def build_mrp_log_payload(
    *,
    user: str,
    plant_filter: str | None,
    mara_file,
    marc_file,
    zplm_marc_file,
    errors_df,
    summary_df,
):
    """
    Monta o payload de log para a página MRP.
    Mantém o mesmo formato das outras páginas para consistência.
    """
    from datetime import datetime

    def _name(f, default):
        return getattr(f, "name", default) if f is not None else default

    total_errors = int(errors_df.shape[0]) if errors_df is not None else 0
    total_fields = int(errors_df['Field'].nunique()) if (errors_df is not None and 'Field' in errors_df.columns) else 0
    total_materials = int(errors_df['Material'].nunique()) if (errors_df is not None and 'Material' in errors_df.columns) else 0

    payload = {
        "type": "mrp",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "user": user,
        "plant_filter": plant_filter if plant_filter else "ALL",
        "metrics": {
            "total_errors": total_errors,
            "fields_with_error": total_fields,
            "materials_with_error": total_materials,
        },
        "files": {
            "mara": _name(mara_file, "MARA.xlsx"),
            "marc": _name(marc_file, "MARC.xlsx"),
            "zplm_marc": _name(zplm_marc_file, "ZPLM MARC.xlsx"),
        },
    }
    return payload