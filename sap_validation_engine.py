import pandas as pd
import io
import os, json
from datetime import datetime
from openpyxl import load_workbook
from utils.fast_loader import load_xlsx_fast
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
    # ➤ NOVO: exceção para Description quando PSMS = 70 ou 75
    if coluna == "Description" and "Plant Specific Material Status" in df.columns:
        df['_PSMS'] = df['Plant Specific Material Status'].apply(
            lambda x: int(float(str(x))) if str(x).strip() not in ['', 'nan'] else None
        )

    if coluna in df.columns:
        missing = df[df[coluna].isnull()]

        for idx in missing.index:
            # ➤ NOVO: ignorar erro se PSMS for 70 ou 75
            if coluna == "Description" and "_PSMS" in df.columns:
                psms = df.at[idx, '_PSMS']
                if psms in (70, 75):
                    continue  # ignorar o erro

            resultados.append({
                'Line': idx + 2,
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

    has_mt = 'Material Type' in df.columns

    for idx, row in df.iterrows():
        erros_linha = []

        raw = row['Consumption Mode']
        val = normalize_code(raw, as_int=True)  # normaliza 0.0 -> 0, etc.

        material = str(row['Material']).strip() if 'Material' in df.columns else ''
        plant = str(row['Plant']).strip() if 'Plant' in df.columns else ''
        mt = str(row['Material Type']).strip().upper() if has_mt else ''

        # --- CASO FERT ---
        if mt == 'FERT':

            # Para FERT: 0 ou vazio são aceitáveis
            if val not in ('', '0'):
                erros_linha.append(
                    f"Invalid value: '{raw}' (Material Type FERT requires Consumption Mode = '0' or empty)"
                )

        else:
            # --- REGRA ANTIGA PARA OUTROS MATERIAIS ---
            if val == '':
                erros_linha.append('Mandatory field blank')
            elif val != '2':
                erros_linha.append(
                    f"Invalid value: '{raw}' (should be '2' for non-FERT materials)"
                )

        if erros_linha:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Consumption Mode',
                'Error': '\n'.join(erros_linha),
                'Material': material,
                'Plant': plant
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

def validar_item_category_group(df, resultados, mvke_df):
    """
    Regras finais:

    - A partir do SAPUI5 Export, obter materiais com Material Type ∈ {FERT, HALB, ROH}.
    - Para cada material:
         * Se NÃO existir na MVKE → ignorar (sem erro)
         * Se EXISTIR na MVKE:
             - TODAS as linhas para esse material têm de ter Item Category Group = 'NORM'
             - Se existir qualquer valor diferente → ERRO
    """

    # SAPUI5 colunas obrigatórias
    if 'Material' not in df.columns or 'Material Type' not in df.columns:
        print("⚠️ Missing columns in SAPUI5 Export for Item Category Group validation.")
        return

    # MVKE obrigatório
    if mvke_df is None or mvke_df.empty:
        print("⚠️ MVKE empty or missing — skipping Item Category Group validation.")
        return

    # Normalizar MVKE
    mvke = mvke_df.copy()
    mvke.columns = [c.strip() for c in mvke.columns]

    # encontrar colunas
    col_mat = None
    col_icg = None

    for c in ['Material', 'MATNR']:
        if c in mvke.columns:
            col_mat = c
            break

    for c in ['Item Category Group', 'Item Cat Group', 'MVGR1', 'KDMAT']:
        if c in mvke.columns:
            col_icg = c
            break

    if col_mat is None or col_icg is None:
        print("⚠️ MVKE missing Material or Item Category Group columns — skipping validation.")
        return

    mvke['MAT_norm'] = mvke[col_mat].astype(str).str.strip()
    mvke['ICG_norm'] = mvke[col_icg].astype(str).str.strip().str.upper()

    # mapear Material -> set de valores encontrados
    mvke_map = {}
    for mat, grp in mvke.groupby('MAT_norm'):
        mvke_map[mat] = set(grp['ICG_norm'])

    # filtrar FERT / HALB / ROH
    sap_subset = df.loc[
        df['Material Type'].astype(str).str.strip().str.upper().isin(['FERT', 'HALB', 'ROH'])
    ]

    for idx, row in sap_subset.iterrows():
        material = str(row['Material']).strip()
        plant = str(row.get('Plant', '')).strip()

        # se não existir na MVKE --> ignorar
        if material not in mvke_map:
            continue

        icg_set = mvke_map[material]

        # regra final:
        # TODOS os valores devem ser NORM — se houver qualquer outro → erro
        if icg_set != {'NORM'}:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Item Category Group (MVKE)',
                'Error': (
                    f"Invalid in MVKE: For material '{material}' all entries must have "
                    f"Item Category Group = 'NORM' (found: {sorted(icg_set)})"
                ),
                'Material': material,
                'Plant': plant
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
    - EX só permitido para Plant Type FP e MP (regra base).
    - ZK só permitido para Plant Type PP (regra base).
    - EXCEÇÃO: para Plant Type MP cujo Material contenha o token '14K450' (ex.: 'VPAKLF-14K450-BCD'),
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
        plant_type = str(row['Plant Type']).strip() if 'Plant Type' in df.columns else ''
        lot_sizing = '' if pd.isna(row['Lot Sizing Procedure']) else str(row['Lot Sizing Procedure']).strip()
        material = str(row['Material']).strip() if 'Material' in df.columns else ''

        # 1) Campo obrigatório
        if lot_sizing == '':
            erros_linha.append('Mandatory field blank')
        else:
            # 2) Exceção MP + token 14K450 → aceita ZK ou EX
            is_halb_14k450 = (plant_type == 'MP' and material_has_14k450_token(material))
            if is_halb_14k450:
                if lot_sizing not in {'EX', 'ZK'}:
                    erros_linha.append(
                        f"Invalid value: '{lot_sizing}' (for Plant Type MP with Material containing token '14K450', allowed: ZK or EX)"
                    )
            else:
                # 3) Regras padrão
                if lot_sizing == 'EX' and plant_type not in ['FP', 'MP']:
                    erros_linha.append(f"Value 'EX' only allowed for Plant Type FP or MP (current: '{plant_type}')")
                if lot_sizing == 'ZK' and plant_type != 'PP':
                    erros_linha.append(f"Value 'ZK' only allowed for Plant Type PP or exception (current: '{plant_type}')")

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
    import pandas as pd

    required = {'Material', 'Plant', 'Plant Type', 'Planning Calendar', 'Plant Specific Material Status'}
    if not required.issubset(df.columns):
        missing = [c for c in required if c not in df.columns]
        print(f"⚠️ Missing columns for Planning Calendar validation: {missing}")
        return

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except:
            return None

    def contains_14k450(material):
        return '14K450' in str(material).upper()

    for idx, row in df.iterrows():
        erros = []

        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        mt       = str(row.get('Plant Type', '')).strip().upper()

        pc_raw   = row.get('Planning Calendar', None)
        pc       = '' if pd.isna(pc_raw) else str(pc_raw).strip().upper()

        psms     = parse_status(row.get('Plant Specific Material Status', None))
        has14k   = contains_14k450(material)

        # Validar apenas quando PSMS ∈ [10, 70]
        if psms is None or not (10 <= psms <= 70):
            continue

        # --------------------------------------------------------------
        # OVERRIDE PARA (MP + 14K450)
        # --------------------------------------------------------------
        if has14k and mt == 'MP':
            if pc == '':
                erros.append("For Plant Type 'MP' with material containing '14K450', Planning Calendar must not be empty")
            elif pc == 'PT2':
                erros.append("For Plant Type 'MP' with material containing '14K450', Planning Calendar must not be 'PT2'")

            if erros:
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'Planning Calendar',
                    'Error': ' \n'.join(erros),
                    'Material': material,
                    'Plant': plant
                })

            continue  # IGNORA regras normais 100%

        # --------------------------------------------------------------
        # REGRAS NORMAIS (sem override)
        # --------------------------------------------------------------
        if plant == 'PT01':
            if mt in ('FP', 'MP'):
                if pc != 'PT1':
                    erros.append(f"For Plant PT01 and Plant Type '{mt}', Planning Calendar must be 'PT1' (Current: '{pc_raw}')")

            elif mt == 'PP':
                if pc == '':
                    erros.append("For Plant PT01 and Plant Type 'PP', Planning Calendar must not be empty")
                elif pc == 'PT2':
                    erros.append("For Plant PT01 and Plant Type 'PP', Planning Calendar must not be 'PT2'")

        elif plant == 'PT02':
            if mt in ('FP', 'MP'):
                if pc != 'PT2':
                    erros.append(f"For Plant PT02 and Plant Type '{mt}', Planning Calendar must be 'PT2' (Current: '{pc_raw}')")

            elif mt == 'PP':
                if pc == '':
                    erros.append("For Plant PT02 and Plant Type 'PP', Planning Calendar must not be empty")

        # Adicionar ao relatório se houver erros
        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'Planning Calendar',
                'Error': ' \n'.join(erros),
                'Material': material,
                'Plant': plant
            })

def validar_planning_time_fence(df, resultados):
    import pandas as pd

    required = {'Material', 'Plant', 'MRP Type', 'Planning Time Fence'}
    if not required.issubset(df.columns):
        print("⚠️ Columns 'Material', 'Plant', 'MRP Type' or 'Planning Time Fence' not found.")
        return

    def parse_tf(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return float(str(val).strip())
        except:
            return 'INVALID'

    def parse_status(val):
        if pd.isna(val) or str(val).strip() == '':
            return None
        try:
            return int(float(str(val).strip()))
        except:
            return None

    for idx, row in df.iterrows():
        erros = []
        material = str(row.get('Material', '')).strip()
        plant    = str(row.get('Plant', '')).strip()
        mt       = str(row.get('MRP Type', '')).strip().upper()

        tf = parse_tf(row.get('Planning Time Fence', None))
        psms_val = parse_status(row.get('Plant Specific Material Status', None))

        # PSMS = 10 -> ignorar linha
        if psms_val == 10:
            continue

        # Valor inválido (letras, símbolos, etc.)
        if tf == 'INVALID':
            erros.append("Planning Time Fence contains non-numeric value.")
        else:

            # Regras por MRP Type
            if mt == 'PD':
                if tf not in (None, 0):
                    erros.append(f"For {mt}, Planning Time Fence should be empty or 0 (Current: '{tf}')")

            elif mt in ('P1', 'M1'):
                if tf is None or tf == 0:
                    erros.append(f"For {mt}, Planning Time Fence should be > 0 (Current: '{tf}')")

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
    Validação do MRP Type segundo regras finais confirmadas.

    Regras:
    - TF com semântica: TF>0 (Yes) ; TF=0 ou vazio (No)
    - MP & FP transfer (ZPLM MARC)
    - PSMS (Plant Specific Material Status): precedências 10, 25, 70, 75
    - PLANT (PT01/PT02) → coerência do Planning Calendar quando PSMS=25

    Precedências por Status:
    - PSMS = 10 → Ignora a linha (não valida nem contabiliza erros).
    - PSMS = 25 → MRP = ND; TF = 0 ou vazio; Planning Calendar:
                     PT01 → 'PT1'
                     PT02 → 'PT2'
    - PSMS ∈ {70, 75} → MRP = ND para TODOS os Plant Types (restantes regras não se aplicam)

    Tabela base (aplicada apenas quando não é 10/25/70/75):
    - Plant Type FP → MRP = M1, TF > 0
    - Plant Type MP & FP transfer → MRP ∈ {M1, P1, PD}, TF pode ser >0 ou 0/vazio
    - Plant Type MP sem transfer → MRP = PD, TF = 0 ou vazio
    - Plant Type PP com TF (>0) → MRP = P1
    - Plant Type PP sem TF (0/vazio) → MRP = PD
    """

    import pandas as pd

    needed = {'MRP Type', 'Plant Type', 'Planning Time Fence', 'Plant'}
    if not needed.issubset(df.columns):
        return

    has_psms = 'Plant Specific Material Status' in df.columns
    has_pc = 'Planning Calendar' in df.columns

    valores_permitidos = {'M1', 'PD', 'P1', 'ND', 'V1'}

    # Mapa de transferências (FP ou MP)
    transfer_map = {}
    if zplm_marc_df is not None:
        try:
            transfer_map = build_material_transfer_map(zplm_marc_df)
        except Exception:
            transfer_map = {}

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
            return None

    for idx, row in df.iterrows():
        erros = []

        mrp = '' if pd.isna(row['MRP Type']) else str(row['MRP Type']).strip().upper()
        pt  = '' if pd.isna(row['Plant Type']) else str(row['Plant Type']).strip().upper()
        plant = '' if pd.isna(row['Plant']) else str(row['Plant']).strip()

        tf_raw = row.get('Planning Time Fence', None)
        tf_val = parse_tf(tf_raw)

        psms = parse_status(row.get('Plant Specific Material Status', None)) if has_psms else None
        pc = '' if (not has_pc or pd.isna(row.get('Planning Calendar'))) else str(row['Planning Calendar']).strip().upper()

        material = str(row.get('Material', '')).strip()
        is_transfer = transfer_map.get((material, plant), False)

        # =====================================================
        # PSMS = 10 → ignorar linha
        # =====================================================
        if psms == 10:
            continue

        # =====================================================
        # PSMS = 25 → regras obrigatórias (ND + TF=0/empty + PC coerente)
        # =====================================================
        if psms == 25:
            # 1) MRP deve ser ND
            if mrp != 'ND':
                if mrp == '':
                    erros.append("Mandatory field blank (MRP Type); For Part Status = 25 it must be 'ND'")
                else:
                    erros.append(f"For Part Status = 25, MRP Type must be 'ND' (Current: '{mrp}')")

            # 2) TF deve ser 0 ou vazio
            if tf_val == 'INVALID':
                erros.append("Planning Time Fence invalid for Part Status = 25 (must be 0 or empty)")
            elif tf_val not in (None, 0):
                erros.append(f"For Part Status = 25, Planning Time Fence must be 0 or empty (Current: '{tf_raw}')")

            # 3) Planning Calendar coerente
            if has_pc:
                if plant == 'PT01' and pc != 'PT1':
                    erros.append(f"For Part Status = 25 & PT01, Planning Calendar must be 'PT1' (Current: '{pc}')")
                elif plant == 'PT02' and pc != 'PT2':
                    erros.append(f"For Part Status = 25 & PT02, Planning Calendar must be 'PT2' (Current: '{pc}')")

            if erros:
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'MRP Type',
                    'Error': ' \n'.join(erros),
                    'Material': material,
                    'Plant': plant
                })
            continue  # não aplicar regras base

        # =====================================================
        # PSMS = 70 ou 75 → somente MRP = ND (resto não aplica)
        # =====================================================
        if psms in (70, 75):
            if mrp != 'ND':
                if mrp == '':
                    erros.append("Mandatory field blank (MRP Type); for Part Status = 70/75 it must be 'ND'")
                else:
                    erros.append(f"For Part Status = {psms}, MRP Type must be 'ND' (Current: '{mrp}')")

            if erros:
                resultados.append({
                    'Line': idx + 2,
                    'Field': 'MRP Type',
                    'Error': ' \n'.join(erros),
                    'Material': material,
                    'Plant': plant
                })
            continue  # não aplicar regras base

        # =====================================================
        # Validações base (apenas quando não é PSMS 10/25/70/75)
        # =====================================================

        # MRP obrigatório e dentro do domínio
        if mrp == '':
            erros.append("Mandatory field blank (MRP Type)")
        elif mrp not in valores_permitidos:
            erros.append(f"Invalid MRP Type: '{mrp}'")

        # TF inválido
        if tf_val == 'INVALID':
            erros.append("Planning Time Fence invalid (must be numeric)")
        else:
            has_tf = (tf_val is not None) and (tf_val > 0)
            no_tf  = (tf_val is None) or (tf_val == 0)

            # -------------------------------------------------
            # FP NORMAL (não transfer)
            # -------------------------------------------------
            if pt == 'FP' and not is_transfer:
                if mrp != 'M1':
                    erros.append(f"For FP, MRP must be 'M1' (Current: '{mrp}')")
                if not has_tf:
                    erros.append(f"For FP, Time Fence must be > 0 (Current: '{tf_raw}')")

            # -------------------------------------------------
            # FP TRANSFER + MP TRANSFER
            # -------------------------------------------------
            elif pt in {'FP', 'MP'} and is_transfer:
                if mrp not in {'M1', 'P1', 'PD'}:
                    erros.append(
                        f"For {pt} with transfer, allowed MRP Types: M1, P1, PD (Current: '{mrp}')"
                    )
                # TF livre → sem validação

            # -------------------------------------------------
            # MP SEM TRANSFER
            # -------------------------------------------------
            elif pt == 'MP' and not is_transfer:
                if mrp != 'PD':
                    erros.append(f"For MP without transfer, MRP must be 'PD' (Current: '{mrp}')")
                if not no_tf:
                    erros.append(f"For MP without transfer, Time Fence must be 0 or empty (Current: '{tf_raw}')")

            # -------------------------------------------------
            # PP
            # -------------------------------------------------
            elif pt == 'PP':
                if has_tf:
                    if mrp != 'P1':
                        erros.append(f"For PP with Time Fence > 0, MRP must be 'P1' (Current: '{mrp}')")
                else:
                    if mrp != 'PD':
                        erros.append(f"For PP with Time Fence = 0/empty, MRP must be 'PD' (Current: '{mrp}')")

        # Final → guardar erros
        if erros:
            resultados.append({
                'Line': idx + 2,
                'Field': 'MRP Type',
                'Error': ' \n'.join(erros),
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

def validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados):
    """
    - Só valida FERT.
    - Extrai o sufixo do Product Hierarchy após 'P' na 13ª posição.
    - Compara com Enterproj Number da ZPLM MARA com normalização:
        * 20922.0 -> 20922
        * NaN/None/""/"nan" -> zeros (zfill do tamanho do sufixo do PH)
        * preserva zeros à esquerda via zfill
    """
    import pandas as pd
    import math
    import re

    needed_mara = {'Material', 'Material Type', 'Product Hierarchy'}
    if not needed_mara.issubset(mara_df.columns):
        print("⚠️ MARA não contém colunas necessárias.")
        return

    if zplm_mara_df is None or zplm_mara_df.empty:
        print("⚠️ ZPLM MARA não carregado → validação ignorada.")
        return

    # --- NÃO converter tudo para str já aqui (evita "nan") ---
    zplm = zplm_mara_df.copy()
    zplm.columns = [c.strip().upper() for c in zplm.columns]
    if not {'MATERIAL', 'ENTERPROJ NUMBER'}.issubset(zplm.columns):
        print("⚠️ ZPLM MARA sem MATERIAL/ENTERPROJ NUMBER.")
        return

    # Mapa Material -> Enterproj (valor cru, pode ser NaN ou float)
    zplm_map = dict(zip(
        zplm['MATERIAL'].astype(str).str.strip(),
        zplm['ENTERPROJ NUMBER']  # <- manter tipo original
    ))

    def _is_nan_like(v):
        return v is None or (isinstance(v, float) and math.isnan(v))

    def _only_digits(s: str) -> str:
        # se quiseres estritamente dígitos, descomenta:
        # return re.sub(r"\D+", "", s)
        return s

    def norm_side(value, target_len=None):
        """
        Normaliza o valor para comparação:
        - NaN/None/""/"nan" -> zeros (com tamanho target_len, se fornecido)
        - '20922.0' -> '20922'
        - preserva zeros à esquerda via zfill(target_len)
        - mantém apenas dígitos se _only_digits estiver ativo
        """
        # NaN-like
        if _is_nan_like(value):
            return ("0" * target_len) if (target_len and target_len > 0) else "0"

        s = str(value).strip()

        # strings 'nan'/'None'
        if s.lower() in ("nan", "none", ""):
            return ("0" * target_len) if (target_len and target_len > 0) else "0"

        # tentar converter floats "xxxxx.0" para inteiro
        try:
            # só converte se for numérico
            # permite '20922.0' ou '00012'
            f = float(s)
            i = int(f)
            s = str(i)
        except Exception:
            # não numérico: mantém, mas pode filtrar dígitos
            pass

        s = _only_digits(s)

        if target_len and target_len > 0:
            s = s.zfill(target_len)

        return s

    for idx, row in mara_df.iterrows():
        material = str(row['Material']).strip()
        mt = str(row['Material Type']).strip().upper()
        if mt != 'FERT':  # só FERT
            continue

        ph_raw = str(row['Product Hierarchy']).strip()

        # Validar 'P' na 13ª posição
        if len(ph_raw) < 13 or ph_raw[12] != 'P':
            resultados.append({
                'Line': '',
                'Field': 'Product Hierarchy',
                'Material': material,
                'Plant': '',
                'Error': "Invalid Product Hierarchy: missing 'P' at 13th position (FERT rule)"
            })
            continue

        # sufixo após 'P'
        suffix = ph_raw[13:].strip()
        # normalizar lado PH (apenas garantir dígitos e usar o comprimento como referência)
        ph_norm = norm_side(suffix, target_len=len(suffix))

        # obter Enterproj cru e normalizar com mesmo comprimento
        enterproj_raw = zplm_map.get(material, None)
        enterproj_norm = norm_side(enterproj_raw, target_len=len(suffix))

        if ph_norm != enterproj_norm:
            resultados.append({
                'Line': '',
                'Field': 'Product Hierarchy',
                'Material': material,
                'Plant': '',
                'Error': (
                    f"PH mismatch: ProductHierarchy→'{ph_norm}' "
                    f"≠ Enterproj→'{enterproj_norm}'"
                )
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
    if 'Plant' not in df.columns or 'Issue Storage Location' not in df.columns:
        print("⚠️ Columns 'Plant' or 'Storage Location' not found.")
        return

    for idx, row in df.iterrows():
        erros_linha = []
        plant = str(row['Plant']).strip()
        storage_location = str(row['Issue Storage Location']).strip()

        # Field obrigatório
        if pd.isna(row['Issue Storage Location']) or storage_location == '':
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
                'Field': 'Issue Storage Location',
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

def validar_inspection_type_qmat(sap_df, marc_df, qmat_df, resultados):
    """
    Validação QMAT — Inspection Type

    Rule 2: Inspection type '08'
      - Material Type in {ROH, HALB}
      - Para cada entrada MARC (Material/Plant)
        -> consultar QMAT (Material/Plant)
        -> se existir Inspection Type '08'
           -> Active flag NÃO pode ser 'X'

    Rule 3: Inspection type '1708'
      - Material Type in {ROH, HALB, FERT}
      - mesmo processo que acima
    """

    if qmat_df is None or qmat_df.empty:
        print("⚠️ QMAT.xlsx em falta, validação QMAT ignorada.")
        return

    qmat = qmat_df.copy()
    qmat.columns = [c.strip().upper() for c in qmat.columns]

    # detectar colunas
    col_mat = next((c for c in ["MATERIAL", "MATNR"] if c in qmat.columns), None)
    col_plant = next((c for c in ["PLANT", "WERKS"] if c in qmat.columns), None)
    col_insp = next((c for c in ["INSPECTION TYPE", "QMTXT", "ART"] if c in qmat.columns), None)
    col_active = next((c for c in ["INSPTYPE-MATL ACTIVE", "ACTIVE", "STATUS"] if c in qmat.columns), None)

    if not all([col_mat, col_plant, col_insp, col_active]):
        print("⚠️ QMAT com colunas insuficientes. Ignorado.")
        return

    qmat["MAT"] = qmat[col_mat].astype(str).str.strip()
    qmat["PLT"] = qmat[col_plant].astype(str).str.strip()
    qmat["INSP"] = qmat[col_insp].astype(str).str.strip().str.upper()
    qmat["ACT"] = qmat[col_active].astype(str).str.strip().str.upper()

    # mapear material -> tipo de material (SAPUI5)
    mt_map = dict(zip(
        sap_df['Material'].astype(str).str.strip(),
        sap_df['Material Type'].astype(str).str.strip().str.upper()
    ))

    marc = marc_df.copy()
    marc.columns = [c.strip() for c in marc.columns]

    for idx, row in marc.iterrows():
        material = str(row.get("Material", "")).strip()
        plant    = str(row.get("Plant", "")).strip()
        mt       = mt_map.get(material, "")

        # RULE 2 — Inspection Type = 08, para HALB/ROH
        if mt in ["HALB", "ROH"]:
            subset = qmat[
                (qmat["MAT"] == material) &
                (qmat["PLT"] == plant) &
                (qmat["INSP"] == "08")
            ]

            for _, r in subset.iterrows():
                if r["ACT"] == "X":
                    resultados.append({
                        "Line": idx + 2,
                        "Field": "Inspection Type (QMAT 08)",
                        "Error": "Inspection Type '08' has Active flag = 'X' in QMAT",
                        "Material": material,
                        "Plant": plant
                    })

        # RULE 3 — Inspection Type = 1708, para FERT/HALB/ROH
        if mt in ["FERT", "HALB", "ROH"]:
            subset = qmat[
                (qmat["MAT"] == material) &
                (qmat["PLT"] == plant) &
                (qmat["INSP"] == "1708")
            ]

            for _, r in subset.iterrows():
                if r["ACT"] == "X":
                    resultados.append({
                        "Line": idx + 2,
                        "Field": "Inspection Type (QMAT 1708)",
                        "Error": "Inspection Type '1708' has Active flag = 'X' in QMAT",
                        "Material": material,
                        "Plant": plant
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

def validar_fp_pt01_pt02_status55(zplm_marc_df, marc_df):
    """
    1. Identifica materiais com Material Plant Type = 'FP' simultaneamente em PT01 e PT02
       (usando ZPLM MARC)
    2. Para esses materiais, verifica na MARC se existe Plant-Specific Material Status = 55
    3. Se sim, retorna DataFrame com os erros
    """

    import pandas as pd

    if zplm_marc_df is None or zplm_marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    if marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    # Normalizar colunas
    zplm = zplm_marc_df.copy()
    zplm.columns = [c.strip().upper() for c in zplm.columns]
    marc = marc_df.copy()
    marc.columns = [c.strip().title() for c in marc.columns]  # Plant-Specific Material Status

    # 1) Filtrar FP → PT01/PT02
    df_fp = zplm[
        (zplm["MATERIAL PLANT TYPE"].astype(str).str.upper() == "FP") &
        (zplm["PLANT"].isin(["PT01", "PT02"]))
    ]

    fp_groups = df_fp.groupby("MATERIAL")["PLANT"].agg(set)

    materiais_fp_2plantas = [
        mat for mat, plants in fp_groups.items()
        if {"PT01", "PT02"}.issubset(plants)
    ]

    if not materiais_fp_2plantas:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    # 2) Procurar status 55 na MARC
    marc_filtered = marc[
        marc["Material"].astype(str).str.strip().isin(materiais_fp_2plantas)
    ]

    erros = []
    for idx, row in marc_filtered.iterrows():
        material = str(row["Material"]).strip()
        plant = str(row["Plant"]).strip()
        status = str(row["Plant-Specific Material Status"]).strip()

        if status == "55":
            erros.append({
                "Material": material,
                "Plant": plant,
                "Error": "Material with FP in PT01 AND PT02 cannot have status 55"
            })

    return pd.DataFrame(erros)

def validar_fp55_eop_vs_service_liability(zplm_marc_df, marc_df):
    """
    Regra Corrigida:
      Para (Material, Plant) com:
        - Material Plant Type = 'FP' (ZPLM MARC)
        - PSMS = 55 (MARC)
      Validar:
        (A) Service Liability Date vazio        -> ERRO
        (B) EOP Date >= Service Liability Date  -> ERRO
    """

    import pandas as pd
    import numpy as np

    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "EOP Date", "Service Liability Date", "Error"
        ])

    # -----------------------------
    # Normalização
    # -----------------------------
    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    # Helper pick of columns
    def pick(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    # Pick columns
    z_mat   = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant = pick(zplm, ["PLANT", "WERKS"])
    z_mptype= pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    z_eop   = pick(zplm, ["EOP DATE", "EOP", "END OF PROD DATE", "END OF PRODUCTION DATE"])
    z_srv   = pick(zplm, ["SERVICE LIABILITY DATE", "SERVICE LIAB DATE", "SRV LIABILITY DATE"])

    m_mat   = pick(marc, ["MATERIAL", "MATNR"])
    m_plant = pick(marc, ["PLANT", "WERKS"])
    m_psms  = pick(marc, ["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])

    needed_z = [z_mat, z_plant, z_mptype]
    needed_m = [m_mat, m_plant, m_psms]

    if any(c is None for c in needed_z) or any(c is None for c in needed_m):
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "EOP Date", "Service Liability Date", "Error"
        ])

    # -----------------------------
    # Filtrar FP no ZPLM
    # -----------------------------
    z = zplm.copy()
    z[z_mat]   = z[z_mat].astype(str).str.strip()
    z[z_plant] = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype]= z[z_mptype].astype(str).str.strip().str.upper()

    z_fp = z[z[z_mptype] == "FP"].copy()

    # -----------------------------
    # Filtrar PSMS 55 no MARC
    # -----------------------------
    m = marc.copy()
    m[m_mat]   = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()

    def _norm_psms(x):
        if pd.isna(x): return ""
        s = str(x).strip()
        try:
            return str(int(float(s)))
        except:
            return s

    m[m_psms] = m[m_psms].apply(_norm_psms)
    m55 = m[m[m_psms] == "55"].copy()

    # -----------------------------
    # Criar chave de interseção
    # -----------------------------
    key_fp  = set(zip(z_fp[z_mat], z_fp[z_plant]))
    key_55  = set(zip(m55[m_mat], m55[m_plant]))
    key_int = key_fp & key_55

    if not key_int:
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "EOP Date", "Service Liability Date", "Error"
        ])

    z_fp["_KEY"] = list(zip(z_fp[z_mat], z_fp[z_plant]))
    df_target = z_fp[z_fp["_KEY"].isin(key_int)].copy()

    # -----------------------------
    # Parse de datas
    # -----------------------------
    def _dt(val):
        if val is None: return pd.NaT
        if isinstance(val, float) and np.isnan(val): return pd.NaT

        s = str(val).strip()
        if s == "" or s.lower() in ("nan", "none"): return pd.NaT

        # tentar excel serial
        try:
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
        except:
            pass

        # parsing genérico
        try:
            return pd.to_datetime(s, errors="coerce", dayfirst=True)
        except:
            return pd.NaT

    errors = []

    for idx, row in df_target.iterrows():
        material = row[z_mat]
        plant    = row[z_plant]
        mptype   = row[z_mptype]

        srv_raw = row.get(z_srv, None)
        eop_raw = row.get(z_eop, None)

        srv = _dt(srv_raw)
        eop = _dt(eop_raw)

        # ❌ A) Service Liability vazio
        if pd.isna(srv):
            errors.append({
                "Material": material,
                "Plant": plant,
                "Material Plant Type": mptype,
                "PSMS": "55",
                "EOP Date": eop_raw,
                "Service Liability Date": srv_raw,
                "Error": "Service Liability Date empty for FP with status 55"
            })
            continue

        # ❌ B) EOP >= Service Liability -> ERRO (REGRA CORRIGIDA)
        if not pd.isna(eop) and eop >= srv:
            errors.append({
                "Material": material,
                "Plant": plant,
                "Material Plant Type": mptype,
                "PSMS": "55",
                "EOP Date": eop_raw,
                "Service Liability Date": srv_raw,
                "Error": "EOP Date must be strictly earlier than Service Liability Date"
            })

    return pd.DataFrame(errors, columns=[
        "Material", "Plant", "Material Plant Type", "PSMS",
        "EOP Date", "Service Liability Date", "Error"
    ])

def validar_fp55_service_liability_expired(zplm_marc_df, marc_df):
    """
    Regra:
      Para (Material, Plant) com:
        - Material Plant Type = 'FP' (ZPLM MARC)
        - PSMS = 55 (MARC)

      ERRO se:
        - Service Liability Date ≤ hoje (apenas quando preenchido)

    Retorna colunas:
      Material, Plant, Material Plant Type, PSMS,
      Service Liability Date, Error
    """

    import pandas as pd
    import numpy as np

    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "Service Liability Date", "Error"
        ])

    # -----------------------------
    # Normalize headers
    # -----------------------------
    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    # helper to pick cols
    def pick(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    # cols from ZPLM
    z_mat   = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant = pick(zplm, ["PLANT", "WERKS"])
    z_mptype= pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])
    z_srv   = pick(zplm, ["SERVICE LIABILITY DATE", "SERVICE LIAB DATE", "SRV LIABILITY DATE"])

    # cols from MARC
    m_mat   = pick(marc, ["MATERIAL", "MATNR"])
    m_plant = pick(marc, ["PLANT", "WERKS"])
    m_psms  = pick(marc, ["PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS", "MTSTB"])

    if not all([z_mat, z_plant, z_mptype, z_srv, m_mat, m_plant, m_psms]):
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "Service Liability Date", "Error"
        ])

    # -----------------------------
    # filter FP
    # -----------------------------
    z = zplm.copy()
    z[z_mat]   = z[z_mat].astype(str).str.strip()
    z[z_plant] = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype]= z[z_mptype].astype(str).str.strip().str.upper()
    z_fp = z[z[z_mptype] == "FP"].copy()

    # -----------------------------
    # filter PSMS=55
    # -----------------------------
    m = marc.copy()
    m[m_mat]   = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()

    def _norm_psms(x):
        if pd.isna(x): return ""
        try:
            return str(int(float(str(x).strip())))
        except:
            return str(x).strip()

    m[m_psms] = m[m_psms].apply(_norm_psms)
    m55 = m[m[m_psms] == "55"]

    # -----------------------------
    # match keys
    # -----------------------------
    key_fp  = set(zip(z_fp[z_mat], z_fp[z_plant]))
    key_55  = set(zip(m55[m_mat], m55[m_plant]))
    key_int = key_fp & key_55

    if not key_int:
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "Service Liability Date", "Error"
        ])

    z_fp["_KEY"] = list(zip(z_fp[z_mat], z_fp[z_plant]))
    df_target = z_fp[z_fp["_KEY"].isin(key_int)].copy()

    # -----------------------------
    # Parse date
    # -----------------------------
    def _dt(val):
        if val is None: return pd.NaT
        if isinstance(val, float) and np.isnan(val): return pd.NaT
        s = str(val).strip()
        if s == "" or s.lower() in ("nan", "none"): return pd.NaT

        # Excel serial
        try:
            if isinstance(val, (int, float)):
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(val), unit="D")
        except:
            pass

        # generic
        return pd.to_datetime(s, errors="coerce", dayfirst=True)

    today = pd.Timestamp.today().normalize()

    errors = []

    for idx, row in df_target.iterrows():

        material = row[z_mat]
        plant    = row[z_plant]
        mptype   = row[z_mptype]
        srv_raw  = row[z_srv]

        srv = _dt(srv_raw)

        # ❌ erro se srv <= hoje
        if not pd.isna(srv) and srv <= today:
            errors.append({
                "Material": material,
                "Plant": plant,
                "Material Plant Type": mptype,
                "PSMS": "55",
                "Service Liability Date": srv_raw,
                "Error": "Service Liability Date is already expired (≤ today)"
            })

    return pd.DataFrame(errors, columns=[
        "Material", "Plant", "Material Plant Type", "PSMS",
        "Service Liability Date", "Error"
    ])

def validar_fppp70_standard_price_mbew(zplm_marc_df, marc_df, mbew_df):
    """
    Regra:
      Para materiais com:
        - Material Plant Type ∈ {FP, MP, PP}
        - PSMS = 70 (MARC)
      ERRO quando:
        - Standard Price (MBEW.STPRS) está vazio ou = 0

    Retorna colunas:
      Material, Plant, Material Plant Type, PSMS, Standard Price, Error
    """
    import pandas as pd
    import numpy as np

    # Validate input
    if any(df is None or df.empty for df in [zplm_marc_df, marc_df, mbew_df]):
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "Standard Price", "Error"
        ])

    # -------------------------
    # Normalize headers
    # -------------------------
    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    mbew = mbew_df.copy()
    mbew.columns = [str(c).strip().upper() for c in mbew.columns]

    # Helper to pick cols
    def pick(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    # Extract relevant columns
    z_mat    = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant  = pick(zplm, ["PLANT", "WERKS"])
    z_mptype = pick(zplm, ["MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"])

    m_mat   = pick(marc, ["MATERIAL", "MATNR"])
    m_plant = pick(marc, ["PLANT", "WERKS"])
    m_psms  = pick(marc, ["PLANT-SPECIFIC MATERIAL STATUS", "MTSTB"])

    mb_mat   = pick(mbew, ["MATERIAL", "MATNR"])
    mb_plant = pick(mbew, ["VALUATION AREA", "BWKEY", "PLANT", "WERKS"])
    mb_price = pick(mbew, ["STANDARD PRICE", "STPRS", "STD PRICE"])

    if not all([z_mat, z_plant, z_mptype, m_mat, m_plant, m_psms, mb_mat, mb_plant, mb_price]):
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "Standard Price", "Error"
        ])

    # -------------------------
    # Filter ZPLM: FP / MP / PP
    # -------------------------
    z = zplm[[z_mat, z_plant, z_mptype]].copy()
    z[z_mat]    = z[z_mat].astype(str).str.strip()
    z[z_plant]  = z[z_plant].astype(str).str.strip().str.upper()
    z[z_mptype] = z[z_mptype].astype(str).str.strip().str.upper()

    z_sel = z[z[z_mptype].isin(["FP", "MP", "PP"])].copy()

    # -------------------------
    # Filter MARC: PSMS = 70
    # -------------------------
    m = marc[[m_mat, m_plant, m_psms]].copy()
    m[m_mat]   = m[m_mat].astype(str).str.strip()
    m[m_plant] = m[m_plant].astype(str).str.strip().str.upper()

    def norm_status(x):
        try:
            return str(int(float(str(x).strip())))
        except:
            return ""
    m[m_psms] = m[m_psms].apply(norm_status)

    m70 = m[m[m_psms] == "70"]

    # -------------------------
    # Filter MBEW: Join for Standard Price
    # -------------------------
    mb = mbew[[mb_mat, mb_plant, mb_price]].copy()
    mb[mb_mat]   = mb[mb_mat].astype(str).str.strip()
    mb[mb_plant] = mb[mb_plant].astype(str).str.strip().str.upper()

    # -------------------------
    # Intersection keys
    # -------------------------
    key_z  = set(zip(z_sel[z_mat], z_sel[z_plant]))
    key_m  = set(zip(m70[m_mat], m70[m_plant]))
    key_ok = key_z & key_m

    if not key_ok:
        return pd.DataFrame(columns=[
            "Material", "Plant", "Material Plant Type", "PSMS",
            "Standard Price", "Error"
        ])

    # Convert ZPLM to key
    z_sel["_KEY"] = list(zip(z_sel[z_mat], z_sel[z_plant]))

    z_final = z_sel[z_sel["_KEY"].isin(key_ok)].copy()
    z_final.drop(columns=["_KEY"], inplace=True)

    # Merge Standard Price
    merged = z_final.merge(
        mb,
        left_on=[z_mat, z_plant],
        right_on=[mb_mat, mb_plant],
        how="left"
    )

    merged = merged.merge(
        m70[[m_mat, m_plant, m_psms]],
        left_on=[z_mat, z_plant],
        right_on=[m_mat, m_plant],
        how="left"
    )

    # -------------------------
    # Validate Standard Price
    # -------------------------
    errors = []

    for idx, row in merged.iterrows():
        mat   = row[z_mat]
        plant = row[z_plant]
        mptype= row[z_mptype]
        psms  = row[m_psms]
        price = row[mb_price]

        # Empty or zero?
        if price in [None, "", " ", 0, 0.0] or (isinstance(price, float) and np.isnan(price)):
            errors.append({
                "Material": mat,
                "Plant": plant,
                "Material Plant Type": mptype,
                "PSMS": psms,
                "Standard Price": price,
                "Error": "Standard Price is empty or zero for FP/MP/PP with status 70"
            })

    return pd.DataFrame(errors, columns=[
        "Material", "Plant", "Material Plant Type", "PSMS",
        "Standard Price", "Error"
    ])

def validar_mara_vs_zplm_mara_existence(mara_df, zplm_mara_df):
    """
    Verifica se TODOS os materiais da MARA existem na ZPLM MARA.
    ⚠ Não usa Plant (nível de material).
    Retorna DF de erros: Material | Error
    """
    import pandas as pd

    if mara_df is None or mara_df.empty or zplm_mara_df is None or zplm_mara_df.empty:
        return pd.DataFrame(columns=["Material", "Error"])

    # Normalizar cabeçalhos
    mara = mara_df.copy()
    mara.columns = [str(c).strip().upper() for c in mara.columns]

    zplm = zplm_mara_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    # Pick columns
    def pick(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None

    m_mat = pick(mara, ["MATERIAL", "MATNR"])
    z_mat = pick(zplm, ["MATERIAL", "MATNR"])

    if not all([m_mat, z_mat]):
        return pd.DataFrame(columns=["Material", "Error"])

    # Normalizar
    mara[m_mat] = mara[m_mat].astype(str).str.strip()
    zplm[z_mat] = zplm[z_mat].astype(str).str.strip()

    mara_set = set(mara[m_mat].unique())
    zplm_set = set(zplm[z_mat].unique())

    missing = sorted(mara_set - zplm_set)
    if not missing:
        return pd.DataFrame(columns=["Material", "Error"])

    return pd.DataFrame(
        [{"Material": mat, "Error": "Missing in ZPLM MARA"} for mat in missing],
        columns=["Material", "Error"]
    )

def validar_zplm_mara_vs_mara_existence(zplm_mara_df, mara_df):
    """
    Verifica materiais que existem em ZPLM MARA mas NÃO existem na MARA.
    Não usa Plant.
    Retorna: Material | Error
    """
    import pandas as pd

    if zplm_mara_df is None or zplm_mara_df.empty or mara_df is None or mara_df.empty:
        return pd.DataFrame(columns=["Material", "Error"])

    # Normalize
    zplm = zplm_mara_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    mara = mara_df.copy()
    mara.columns = [str(c).strip().upper() for c in mara.columns]

    def pick(df, names):
        for n in names:
            if n in df.columns:
                return n
        return None

    z_mat = pick(zplm, ["MATERIAL", "MATNR"])
    m_mat = pick(mara, ["MATERIAL", "MATNR"])

    if not all([z_mat, m_mat]):
        return pd.DataFrame(columns=["Material", "Error"])

    zplm[z_mat] = zplm[z_mat].astype(str).str.strip()
    mara[m_mat] = mara[m_mat].astype(str).str.strip()

    z_set = set(zplm[z_mat].unique())
    m_set = set(mara[m_mat].unique())

    missing = sorted(z_set - m_set)

    if not missing:
        return pd.DataFrame(columns=["Material", "Error"])

    return pd.DataFrame(
        [{"Material": mat, "Error": "Exists in ZPLM MARA but missing in MARA"} for mat in missing],
        columns=["Material", "Error"]
    )

def validar_marc_vs_zplm_marc_existence(marc_df, zplm_marc_df):
    """
    Verifica se TODAS as combinações (Material, Plant) da MARC existem na ZPLM MARC.
    Retorna DF de erros: Material | Plant | 
    """

    #Se faltar input, devolve DF vazio com colunas padrão
    if marc_df is None or marc_df.empty or zplm_marc_df is None or zplm_marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])
    
    # Normalizar colunas
    marc = marc_df.copy()
    marc.columns = [c.strip().upper() for c in marc.columns]

    zplm = zplm_marc_df.copy()
    zplm.columns = [c.strip().upper() for c in zplm.columns]

    #Pick columns helper
    def pick(df, candidates):
        for c in candidates:
            if c in df.columns:
                return c
        return None
    
    m_mat   = pick(marc, ["MATERIAL", "MATNR"])
    m_plant = pick(marc, ["PLANT", "WERKS"])

    z_mat   = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant = pick(zplm, ["PLANT", "WERKS"])

    if not all([m_mat, m_plant, z_mat, z_plant]):
        return pd.DataFrame(columns=["Material", "Plant", "Error"])
    
    #Normazilar valores
    marc[m_mat]   = marc[m_mat].astype(str).str.strip()
    marc[m_plant] = marc[m_plant].astype(str).str.strip().str.upper()

    zplm[z_mat]   = zplm[z_mat].astype(str).str.strip()
    zplm[z_plant] = zplm[z_plant].astype(str).str.strip().str.upper()

    #Conjuntos de valores
    marc_keys = set(zip(marc[m_mat], marc[m_plant]))
    zplm_keys = set(zip(zplm[z_mat], zplm[z_plant]))

    #Fatlas = MARC - ZPLM
    missing = marc_keys - zplm_keys

    if not missing:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])
    
    #Construir DF de erros
    errors = [{"Material": mat, "Plant": plant, "Error": "Missing in ZPLM MARC"} for (mat, plant) in sorted(missing)]

    return pd.DataFrame(errors, columns=["Material", "Plant", "Error"])

def validar_zplm_marc_vs_marc_existence(zplm_marc_df, marc_df):
    """
    Verifica combinações (Material, Plant) que existem em ZPLM MARC mas NÃO existem na MARC.
    Retorna: Material | Plant | Error
    """
    import pandas as pd

    if zplm_marc_df is None or zplm_marc_df.empty or marc_df is None or marc_df.empty:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    # Normalize headers
    zplm = zplm_marc_df.copy()
    zplm.columns = [str(c).strip().upper() for c in zplm.columns]

    marc = marc_df.copy()
    marc.columns = [str(c).strip().upper() for c in marc.columns]

    # Pick columns
    def pick(df, names):
        for n in names:
            if n in df.columns:
                return n
        return None

    z_mat   = pick(zplm, ["MATERIAL", "MATNR"])
    z_plant = pick(zplm, ["PLANT", "WERKS"])

    m_mat   = pick(marc, ["MATERIAL", "MATNR"])
    m_plant = pick(marc, ["PLANT", "WERKS"])

    if not all([z_mat, z_plant, m_mat, m_plant]):
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    # Normalize values
    zplm[z_mat]   = zplm[z_mat].astype(str).str.strip()
    zplm[z_plant] = zplm[z_plant].astype(str).str.strip().str.upper()

    marc[m_mat]   = marc[m_mat].astype(str).str.strip()
    marc[m_plant] = marc[m_plant].astype(str).str.strip().str.upper()

    z_keys = set(zip(zplm[z_mat], zplm[z_plant]))
    m_keys = set(zip(marc[m_mat], marc[m_plant]))

    missing = z_keys - m_keys

    if not missing:
        return pd.DataFrame(columns=["Material", "Plant", "Error"])

    return pd.DataFrame(
        [{"Material": mat, "Plant": plant, "Error": "Exists in ZPLM MARC but missing in MARC"} 
         for (mat, plant) in sorted(missing)],
        columns=["Material", "Plant", "Error"]
    )
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

    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None
    zplm_mara_df = load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

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
        prev_original_df = load_xlsx_fast(latest_report, sheet_name="Original Data")
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
        'Inspection Type': lambda mt, pt: mt in ['HALB', 'ROH'],
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
        if count > 0:
            coverage_data.append({
                'View': letter,
                'View_Description': view_to_desc.get(letter, ''),
                'Materiais without Active View': count,
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
            mt = str(mt_map.get(str(r['Material']).strip(), '')).strip().upper()
            pt = str(pt_map.get((str(r['Material']).strip(), str(r['Plant']).strip()), '')).strip().upper() if pt_map else ''
            present = set(r.get('Letters_Present', []) or [])
            if v not in present and is_view_required(mt, pt, v):
                rows.append({
                    'Material': str(r['Material']).strip(),
                    'Plant': str(r['Plant']).strip(),
                    'Missing_View': v,
                    'View_Description': view_to_desc.get(v, ''),
                    'Fields_Expected': ','.join(view_to_fields.get(v, []))
                })
        # 👉 NOVO: só adicionar se houver linhas
        if rows:
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
    qmat_file,
    roles_file=None,
    zplm_marc_file=None,
    zplm_mara_file=None,
    previous_report=None,
    plant_filter: str | None = None,
    progress_callback=None
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

    
    print("DEBUG TYPES:", {
        "sap": type(sap_file),
        "mara": type(mara_file),
        "marc": type(marc_file),
        "mvke": type(mvke_file),
        "zplm_marc": type(zplm_marc_file),
        "zplm_mara": type(zplm_mara_file),
        "qmat": type(qmat_file),
    })

    
    if progress_callback:
        progress_callback(5, "Loading SAPUI5 Export...")
    sap_df = load_xlsx_fast(sap_file, sheet_name="SAPUI5 Export")

    if progress_callback:
        progress_callback(7, "Loading MARA...")
    mara_df = load_xlsx_fast(mara_file, sheet_name=0)

    if progress_callback:
        progress_callback(10, "Loading MARC...")
    marc_df = load_xlsx_fast(marc_file, sheet_name=0)    

    if progress_callback:
        progress_callback(12, "Loading MVKE...")
    mvke_df = load_xlsx_fast(mvke_file, sheet_name=0)

    if progress_callback:
        progress_callback(14, "Loading QMAT...")
    qmat_df = load_xlsx_fast(qmat_file, sheet_name=0)

    # Fallback: se não houver roles_file, usa o Excel embutido via Base64
    if roles_file is None:
        roles_file = get_rolesviews_bytesio()  # BytesIO

    rolesviews_dict = load_xlsx_fast(roles_file, sheet_name=None)

    if progress_callback:
        progress_callback(16, "Loading ZPLM data...")
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

    if progress_callback:
        progress_callback(18, "Preparing ZPLM data...")         
    zplm_mara_df = load_xlsx_fast(zplm_mara_file, sheet_name=0) if zplm_mara_file else None

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
    if qmat_df is None:
        raise ValueError("QMAT.xlsx não carregado.")

    
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
        if progress_callback:
            progress_callback(20, f"Validating mandatory field: {coluna}...")
        validar_campo_obrigatorio(sap_df, coluna, resultados)

    # Chamadas em sequência
    if progress_callback:
        progress_callback(22, "Running field validations...")
    validar_procurement_type(sap_df, resultados)
    if progress_callback:
        progress_callback(24, "Running field validations...")
    validar_account_assignment_group(sap_df, resultados, mvke_df)
    if progress_callback:
        progress_callback(26, "Running field validations...")
    validar_automatic_purchase_order(sap_df, resultados)
    if progress_callback:
        progress_callback(28, "Running field validations...")
    validar_batch_management(sap_df, resultados)
    if progress_callback:
        progress_callback(30, "Running field validations...")
    validar_consumption_mode(sap_df, resultados)
    if progress_callback:
        progress_callback(32, "Running field validations...")   
    validar_division(sap_df, resultados)
    if progress_callback:   
        progress_callback(34, "Running field validations...")
    validar_export_import_group(sap_df, resultados)
    if progress_callback:
        progress_callback(36, "Running field validations...")
    validar_general_item_category_group(sap_df, resultados)
    if progress_callback:
        progress_callback(38, "Running field validations...")  
    validar_gross_weight(sap_df, resultados)
    if progress_callback:   
        progress_callback(40, "Running field validations...")
    validar_repetitive_manufacturing(sap_df, resultados)
    if progress_callback:   
        progress_callback(42, "Running field validations...")  
    validar_item_category_group(sap_df, resultados, mvke_df)
    if progress_callback:   
        progress_callback(44, "Running field validations...")  
    validar_jit_delivery_sched(sap_df, resultados)
    if progress_callback:   
        progress_callback(46, "Running field validations...")  
    validar_loading_group(sap_df, resultados)
    if progress_callback:   
        progress_callback(48, "Running field validations...")
    validar_lot_sizing_procedure(sap_df, resultados)
    if progress_callback:   
        progress_callback(50, "Running field validations...")
    validar_material_group(sap_df, resultados)
    if progress_callback:   
        progress_callback(52, "Running field validations...")
    validar_net_weight(sap_df, resultados)
    if progress_callback:   
        progress_callback(54, "Running field validations...")
    validar_period_type(sap_df, resultados)
    if progress_callback:   
        progress_callback(56, "Running field validations...")
    validar_planning_calendar(sap_df, resultados)
    if progress_callback:   
        progress_callback(58, "Running field validations...")  
    validar_planning_time_fence(sap_df, resultados)
    if progress_callback:   
        progress_callback(60, "Running field validations...")
    validar_mrp_type(sap_df, resultados, zplm_marc_df=zplm_marc_df)
    if progress_callback:   
        progress_callback(62, "Running field validations...")
    validar_price_control(sap_df, resultados)
    if progress_callback:   
        progress_callback(64, "Running field validations...")
    validar_price_unit(sap_df, resultados)
    if progress_callback:   
        progress_callback(66, "Running field validations...")
    validar_product_hierarchy(sap_df, resultados)
    if progress_callback:   
        progress_callback(68, "Running field validations...")
    validar_product_hierarchy_vs_enterproj(mara_df, zplm_mara_df, resultados)
    if progress_callback:   
        progress_callback(69, "Running field validations...")
    validar_production_scheduling_profile(sap_df, resultados)
    if progress_callback:   
        progress_callback(70, "Running field validations...")
    validar_profit_center(sap_df, resultados)
    if progress_callback:   
        progress_callback(72, "Running field validations...")
    validar_purchasing_group(sap_df, resultados)
    if progress_callback:   
        progress_callback(74, "Running field validations...")
    validar_repetitive_manufacturing_profile(sap_df, resultados)
    if progress_callback:   
        progress_callback(76, "Running field validations...")
    validar_source_list_required(sap_df, resultados)
    if progress_callback:   
        progress_callback(78, "Running field validations...")
    validar_storage_location(sap_df, resultados)
    if progress_callback:   
        progress_callback(80, "Running field validations...")
    validar_valuation_class(sap_df, resultados)
    if progress_callback:   
        progress_callback(82, "Running field validations...")
    validar_with_qty_structure(sap_df, resultados)
    if progress_callback:   
        progress_callback(84, "Running field validations...")
    validar_enterproj_number(sap_df, resultados)

    if progress_callback:   
        progress_callback(86, "Running field validations...")
    validar_inspection_type_qmat(sap_df, marc_df, qmat_df, resultados)
    if progress_callback:   
        progress_callback(88, "Running field validations...")
    validar_mrp_controller(sap_df, resultados)
    if progress_callback:   
        progress_callback(90, "Running field validations...")
    validar_plant_specific_material_status(sap_df, resultados)
    if progress_callback:   
        progress_callback(92, "Running field validations...")
    validar_backflush(sap_df, resultados)
    if progress_callback:   
        progress_callback(94, "Running field validations...")
    validar_strategy_group(sap_df, resultados)
    if progress_callback:   
        progress_callback(96, "Running field validations...")
    validar_product_related_origin(sap_df, resultados)
    if progress_callback:   
        progress_callback(98, "Running field validations...")
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
    
    if progress_callback:
        progress_callback(100, "Generating final Excel report...")
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

        
    
    def _style_adjusted_summary_sheet(xlsx_bytes: io.BytesIO) -> io.BytesIO:
        from openpyxl import load_workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        xlsx_bytes.seek(0)
        wb = load_workbook(xlsx_bytes)

        if "Adjusted Summary" not in wb.sheetnames:
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return out

        ws = wb["Adjusted Summary"]

        # Pintar tab azul bebé
        ws.sheet_properties.tabColor = "AEC6FF"

        # Estilos dos títulos
        title_fill = PatternFill(start_color="E8F0FF", end_color="E8F0FF", fill_type="solid")
        title_font = Font(bold=True)
        title_align = Alignment(horizontal="center", vertical="center")
        thin = Side(style="thin", color="BBD0FF")
        title_border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # -------------------------------------------------------------------
        # 1) DETETAR PRIMEIRA TABELA (Field Errors)
        # -------------------------------------------------------------------
        first_header_row = None
        first_empty_after_table1 = None

        # a) Encontrar a primeira linha NÃO vazia = cabeçalho da tabela 1
        for r in range(1, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                first_header_row = r
                break

        if first_header_row is None:
            # Nada para estilizar
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            return out

        # b) Encontrar a primeira linha totalmente vazia depois da tabela 1
        for r in range(first_header_row + 1, ws.max_row + 1):
            if all(ws.cell(row=r, column=c).value in (None, "") for c in range(1, ws.max_column + 1)):
                first_empty_after_table1 = r
                break

        # -------------------------------------------------------------------
        # 2) DETETAR SEGUNDA TABELA (Coverage)
        # -------------------------------------------------------------------
        second_header_row = None

        if first_empty_after_table1:
            for r in range(first_empty_after_table1 + 1, ws.max_row + 1):
                if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                    second_header_row = r
                    break

        # -------------------------------------------------------------------
        # 3) Criar TÍTULO 1 (Field Errors)
        # -------------------------------------------------------------------
        title_row_1 = first_header_row - 1
        last_col_1 = max(
            c for c in range(1, ws.max_column + 1)
            if ws.cell(row=first_header_row, column=c).value not in (None, "")
        )

        if title_row_1 >= 1:
            ws.merge_cells(start_row=title_row_1, start_column=1,
                        end_row=title_row_1, end_column=last_col_1)
            cell = ws.cell(row=title_row_1, column=1)
            cell.value = "Field Errors"
            cell.fill = title_fill
            cell.font = title_font
            cell.alignment = title_align
            cell.border = title_border

        # -------------------------------------------------------------------
        # 4) Criar TÍTULO 2 (Materials Without View)
        # -------------------------------------------------------------------
        if second_header_row:
            title_row_2 = second_header_row - 1
            last_col_2 = max(
                c for c in range(1, ws.max_column + 1)
                if ws.cell(row=second_header_row, column=c).value not in (None, "")
            )

            ws.merge_cells(start_row=title_row_2, start_column=1,
                        end_row=title_row_2, end_column=last_col_2)
            cell2 = ws.cell(row=title_row_2, column=1)
            cell2.value = "Materials Without View"
            cell2.fill = title_fill
            cell2.font = title_font
            cell2.alignment = title_align
            cell2.border = title_border

        # -------------------------------------------------------------------
        # Guardar Excel formatado
        # -------------------------------------------------------------------
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out

    file_bytes.seek(0)
    file_bytes = _style_adjusted_summary_sheet(file_bytes)

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
    df_auto   = load_xlsx_fast(auto_file, sheet_name=0)      # "PLM Automotive Validation.xlsx"
    df_plm    = load_xlsx_fast(plm_file, sheet_name=0)       # "PLM TimingLetter Validation.xlsx"
    df_timing = load_xlsx_fast(timing_file, sheet_name=0)    # "Timing Letters.xlsx"

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

# ==========================================================
# SAP MAterial Comparisson 
# ==========================================================

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
        rolesviews_dict = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
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
    current_dict  = load_xlsx_fast(current_file, sheet_name=None)
    previous_dict = load_xlsx_fast(previous_file, sheet_name=None)

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
# REPORTS VAULT (guardar Excel gerado) + payload helper
# ==========================================================

def save_report_to_vault(excel_bytes: io.BytesIO, *, report_type: str = "sap", plant_suffix: str = "ALL") -> str:
    """
    Grava o Excel gerado num diretório local tipo ./reports/<report_type>/YYYY-MM/
    Retorna o caminho do ficheiro gravado.
    """
    # Diretório base
    base_dir = os.path.join(".", "reports", report_type)
    # Subpasta por mês (organização simples)
    month_dir = datetime.now().strftime("%Y-%m")
    out_dir = os.path.join(base_dir, month_dir)
    os.makedirs(out_dir, exist_ok=True)

    # Nome do ficheiro
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"{ts}_SAP_Validation_Adjusted_Report_{plant_suffix}.xlsx"
    fpath = os.path.join(out_dir, fname)

    # Escrever bytes
    excel_bytes.seek(0)
    with open(fpath, "wb") as f:
        f.write(excel_bytes.getvalue())

    return os.path.abspath(fpath)

def build_sap_log_payload(
    *,
    user: str,
    plant_filter: str | None,
    sap_file,
    mara_file,
    marc_file,
    mvke_file,
    errors_df,
    summary_df,
    report_path: str | None = None,   # <-- NOVO (opcional)
):
    """
    Mantém API compatível; acrescenta caminho do relatório no payload quando fornecido.
    """
    from datetime import datetime

    def _name(f, default):
        return getattr(f, "name", default) if f is not None else default

    total_errors = int(errors_df.shape[0]) if errors_df is not None else 0
    total_fields = int(errors_df['Field'].nunique()) if (errors_df is not None and 'Field' in errors_df.columns) else 0
    total_materials = int(errors_df['Material'].nunique()) if (errors_df is not None and 'Material' in errors_df.columns) else 0

    payload = {
        "run_type": "SAP_VALIDATION",
        "type": "sap",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "user": user,
        "plant_filter": plant_filter if plant_filter else "ALL",
        "metrics": {
            "total_errors": total_errors,
            "fields_with_error": total_fields,
            "materials_with_error": total_materials,
        },
        "files": {
            "sap": _name(sap_file, "Item Browse Report.xlsx"),
            "mara": _name(mara_file, "MARA.xlsx"),
            "marc": _name(marc_file, "MARC.xlsx"),
            "mvke": _name(mvke_file, "MVKE.xlsx"),
        },
    }

    # Novo: anexar caminho do relatório gerado (se existir)
    if report_path:
        payload["report_path"] = report_path

    return payload

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
    mara_df = load_xlsx_fast(mara_file, sheet_name=0)
    marc_df = load_xlsx_fast(marc_file, sheet_name=0)
    zplm_marc_df = load_xlsx_fast(zplm_marc_file, sheet_name=0) if zplm_marc_file else None

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
    rolesviews_dict = load_xlsx_fast(get_rolesviews_bytesio(), sheet_name=None)
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

# ==========================================================
# DATA CLEANUPS (MARA + MARC + ZPLM MARC + ZPLM MARA)
# ==========================================================

def _filter_materials_by_base_duplicates(df: pd.DataFrame, material_col: str = "Material") -> pd.DataFrame:
    """
    Mantém no DataFrame apenas as linhas cujo 'Material' (depois de remover o último char)
    forme uma 'base' com pelo menos 2 ocorrências no scope atual.
    - Não altera o 'Material' original (usa uma coluna temporária).
    """
    if df is None or df.empty or material_col not in df.columns:
        return df

    tmp = df.copy()
    # cria a chave base removendo o último carácter; se string vazia, fica vazia
    tmp["_base"] = tmp[material_col].astype(str).str.strip().str[:-1]

    # grupos com contagem >= 2
    base_counts = tmp["_base"].value_counts(dropna=False)
    bases_keep = set(base_counts[base_counts >= 2].index)

    # filtra apenas bases com duplicados
    out = tmp[tmp["_base"].isin(bases_keep)].drop(columns=["_base"])
    return out

def _pick_col_upper(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _normalize_code_as_str_basic(x):
    import pandas as pd
    if pd.isna(x):
        return ""
    s = str(x).strip()
    try:
        return str(int(float(s)))
    except Exception:
        return s

def _classify_fp_versions(df):
    """
    Marca versões por família: tudo igual exceto o último carácter.
    Escopo do grupo: (Plant, Base).
    Regras:
      - Ordena por LastChar (A->Z) dentro do grupo.
      - Última é 'Correct'; restantes 'Outdated'.
    """
    import pandas as pd
    if df.empty:
        df["Status"] = []
        return df

    work = df.copy()
    work["Material"] = work["Material"].astype(str).str.strip()
    # Evita materiais com comprimento < 1
    work = work[work["Material"].str.len() > 0].copy()

    work["Base"] = work["Material"].str[:-1]
    work["LastChar"] = work["Material"].str[-1]

    # Status default
    work["Status"] = "Outdated"
    # Para cada (Plant, Base), encontrar a maior LastChar
    for (plant, base), grp in work.groupby(["Plant", "Base"]):
        if len(grp) == 0:
            continue
        # Ordena lexicograficamente pela última letra; última é a mais recente
        idx_latest = grp.sort_values("LastChar").index[-1]
        work.loc[idx_latest, "Status"] = "Correct"

    return work.drop(columns=["Base", "LastChar"])

def run_fp_status_validation(
    mara_file,
    marc_file,
    zplm_marc_file,
    zplm_mara_file,  # mantido para compatibilidade
    plant_filter=None,
):
    """
    FP – Regras finais (Rodrigo), agora em modo "clean":
    - devolve apenas df_fp e kpis
    - não exporta Excel
    - não cria sheets
    """

    import pandas as pd

    # -----------------------------
    # 1) LOAD FILES
    # -----------------------------
    mara = load_xlsx_fast(mara_file, sheet_name=0)
    marc = load_xlsx_fast(marc_file, sheet_name=0)
    zplm = load_xlsx_fast(zplm_marc_file, sheet_name=0)

    if mara is None or marc is None or zplm is None:
        raise ValueError("Missing required inputs")

    # -----------------------------
    # 2) NORMALIZE HEADERS
    # -----------------------------
    mara.columns = mara.columns.astype(str).str.strip().str.upper()
    marc.columns = marc.columns.astype(str).str.strip().str.upper()
    zplm.columns = zplm.columns.astype(str).str.strip().str.upper()

    # -----------------------------
    # 3) PICK COLUMNS
    # -----------------------------
    col_mat_mara     = _pick_col_upper(mara, ["MATERIAL", "MATNR"])
    col_mat_marc     = _pick_col_upper(marc, ["MATERIAL", "MATNR"])
    col_plant_marc   = _pick_col_upper(marc, ["PLANT", "WERKS"])
    col_psms         = _pick_col_upper(marc, [
        "PLANT-SPECIFIC MATERIAL STATUS", "PLANT SPECIFIC MATERIAL STATUS",
        "PLANT SPEC. MATERIAL STATUS", "MTSTB"
    ])
    col_mat_zplm     = _pick_col_upper(zplm, ["MATERIAL", "MATNR"])
    col_plant_zplm   = _pick_col_upper(zplm, ["PLANT", "WERKS"])
    col_mptype       = _pick_col_upper(zplm, [
        "MATERIAL PLANT TYPE", "PLANT TYPE", "MATERIAL-PLANT TYPE"
    ])

    # -----------------------------
    # 4) PLANT FILTER
    # -----------------------------
    pf = str(plant_filter).strip().upper() if plant_filter else None

    # -----------------------------
    # 5) MARA BASE SET
    # -----------------------------
    mara_set = set(mara[col_mat_mara].dropna().astype(str).str.strip())

    # -----------------------------
    # 6) ZPLM → FP ROWS
    # -----------------------------
    df_fp = zplm[
        zplm[col_mptype].astype(str).str.strip().str.upper() == "FP"
    ][[col_mat_zplm, col_plant_zplm, col_mptype]].copy()

    df_fp[col_mat_zplm]   = df_fp[col_mat_zplm].astype(str).str.strip()
    df_fp[col_plant_zplm] = df_fp[col_plant_zplm].astype(str).str.strip().str.upper()

    df_fp = df_fp[df_fp[col_mat_zplm].isin(mara_set)]

    if pf:
        df_fp = df_fp[df_fp[col_plant_zplm] == pf]

    # -----------------------------
    # 7) MARC → PSMS
    # -----------------------------
    marc_small = marc[[col_mat_marc, col_plant_marc, col_psms]].copy()

    marc_small[col_mat_marc]   = marc_small[col_mat_marc].astype(str).str.strip()
    marc_small[col_plant_marc] = marc_small[col_plant_marc].astype(str).str.strip().str.upper()
    marc_small[col_psms]       = marc_small[col_psms].apply(_normalize_code_as_str_basic)

    allowed = {"40", "45", "50", "55"}
    marc_small = marc_small[marc_small[col_psms].isin(allowed)]

    if pf:
        marc_small = marc_small[marc_small[col_plant_marc] == pf]

    merged = df_fp.merge(
        marc_small,
        left_on=[col_mat_zplm, col_plant_zplm],
        right_on=[col_mat_marc, col_plant_marc],
        how="inner"
    )

    # -----------------------------
    # 8) FINAL DF
    # -----------------------------
    df = merged.rename(columns={
        col_mat_zplm:   "Material",
        col_plant_zplm: "Plant",
        col_mptype:     "Material Plant Type",
        col_psms:       "PSMS",
    })[["Material", "Plant", "Material Plant Type", "PSMS"]]

    df["Material"] = df["Material"].astype(str).str.strip()
    df = df[~df["Material"].str.match(r"^\d+$")]

    # -----------------------------
    # 9) ADD PSMS NAME
    # -----------------------------
    PSMS_NAME = {"40": "NewProd", "45": "Prod", "50": "Balout", "55": "Service"}
    df["PSMS Name"] = df["PSMS"].map(PSMS_NAME)

    # -----------------------------
    # 10) BUILD FAMILY BASE
    # -----------------------------
    df["_base"] = df["Material"].str[:-1]
    base_counts = df["_base"].value_counts()
    keep = set(base_counts[base_counts >= 2].index)
    df = df[df["_base"].isin(keep)].copy()

    # -----------------------------
    # 11) CLASSIFICATION LOGIC
    # -----------------------------
    def classify_family(g: pd.DataFrame) -> pd.DataFrame:
        g = g.copy()

        g["suffix"] = g["Material"].str[-1]
        g = g.sort_values("suffix")

        psms_set = set(g["PSMS"])

        if psms_set == {"55"} and len(g) >= 2:
            g["Status"] = "Need verify"
            return g.drop(columns=["suffix"])

        priority = ["45", "40", "55", "50"]
        correct_psms = next((p for p in priority if p in psms_set), None)

        g["Status"] = "Correct"

        has40 = "40" in psms_set
        has45 = "45" in psms_set
        has55 = "55" in psms_set

        if correct_psms in ("45", "40") and has55:
            g.loc[g["PSMS"] == "55", "Status"] = "Need verify"
        if correct_psms == "55" and (has40 or has45):
            g.loc[g["PSMS"].isin(["40", "45"]), "Status"] = "Need verify"

        if correct_psms == "45" and has40:
            g.loc[g["PSMS"] == "40", "Status"] = "Need verify"
        if correct_psms == "40" and has45:
            g.loc[g["PSMS"] == "45", "Status"] = "Need verify"

        dup50 = g[g["PSMS"] == "50"]
        if len(dup50) >= 2:
            oldest = dup50.index[0]
            g.loc[oldest, "Status"] = "Outdated"

        return g.drop(columns=["suffix"])

    df = df.groupby(["Plant", "_base"], group_keys=False).apply(classify_family)
    df = df.drop(columns=["_base"])

    # -----------------------------
    # 12) KPIs
    # -----------------------------
    if df.empty:
        kpis = {"materials": 0, "plants": 0, "psms": 0, "outdated": 0, "correct": 0}
    else:
        kpis = {
            "materials": df["Material"].nunique(),
            "plants":    df["Plant"].nunique(),
            "psms":      df["PSMS"].nunique(),
            "outdated":  (df["Status"] == "Outdated").sum(),
            "correct":   (df["Status"] == "Correct").sum(),
        }

    # -----------------------------
    # RETURN CLEAN RESULTS
    # -----------------------------
    return df, kpis

def run_data_cleanups_validation(
    mara_file,
    marc_file,
    zplm_marc_file,
    zplm_mara_file,
    mbew_file,
    plant_filter=None,
):
    """
    DATA CLEANUPS ENGINE
    --------------------
    Este motor consolida todas as validações da página “Data Cleanups”
    num único Excel multi-sheet.

    OUTPUT:
        - excel_bytes : Excel com N sheets (uma por validação)
        - results     : dict {sheet_name : dataframe_resultado}
        - kpis        : KPIs da validação FP (pode ser expandido)
    """

    import pandas as pd
    import io

    # -----------------------------------------------------
    # 1. Carregar ficheiros
    # -----------------------------------------------------
    mara       = load_xlsx_fast(mara_file, sheet_name=0)
    marc       = load_xlsx_fast(marc_file, sheet_name=0)
    zplm_marc  = load_xlsx_fast(zplm_marc_file, sheet_name=0)
    zplm_mara  = load_xlsx_fast(zplm_mara_file, sheet_name=0)
    mbew = load_xlsx_fast(mbew_file, sheet_name=0)

    if mara is None or marc is None or zplm_marc is None:
        raise ValueError("Missing required files for Data Cleanups")

    # Dicionário que vai conter todas as sheets
    results = {}

    # =====================================================
    # 2. Validação principal FP Status  (o antigo FP Extract)
    # =====================================================
    df_fp, kpi_fp = run_fp_status_validation(
        mara_file=mara_file,
        marc_file=marc_file,
        zplm_marc_file=zplm_marc_file,
        zplm_mara_file=zplm_mara_file,
        plant_filter=plant_filter
    )

    results["FP Level Status"] = df_fp

    # =====================================================
    # 3. Validação FP PT01 + PT02 + Status 55
    # =====================================================
    df_status55 = validar_fp_pt01_pt02_status55(zplm_marc, marc)
    results["FP in PT02"] = df_status55

    # =====================================================
    # 4. OUTRAS VALIDAÇÕES (exemplo: EOP vs Service Liability)
    # =====================================================
    df_eop_vs_srv = validar_fp55_eop_vs_service_liability(zplm_marc, marc)
    results["EOP vs Service Liability (Blank or Swapped)"] = df_eop_vs_srv

    df_srv_expired = validar_fp55_service_liability_expired(zplm_marc, marc)
    results["Past Service Liability"] = df_srv_expired

    df_std_price = validar_fppp70_standard_price_mbew(zplm_marc, marc, mbew)
    results["No Standard Cost"] = df_std_price

    # --- Cobertura: MARC deve existir em ZPLM MARC (por planta) ---
    df_marc_cov = validar_marc_vs_zplm_marc_existence(marc, zplm_marc)
    results["Missing in ZPLM_MARC"] = df_marc_cov

    # --- Cobertura: MARA deve existir em ZPLM MARA (só material; sem Plant) ---
    df_mara_cov = validar_mara_vs_zplm_mara_existence(mara, zplm_mara)
    results["Missing in ZPLM_MARA"] = df_mara_cov

    # --- Inverse: ZPLM MARC should exist in MARC ---
    df_zplm_marc_cov = validar_zplm_marc_vs_marc_existence(zplm_marc, marc)
    results["Missing in MARC"] = df_zplm_marc_cov

    # --- Inverse: ZPLM MARA should exist in MARA (sem Plant) ---
    df_zplm_mara_cov = validar_zplm_mara_vs_mara_existence(zplm_mara, mara)
    results["Missing in MARA"] = df_zplm_mara_cov

    # =======================================================
    # 5) WRITE THE EXCEL (SUMMARY + VALIDATION SHEETS)
    # =======================================================
    out = io.BytesIO()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:

        # =======================================================
        # SUMMARY SHEET
        # =======================================================
        summary_rows = []

        for sheet_name, df_val in results.items():

            # No "Plant" column → treat as ALL
            if df_val is None or df_val.empty or "Plant" not in df_val.columns:
                summary_rows.append({
                    "Validation": sheet_name,
                    "Plant": "ALL",
                    "Issues": 0
                })
                continue

            # Per-plant summary
            for plant in ["PT01", "PT02"]:
                dfp = df_val[df_val["Plant"] == plant]

                summary_rows.append({
                    "Validation": sheet_name,
                    "Plant": plant,
                    "Issues": len(dfp),
                })

        df_summary = pd.DataFrame(summary_rows)
        df_summary.to_excel(writer, sheet_name="SUMMARY", index=False)

        # =======================================================
        # VALIDATION SHEETS (with plant split when applicable)
        # =======================================================
        for sheet_name, df_validation in results.items():

            # Empty DF → still create PT01 + PT02 sheets (se a validação for por planta)
            if df_validation is None or df_validation.empty:
                # Se não tiver coluna Plant: cria uma única sheet 'sheet_name'
                if df_validation is None or "Plant" not in (df_validation.columns if df_validation is not None else []):
                    safe_name = sheet_name[:31]
                    pd.DataFrame(columns=["Material"]).to_excel(writer, sheet_name=safe_name, index=False)
                    continue

                # Tem Plant → cria PT01/PT02 vazias
                for plant in ["PT01", "PT02"]:
                    safe_name = f"{sheet_name}_{plant}"[:31]
                    pd.DataFrame(columns=["Material", "Plant"]).to_excel(
                        writer, sheet_name=safe_name, index=False
                    )
                continue

            # ⚠️ Caso a validação NÃO tenha coluna Plant → escrever numa única sheet
            if "Plant" not in df_validation.columns:
                safe_name = sheet_name[:31]
                df_validation.to_excel(writer, sheet_name=safe_name, index=False)
                continue

            # IF FILTERED → ONLY ONE PLANT
            if plant_filter and plant_filter in ("PT01", "PT02"):
                safe_name = f"{sheet_name}_{plant_filter}"[:31]
                df_validation[df_validation["Plant"] == plant_filter].to_excel(
                    writer, sheet_name=safe_name, index=False
                )
                continue

            # ALL PLANTS → GENERATE PT01 + PT02
            for plant in ["PT01", "PT02"]:
                dfp = df_validation[df_validation["Plant"] == plant]
                safe_name = f"{sheet_name}_{plant}"[:31]

                if dfp.empty:
                    pd.DataFrame(columns=df_validation.columns).to_excel(
                        writer, sheet_name=safe_name, index=False
                    )
                else:
                    dfp.to_excel(writer, sheet_name=safe_name, index=False)

    # =======================================================
    # 2) APPLY STYLING + HYPERLINKS TO SUMMARY SHEET
    # =======================================================
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import quote_sheetname

    out.seek(0)
    wb = load_workbook(out)

    if "SUMMARY" in wb.sheetnames:
        ws = wb["SUMMARY"]

        # --- Add hyperlink column ---
        link_col = ws.max_column + 1
        ws.cell(row=1, column=link_col, value="Link")

        # Styles
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        zebra_fill  = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
        border = Border(
            left=Side(style="thin", color="808080"),
            right=Side(style="thin", color="808080"),
            top=Side(style="thin", color="808080"),
            bottom=Side(style="thin", color="808080")
        )

        # Format header
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        # Build hyperlinks row-by-row
        for r in range(2, ws.max_row + 1):

            validation = ws.cell(row=r, column=1).value.strip()
            plant      = ws.cell(row=r, column=2).value.strip() if ws.cell(row=r, column=2).value else None

            # ===================================================
            # DETECT REAL TAB NAME CREATED IN THE ENGINE
            # ===================================================
            
            # CASE 1: Validation without Plant → tab is exactly validation name
            if plant not in ("PT01", "PT02"):
                target_sheet = validation
            
            # CASE 2: Validation with Plant → tab = validation + "_" + PLANT
            else:
                target_sheet = f"{validation}_{plant}"

            # If sheet exists → add hyperlink
            if target_sheet in wb.sheetnames:
                ws.cell(row=r, column=link_col).value = f"=HYPERLINK(\"#'{quote_sheetname(target_sheet)}'!A1\", \"Open\")"
            else:
                ws.cell(row=r, column=link_col).value = "Not found"

            # Format hyperlink cell
            link_cell = ws.cell(row=r, column=link_col)
            link_cell.font = Font(color="0563C1", underline="single")
            link_cell.alignment = Alignment(horizontal="center")
            link_cell.border = border

        # Zebra + borders + autofit
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                if r % 2 == 0:
                    cell.fill = zebra_fill
                cell.border = border

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                val = str(cell.value) if cell.value else ""
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max_len + 2

    # FINAL SAVE
    out2 = io.BytesIO()
    wb.save(out2)
    out2.seek(0)

    return out2, results, kpi_fp

# ===========================
# Logging payload (opcional)
# ===========================

def build_fp_log_payload(
    user: str,
    plant_filter: str | None,
    mara_file, marc_file, zplm_marc_file, zplm_mara_file,
    df_result
):
    import os
    mats = int(df_result["Material"].nunique()) if df_result is not None and not df_result.empty else 0
    plants = int(df_result["Plant"].nunique()) if df_result is not None and not df_result.empty else 0
    outdated = int((df_result["Status"] == "Outdated").sum()) if df_result is not None and not df_result.empty else 0
    correct = int((df_result["Status"] == "Correct").sum()) if df_result is not None and not df_result.empty else 0

    return {
        "kind": "fp",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user": user,
        "plant_filter": plant_filter,
        "files": {
            "mara": getattr(mara_file, "name", "MARA.xlsx"),
            "marc": getattr(marc_file, "name", "MARC.xlsx"),
            "zplm_marc": getattr(zplm_marc_file, "name", "ZPLM MARC.xlsx"),
            "zplm_mara": getattr(zplm_mara_file, "name", "ZPLM MARA.xlsx"),
        },
        "metrics": {
            "materials": mats,
            "plants": plants,
            "outdated": outdated,
            "correct": correct,
        }
    }