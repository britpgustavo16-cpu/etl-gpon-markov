import os
import re
import glob
import time
import zipfile
import unicodedata
import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if os.name == "nt" and os.path.exists(r"C:\Program Files\Tesseract-OCR\tesseract.exe"):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PASTA_ENTRADA = os.path.join(BASE_DIR, "entrada_materiais")
ANO_FILTRO = 2025
ARQUIVO_SAIDA = os.path.join(BASE_DIR, f"Mestre_GPON_Markov_{ANO_FILTRO}.xlsx")

# Ignora contagens contábeis de balanço físico que não representam movimentação real
IGNORAR_INVENTARIO_CONTABIL = True

# Regra operacional: atualmente apenas equipamentos da família 168 possuem fluxo de RMA
MODELOS_RMA_PERMITIDOS = {"168", "R168", "E168"}

DESCRICAO_ESTADOS = {
    "S0": "S0 - Estoque Novo / Bom (Gradeado)",
    "S1": "S1 - Expedição Campo / Equipe Técnica",
    "S2": "S2 - Retorno de Campo (Material RET)",
    "S3": "S3 - Laboratório / Triagem Técnica",
    "S4": "S4 - Conversão de Material (RET/BOM)",
    "S5": "S5 - Estoque Reaproveitado (Reuso)",
    "S6": "S6 - Defeito / Reprovado (Material DEF)",
    "S7": "S7 - RMA / Fornecedor (Exclusivo Mod. 168)",
    "S8": "S8 - Quarentena Obsoletos / Descarte",
    "S9": "S9 - Sucata / Baixa Definitiva",
}

PRECEDENCIA_ESTADO = {
    "S9": 10,
    "S8": 9,
    "S7": 8,
    "S6": 7,
    "S4": 6,
    "S3": 5,
    "S2": 4,
    "S1": 3,
    "S5": 2,
    "S0": 1,
}

ESTADOS_ABSORVENTES = {"S9"}

COLUNAS_PADRAO = [
    "GPON / Serial",
    "Cod. Material",
    "Dt. Movimento",
    "Dia Dig.",
    "Hora Dig.",
    "Contr / Proj Origem",
    "Sentido Mov.",
    "Nº. Documento",
    "Situação Mat.",
    "Posição",
    "Estado_Markov",
    "Proximo_Estado",
    "Dias_no_Estado",
    "Horas_no_Estado",
    "Status_Transicao",
    "Responsável",
    "Observação",
    "Usuário Dig.",
]

PATTERN_GPON_STR = (
    r"\b(?:ZTE[A-Z0-9]{8,12}|FHTT[A-Z0-9]{8,12}|EF[A-Z0-9]{10,14}|"
    r"(?:0525|CC06|D80D|DD16|68FF|C471|5450)[A-Z0-9]{6,12}|"
    r"(?:[0-9A-F]{2}[:-]){5}[0-9A-F]{2}|[0-9A-F]{12,16})\b"
)
REGEX_GPON = re.compile(PATTERN_GPON_STR, re.IGNORECASE)

TERMOS_PASSIVOS = [
    "CABO", "DROP", "CONECTOR", "APC", "UPC", "PTO", "ALCA", "ALÇA",
    "ROLDANA", "FITA", "ESTICADOR", "FECHO", "ABRACADEIRA", "ABRAÇADEIRA",
    "TUBETE", "SPLITTER", "CAIXA EMENDA", "CEO", "CTO", "PARAFUSO", "BUCHA",
    "ANEL DE DISTRIBUICAO", "CORDA 10MM"
]

ABAS_IGNORAR = {
    "MANUAL DE USO",
    "DASH GERAL",
    "DASH SEGMENTADO",
    "TABELA GERAL",
    "TABELA SEGMENTADO",
    "BASE",
    "ANALSE DIARIA POR COLABORADOR",
    "ANALISE DIARIA POR COLABORADOR",
    "TABELA DINAMICA 8",
    "TABELA DINAMICA 6",
    "1 - ENTRADAS E SAIDAS",
    "3 - SAIDA PARA LOJAS",
    "4 - ADESIVOCAPA",
    "2.2 - COD. DE MATERIAL",
    "6 - RECOLHIMENTO DE COMODATO",
}


def normalizar_texto(texto) -> str:
    if pd.isna(texto) or texto is None:
        return ""
    s = str(texto).strip().upper()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    return s


def eh_passivo(linha_concatenada: str) -> bool:
    norm = normalizar_texto(linha_concatenada)
    return any(termo in norm for termo in TERMOS_PASSIVOS)


def eh_modelo_168(cod_mat: str) -> bool:
    """Verifica se o código pertence à família 168 (168, R168, E168), única elegível para RMA."""
    c = normalizar_texto(cod_mat)
    if c in MODELOS_RMA_PERMITIDOS:
        return True
    return bool(re.match(r"^[A-Z]?168$", c))


def serial_valido(serial: str) -> bool:
    if not serial:
        return False
    s = re.sub(r"[^A-Z0-9]", "", str(serial).upper())
    if not (12 <= len(s) <= 16):
        return False
    if len(set(s)) <= 1:
        return False
    return True


def extrair_gpon(texto: str) -> str:
    if not texto or pd.isna(texto):
        return ""
    match = REGEX_GPON.search(str(texto).upper())
    if not match:
        return ""
    candidato = match.group(0).replace(":", "").replace("-", "")
    return candidato if serial_valido(candidato) else ""


def excel_num_to_date_str(val_str: str) -> str:
    try:
        num = float(val_str)
        if 20000 <= num <= 60000:
            base = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=num)
            return base.strftime("%d/%m/%Y")
    except ValueError:
        pass
    return val_str


def col_letter_to_idx(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref)
    if not m:
        return 0
    idx = 0
    for ch in m.group(1):
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def padronizar_coluna_data(serie: pd.Series) -> pd.Series:
    valores = serie.fillna("").astype(str).str.strip().tolist()

    eh_americano = False
    for v in valores:
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})", v)
        if m:
            p1, p2 = int(m.group(1)), int(m.group(2))
            if p1 <= 12 and p2 > 12:
                eh_americano = True
                break

    datas_padronizadas = []
    for v in valores:
        if not v or v.upper() in {"#REF!", "#N/A", "NAT", "NONE", "NAN", "-"}:
            datas_padronizadas.append("")
            continue

        m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", v)
        if m_iso:
            ano, mes, dia = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
            try:
                datas_padronizadas.append(datetime.date(ano, mes, dia).strftime("%d/%m/%Y"))
            except ValueError:
                datas_padronizadas.append("")
            continue

        m_bar = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})", v)
        if m_bar:
            p1, p2, ano_raw = int(m_bar.group(1)), int(m_bar.group(2)), m_bar.group(3)
            ano = 2000 + int(ano_raw) if len(ano_raw) == 2 else int(ano_raw)
            if eh_americano or (p1 <= 12 and p2 > 12):
                mes, dia = p1, p2
            else:
                dia, mes = p1, p2
            try:
                datas_padronizadas.append(datetime.date(ano, mes, dia).strftime("%d/%m/%Y"))
            except ValueError:
                datas_padronizadas.append("")
            continue

        datas_padronizadas.append("")

    return pd.Series(datas_padronizadas, index=serie.index)


def parse_timestamp_padrao(dt_str: str, dia_dig_str: str, hora_str: str) -> pd.Timestamp:
    candidato = dt_str if dt_str else dia_dig_str
    if not candidato:
        return pd.NaT

    hora_limpa = str(hora_str).strip()
    if not re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", hora_limpa):
        hora_limpa = "00:00:00"

    return pd.to_datetime(f"{candidato} {hora_limpa}", format="%d/%m/%Y %H:%M:%S", errors="coerce")


def ajustar_cabecalho_real(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    palavras_chave = {
        "GPON", "COD. GPON", "COD GPON", "SERIAL", "NUM SERIE",
        "MAC", "EQUIPAMENTO", "COD.COMPLEM.", "COD COMPLEM"
    }
    cols_norm = {normalizar_texto(c) for c in df.columns}
    if cols_norm & palavras_chave:
        return df.reset_index(drop=True)

    for i in range(min(12, len(df))):
        valores_linha = [str(v).strip() if pd.notna(v) else "" for v in df.iloc[i].values]
        valores_norm = {normalizar_texto(v) for v in valores_linha}
        if valores_norm & palavras_chave:
            novas_cols = [v if v else f"COL_VAZIA_{idx}" for idx, v in enumerate(valores_linha)]
            df_limpo = df.iloc[i + 1:].copy()
            df_limpo.columns = novas_cols
            return df_limpo.reset_index(drop=True)

    return df.reset_index(drop=True)


def mapear_colunas_brutas(df: pd.DataFrame, origem_arquivo: str = "", nome_aba: str = "") -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

    df = ajustar_cabecalho_real(df)
    if df.empty:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

    mapa_aliases = {
        "GPON / Serial": [
            "GPON", "COD. GPON", "COD GPON", "COD.COMPLEM.", "COD COMPLEM",
            "COD. COMPLEM", "SERIAL", "NUM SERIE", "NUMERO DE SERIE", "MAC", "EQUIPAMENTO", "SN"
        ],
        "Cod. Material": ["COD.MATER.", "COD MATER", "COD. MAT.", "COD MAT", "COD. MAT", "CODIGO MATERIAL"],
        "Dt. Movimento": ["DATA MOVTO.", "DATA MOVTO", "DT. MOVIMENTO", "DT MOVIMENTO", "DATA MOVIMENTO", "DATA", "DT MOV"],
        "Dia Dig.": ["DATA BAIXA", "DIA DIG.", "DIA DIG", "DATA DIGITACAO", "DT DIGITACAO"],
        "Hora Dig.": ["HORA DA BAIXA", "HORA DIG.", "HORA DIG", "HORA", "HR DIGITACAO"],
        "Contr / Proj Origem": ["PROJETO", "CONTR / PROJ ORIGEM", "CONTRATO", "PROJ ORIGEM", "CONTR/PROJ"],
        "Nº. Documento": ["NRM/DM", "NºRM/DM", "Nº. DOCUMENTO", "N. DOCUMENTO", "DOCUMENTO", "NTRANSF.", "NºTRANSF.", "DOCUM."],
        "Situação Mat.": ["SITUAÇÃO MAT.", "SITUACAO MAT", "STATUS DO TESTE", "CONDIÇÃO DE USO", "CONDICAO DE USO", "SITUACAO"],
        "Posição": ["POSIÇÃO", "POSICAO", "LOCAL", "STATUS POSICAO"],
        "Responsável": ["RESPONSÁVEL", "RESPONSAVEL", "RESPONSAVÉL", "TECNICO", "COLABORADOR"],
        "Observação": ["OBSERVAÇÃO", "OBSERVACAO", "OBS", "AVARIA", "DEFEITO", "TIPO DE DEFEITO", "RELATO DA OCORRENCIA"],
        "Usuário Dig.": ["USUÁRIO QUE BAIXOU", "USUARIO QUE BAIXOU", "USUÁRIO DIG.", "USUARIO DIG", "USUARIO"],
    }

    df_out = pd.DataFrame(index=df.index)
    colunas_originais = {}
    for c in df.columns:
        c_norm = normalizar_texto(c)
        if c_norm and c_norm not in colunas_originais:
            colunas_originais[c_norm] = c

    for col_alvo, aliases in mapa_aliases.items():
        encontrada = next((colunas_originais[a] for a in aliases if a in colunas_originais), None)
        if encontrada is not None:
            serie = df[encontrada]
            if isinstance(serie, pd.DataFrame):
                serie = serie.iloc[:, 0]
            df_out[col_alvo] = serie.fillna("").astype(str).str.strip()
        else:
            df_out[col_alvo] = ""

    col_ent = colunas_originais.get("ENTRADA")
    col_sai = colunas_originais.get("SAIDA")
    if col_ent and col_sai:
        ent_num = pd.to_numeric(df[col_ent].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
        sai_num = pd.to_numeric(df[col_sai].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
        sentido = []
        for e, s in zip(ent_num, sai_num):
            if e > 0 and s == 0:
                sentido.append("ENTRADA")
            elif s > 0 and e == 0:
                sentido.append("SAIDA")
            else:
                sentido.append("MOVIMENTO")
        df_out["Sentido Mov."] = sentido
    else:
        df_out["Sentido Mov."] = ""

    df_out["Cod. Material"] = df_out["Cod. Material"].str.replace(r"\.0$", "", regex=True).str.upper()

    nome_arq_norm = normalizar_texto(origem_arquivo)
    nome_aba_norm = normalizar_texto(nome_aba)
    if "LABORATORIO" in nome_arq_norm:
        mask_sem_proj = df_out["Contr / Proj Origem"] == ""
        df_out.loc[mask_sem_proj, "Contr / Proj Origem"] = f"LABORATORIO ({nome_aba.strip()})"
        mask_sem_pos = df_out["Posição"] == ""
        df_out.loc[mask_sem_pos, "Posição"] = "LABORATORIO"
        if "REPROVADOS" in nome_aba_norm:
            df_out.loc[df_out["Situação Mat."] == "", "Situação Mat."] = "REPROVADO"
        elif "DEFEITO EM CAMPO" in nome_aba_norm:
            df_out.loc[df_out["Situação Mat."] == "", "Situação Mat."] = "DEFEITO CAMPO"

    df_out["Dt. Movimento"] = padronizar_coluna_data(df_out["Dt. Movimento"])
    df_out["Dia Dig."] = padronizar_coluna_data(df_out["Dia Dig."])
    mask_dt_vazia = df_out["Dt. Movimento"] == ""
    df_out.loc[mask_dt_vazia, "Dt. Movimento"] = df_out.loc[mask_dt_vazia, "Dia Dig."]

    gpon_col = (
        df_out["GPON / Serial"]
        .str.upper()
        .str.extract(f"({PATTERN_GPON_STR})", expand=False)
        .fillna("")
        .str.replace(":", "", regex=False)
        .str.replace("-", "", regex=False)
    )

    mask_vazio = gpon_col == ""
    if mask_vazio.any():
        concat_vazio = df.loc[mask_vazio].fillna("").astype(str).agg(" ".join, axis=1).str.upper()
        gpon_col.loc[mask_vazio] = (
            concat_vazio.str.extract(f"({PATTERN_GPON_STR})", expand=False)
            .fillna("")
            .str.replace(":", "", regex=False)
            .str.replace("-", "", regex=False)
        )

    df_out["GPON / Serial"] = [g if serial_valido(g) else "" for g in gpon_col]
    df_out = df_out[df_out["GPON / Serial"] != ""].reset_index(drop=True)
    return df_out


def ler_xlsx_streaming_xml(caminho: str) -> dict[str, pd.DataFrame]:
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    abas_dfs = {}

    with zipfile.ZipFile(caminho, "r") as z:
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            with z.open("xl/sharedStrings.xml") as f:
                for _, elem in ET.iterparse(f, events=("end",)):
                    if elem.tag == f"{ns_main}si":
                        shared_strings.append("".join(elem.itertext()))
                        elem.clear()

        date_style_ids = set()
        if "xl/styles.xml" in z.namelist():
            st_root = ET.fromstring(z.read("xl/styles.xml"))
            num_fmts = {
                int(nf.attrib["numFmtId"]): nf.attrib.get("formatCode", "")
                for nf in st_root.findall(f".//{ns_main}numFmt")
            }
            cell_xfs = st_root.find(f"{ns_main}cellXfs")
            if cell_xfs is not None:
                for idx_xf, xf in enumerate(cell_xfs.findall(f"{ns_main}xf")):
                    nf_id = int(xf.attrib.get("numFmtId", 0))
                    fmt_code = num_fmts.get(nf_id, "").lower()
                    if nf_id in {14, 15, 16, 17, 22} or any(k in fmt_code for k in ("yy", "dd", "mm/")):
                        date_style_ids.add(str(idx_xf))

        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root}

        for sheet in wb_root.findall(f".//{ns_main}sheet"):
            name = sheet.attrib["name"]
            if normalizar_texto(name) in ABAS_IGNORAR:
                continue

            rid = sheet.attrib[f"{ns_rel}id"]
            target = rid_to_target.get(rid, "")
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            if target not in z.namelist():
                continue

            linhas = []
            with z.open(target) as f:
                for _, elem in ET.iterparse(f, events=("end",)):
                    if elem.tag == f"{ns_main}row":
                        row_dict = {}
                        for c in elem.findall(f"{ns_main}c"):
                            c_idx = col_letter_to_idx(c.attrib.get("r", "A1"))
                            if c_idx > 25:
                                continue
                            c_type = c.attrib.get("t", "")
                            c_style = c.attrib.get("s", "")
                            v_elem = c.find(f"{ns_main}v")
                            is_elem = c.find(f"{ns_main}is")
                            val = ""
                            if c_type == "s" and v_elem is not None and v_elem.text is not None:
                                val = shared_strings[int(v_elem.text)]
                            elif c_type == "inlineStr" and is_elem is not None:
                                val = "".join(is_elem.itertext())
                            elif v_elem is not None and v_elem.text is not None:
                                val = v_elem.text
                                if c_style in date_style_ids and c_type in ("", "n"):
                                    val = excel_num_to_date_str(val)
                            row_dict[c_idx] = str(val).strip()

                        if row_dict:
                            max_c = max(row_dict.keys())
                            vals = [row_dict.get(i, "") for i in range(max_c + 1)]
                            if any(vals):
                                linhas.append(vals)
                        elem.clear()

            if len(linhas) > 1:
                max_cols = max(len(r) for r in linhas)
                linhas_pad = [r + [""] * (max_cols - len(r)) for r in linhas]
                abas_dfs[name] = pd.DataFrame(linhas_pad[1:], columns=linhas_pad[0])
                print(f"    -> Aba '{name.strip()}': {len(abas_dfs[name])} linhas lidas", flush=True)

    return abas_dfs


def ler_ods_streaming(caminho: str) -> dict[str, pd.DataFrame]:
    abas = {}
    ns_table = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
    tag_table = f"{ns_table}table"
    tag_row = f"{ns_table}table-row"
    tag_cell = f"{ns_table}table-cell"
    attr_name = f"{ns_table}name"
    attr_rep_row = f"{ns_table}number-rows-repeated"
    attr_rep_col = f"{ns_table}number-columns-repeated"

    with zipfile.ZipFile(caminho, "r") as z:
        with z.open("content.xml") as xml_file:
            aba_atual = None
            linhas_aba = []
            for event, elem in ET.iterparse(xml_file, events=("start", "end")):
                if event == "start" and elem.tag == tag_table:
                    aba_atual = elem.attrib.get(attr_name, "Sheet")
                    linhas_aba = []
                elif event == "end" and elem.tag == tag_row:
                    if aba_atual and normalizar_texto(aba_atual) not in ABAS_IGNORAR:
                        rep_row = int(elem.attrib.get(attr_rep_row, 1))
                        vals = []
                        for cell in elem.findall(tag_cell):
                            rep_col = int(cell.attrib.get(attr_rep_col, 1))
                            txt = "".join(cell.itertext()).strip()
                            vals.extend([txt] * min(rep_col, 25))
                        if any(vals):
                            for _ in range(min(rep_row, 50)):
                                linhas_aba.append(vals[:20])
                    elem.clear()
                elif event == "end" and elem.tag == tag_table:
                    if len(linhas_aba) > 1 and aba_atual:
                        max_cols = max(len(r) for r in linhas_aba)
                        linhas_pad = [r + [""] * (max_cols - len(r)) for r in linhas_aba]
                        abas[aba_atual] = pd.DataFrame(linhas_pad[1:], columns=linhas_pad[0])
                    elem.clear()
    return abas


def ler_planilhas(caminho: str) -> list[pd.DataFrame]:
    nome_arquivo = os.path.basename(caminho)
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".ods":
        abas = ler_ods_streaming(caminho)
    elif ext in [".xlsx", ".xlsm"]:
        abas = ler_xlsx_streaming_xml(caminho)
    else:
        abas = pd.read_excel(caminho, sheet_name=None, dtype=str)

    resultados = []
    for nome_aba, df in abas.items():
        if normalizar_texto(nome_aba) in ABAS_IGNORAR or df is None or df.empty:
            continue
        df_tratado = mapear_colunas_brutas(df, origem_arquivo=nome_arquivo, nome_aba=nome_aba)
        if not df_tratado.empty:
            resultados.append(df_tratado)
    return resultados


def ler_pdf(caminho: str) -> list[pd.DataFrame]:
    nome_arquivo = os.path.basename(caminho)
    dfs = []
    with pdfplumber.open(caminho) as pdf:
        for idx_p, page in enumerate(pdf.pages, start=1):
            tabelas = page.extract_tables()
            for tab in tabelas:
                if len(tab) > 1:
                    df = pd.DataFrame(tab[1:], columns=tab[0])
                    dfs.append(mapear_colunas_brutas(df, origem_arquivo=nome_arquivo, nome_aba=f"Pag_{idx_p}"))
    return dfs


def ler_imagem_ocr(caminho: str) -> list[pd.DataFrame]:
    texto = pytesseract.image_to_string(Image.open(caminho), lang="por")
    registros = []
    for linha in texto.splitlines():
        if eh_passivo(linha):
            continue
        gpon = extrair_gpon(linha)
        if gpon:
            datas = re.findall(r"\b\d{2}/\d{2}/\d{2,4}\b", linha)
            horas = re.findall(r"\b\d{2}:\d{2}(?::\d{2})?\b", linha)
            dt_mov = padronizar_coluna_data(pd.Series([datas[0] if datas else ""])).iloc[0]
            dia_dig = padronizar_coluna_data(pd.Series([datas[1] if len(datas) >= 2 else (datas[0] if datas else "")])).iloc[0]
            registros.append({
                "GPON / Serial": gpon,
                "Cod. Material": "",
                "Dt. Movimento": dt_mov,
                "Dia Dig.": dia_dig,
                "Hora Dig.": horas[0] if horas else "",
                "Contr / Proj Origem": "",
                "Sentido Mov.": "",
                "Nº. Documento": "",
                "Situação Mat.": "BOM" if "BOM" in normalizar_texto(linha) else ("DEF. ESTOQUE" if "DEF" in normalizar_texto(linha) else ""),
                "Posição": next((p for p in ["ASSINANTE", "ESTOQUE", "VOLANTE", "CONVERTIDO", "FORNECEDOR", "PERDIDO"] if p in normalizar_texto(linha)), ""),
                "Responsável": "",
                "Observação": linha.strip(),
                "Usuário Dig.": "OCR_ANIEL",
            })
    return [pd.DataFrame(registros)] if registros else []


def classificar_estados_markov_ajustados(df: pd.DataFrame) -> pd.DataFrame:
    """
    Classifica cada movimentação no espaço de estados S0-S9:
      - S7 (RMA / Fornecedor) restrito exclusivamente à família de modelo 168 (168, R168, E168)
      - Prefixo 'E' em Cod. Material (DEF-ONU...) reconhecido como S6 (Defeito / Reprovado)
      - Prefixo 'R' em Cod. Material (RET-ONU...) reconhecido como S2 (Retorno de Campo)
    """
    estados = []
    ja_circulou_ou_triado = {}

    gpons = df["GPON / Serial"].tolist()
    cods = [normalizar_texto(x) for x in df["Cod. Material"].tolist()]
    projs = [normalizar_texto(x) for x in df["Contr / Proj Origem"].tolist()]
    sentidos = [normalizar_texto(x) for x in df["Sentido Mov."].tolist()]
    sits = [normalizar_texto(x) for x in df["Situação Mat."].tolist()]
    poss = [normalizar_texto(x) for x in df["Posição"].tolist()]
    obss = [normalizar_texto(x) for x in df["Observação"].tolist()]
    obs_orig = df["Observação"].tolist()

    for idx_row, (gpon, cod, proj, sent, sit, pos, obs) in enumerate(
        zip(gpons, cods, projs, sentidos, sits, poss, obss)
    ):
        if gpon not in ja_circulou_ou_triado:
            ja_circulou_ou_triado[gpon] = False

        eh_cod_retirado = cod.startswith("R") and len(cod) >= 2 and cod[1:].isdigit()
        eh_cod_defeito = cod.startswith("E") and len(cod) >= 2 and cod[1:].isdigit()
        eh_168 = eh_modelo_168(cod)
        menciona_rma = any(k in obs for k in ("RMA", "REMESSA PARA MANUTENCAO")) or ("FORNECEDOR" in pos)

        # S9 - Sucata / Baixa Definitiva (Estado Absorvente)
        if (
            "CONVERSAO P/ SUCATA" in obs
            or "CONVERSAO P/ DESCARTE" in obs
            or "PERDIDO" in pos
            or "SUCATA" in pos
        ):
            estado = "S9"

        # S7 - RMA / Fornecedor (Exclusivo Modelo 168)
        elif eh_168 and menciona_rma:
            estado = "S7"
            ja_circulou_ou_triado[gpon] = True

        # S6 - Defeito / Reprovado (inclui prefixo E- e lançamentos de remessa em modelos != 168)
        elif (
            eh_cod_defeito
            or menciona_rma
            or "REPROVADO" in sit
            or "DEF" in sit
            or "C/DEFEITO" in obs
            or any(k in obs for k in ("QUEIMADO", "QUEBRADO", "DANIFICADO", "APENAS LUZ POWER"))
        ):
            estado = "S6"
            ja_circulou_ou_triado[gpon] = True
            if menciona_rma and not eh_168:
                obs_orig[idx_row] = f"{str(obs_orig[idx_row]).strip()} [RMA IGNORADO - APENAS MODELO 168]".strip()

        # S4 - Conversão de Material (RET <-> BOM)
        elif "CONVERSAO DE MATERIAL" in obs or "CONVERTIDO" in pos:
            estado = "S4"
            ja_circulou_ou_triado[gpon] = True

        # S8 - Quarentena Obsoletos / Descarte
        elif "OBSOLETO" in proj or "DESCARTE" in proj or "OBSOLETO" in pos:
            estado = "S8"
            ja_circulou_ou_triado[gpon] = True

        # S3 - Laboratório / Triagem Técnica
        elif "LABORATORIO" in proj or "LABORATORIO" in pos or "VISTORIA" in proj or "VISTORIA" in pos:
            estado = "S3"
            ja_circulou_ou_triado[gpon] = True

        # S2 - Retorno de Campo (Material RET-)
        elif (
            "DEVOLUCAO AO ESTOQUE-RET" in obs
            or "RETIRADO" in sit
            or (eh_cod_retirado and sent == "ENTRADA")
        ):
            estado = "S2"
            ja_circulou_ou_triado[gpon] = True

        # S1 - Expedição Campo / Equipe Técnica
        elif (
            "VOLANTE" in pos
            or "ASSINANTE" in pos
            or "INV.EQUIPE" in obs
            or "LOJA SEA" in obs
            or "PROJ " in obs
            or (sent == "SAIDA" and any(k in proj for k in ("GRADEADO", "ANALISE DE ESTOQUE", "AEREO")))
        ):
            estado = "S1"
            ja_circulou_ou_triado[gpon] = True

        # S5 vs S0 - Estoque Reaproveitado (Reuso) vs Estoque Novo / Bom
        elif (
            "REAPROVEITADO" in sit
            or "DEVOLUCAO AO ESTOQUE-BOM" in obs
            or "ESTOQUE" in pos
            or any(k in proj for k in ("GRADEADO", "ANALISE DE ESTOQUE", "ARMARIO NOC"))
        ):
            if "REAPROVEITADO" in sit or "DEVOLUCAO AO ESTOQUE-BOM" in obs or ja_circulou_ou_triado[gpon]:
                estado = "S5"
                ja_circulou_ou_triado[gpon] = True
            else:
                estado = "S0"

        else:
            estado = "S5" if ja_circulou_ou_triado[gpon] else "S0"

        estados.append(estado)

    df = df.copy()
    df["Estado_Markov"] = estados

    # Reconciliação de S9 (Sucata Absorvente): se houver movimentação posterior válida, reclassifica para S8
    for _, grupo in df.groupby("GPON / Serial", sort=False):
        idxs = grupo.index.tolist()
        ests = grupo["Estado_Markov"].tolist()
        for pos_rel, (row_i, st) in enumerate(zip(idxs, ests)):
            if st == "S9" and pos_rel < len(idxs) - 1:
                if any(s != "S9" for s in ests[pos_rel + 1:]):
                    estados[row_i] = "S8"
                    obs_orig[row_i] = f"{str(obs_orig[row_i]).strip()} [S9 RECLASSIFICADO - MOVIMENTO POSTERIOR]".strip()

    df["Estado_Markov"] = estados
    df["Observação"] = obs_orig
    return df


def calcular_tempos_e_transicoes(df: pd.DataFrame, ano_filtro: int) -> pd.DataFrame:
    df = df.copy()

    if IGNORAR_INVENTARIO_CONTABIL:
        obs_norm = df["Observação"].apply(normalizar_texto)
        df = df[~obs_norm.str.contains("INVENTARIO DO ESTOQUE FISICO", na=False)].copy()

    df["_dt_evento"] = [
        parse_timestamp_padrao(d1, d2, h)
        for d1, d2, h in zip(df["Dt. Movimento"], df["Dia Dig."], df["Hora Dig."])
    ]

    fim_ano = pd.Timestamp(year=ano_filtro, month=12, day=31, hour=23, minute=59, second=59)
    df = df[df["_dt_evento"].notna() & (df["_dt_evento"] <= fim_ano)].copy()

    df = df.sort_values(
        by=["GPON / Serial", "_dt_evento", "Dia Dig.", "Hora Dig."],
        ascending=[True, True, True, True],
    ).reset_index(drop=True)

    df = classificar_estados_markov_ajustados(df)

    # Remove duplicidades exatas
    df = df.drop_duplicates(subset=["GPON / Serial", "_dt_evento", "Estado_Markov"], keep="first").reset_index(drop=True)

    # Resolve eventos simultâneos (0,00 dias) mantendo o estado de maior precedência
    df["_prec"] = df["Estado_Markov"].map(PRECEDENCIA_ESTADO).fillna(0)
    df = (
        df.sort_values(by=["GPON / Serial", "_dt_evento", "_prec"], ascending=[True, True, False])
        .drop_duplicates(subset=["GPON / Serial", "_dt_evento"], keep="first")
        .drop(columns=["_prec"])
        .sort_values(by=["GPON / Serial", "_dt_evento"], ascending=[True, True])
        .reset_index(drop=True)
    )

    # Comprime auto-transições consecutivas idênticas (atualizações internas sem mudança física de estado)
    estado_anterior = df.groupby("GPON / Serial")["Estado_Markov"].shift(1)
    df = df[df["Estado_Markov"] != estado_anterior].reset_index(drop=True)

    df["Proximo_Estado"] = df.groupby("GPON / Serial")["Estado_Markov"].shift(-1)
    df["_dt_proximo"] = df.groupby("GPON / Serial")["_dt_evento"].shift(-1)

    df_ano = df[df["_dt_evento"].dt.year == ano_filtro].copy()
    if df_ano.empty:
        return pd.DataFrame(columns=COLUNAS_PADRAO)

    prox_out, dias_out, horas_out, status_out = [], [], [], []
    for dt_ini, dt_fim, est_atual, prox_est in zip(
        df_ano["_dt_evento"], df_ano["_dt_proximo"], df_ano["Estado_Markov"], df_ano["Proximo_Estado"]
    ):
        if est_atual in ESTADOS_ABSORVENTES:
            prox_out.append(est_atual)
            dias_out.append(0.0)
            horas_out.append(0.0)
            status_out.append("Estado Absorvente")
        elif pd.notna(dt_fim) and dt_fim <= fim_ano:
            delta_seg = max((dt_fim - dt_ini).total_seconds(), 0.0)
            prox_out.append(prox_est)
            dias_out.append(round(delta_seg / 86400.0, 2))
            horas_out.append(round(delta_seg / 3600.0, 2))
            status_out.append("Concluída")
        else:
            delta_seg = max((fim_ano - dt_ini).total_seconds(), 0.0)
            prox_out.append("EM_ABERTO_FIM_2025")
            dias_out.append(round(delta_seg / 86400.0, 2))
            horas_out.append(round(delta_seg / 3600.0, 2))
            status_out.append("Em Aberto (Fim 2025)")

    df_ano["Proximo_Estado"] = prox_out
    df_ano["Dias_no_Estado"] = dias_out
    df_ano["Horas_no_Estado"] = horas_out
    df_ano["Status_Transicao"] = status_out

    return df_ano.drop(columns=["_dt_evento", "_dt_proximo"]).reset_index(drop=True)


def gerar_tabelas_analiticas_markov(df_2025: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    estados_presentes = sorted(df_2025["Estado_Markov"].unique().tolist())

    # 1. Matriz de Transição Discreta (apenas transições Concluídas + Absorventes)
    df_conc = df_2025[df_2025["Status_Transicao"].isin(["Concluída", "Estado Absorvente"])].copy()
    contagem_conc = pd.crosstab(df_conc["Estado_Markov"], df_conc["Proximo_Estado"])
    contagem_conc = contagem_conc.reindex(index=estados_presentes, columns=estados_presentes, fill_value=0)

    for st_abs in ESTADOS_ABSORVENTES:
        if st_abs in estados_presentes:
            contagem_conc.loc[st_abs, :] = 0
            contagem_conc.loc[st_abs, st_abs] = max((df_2025["Estado_Markov"] == st_abs).sum(), 1)

    somas_linha = contagem_conc.sum(axis=1)
    prob_discreta = contagem_conc.div(somas_linha.replace(0, 1), axis=0) * 100.0
    prob_discreta = prob_discreta.round(2)

    df_matriz_discreta = prob_discreta.copy()
    df_matriz_discreta.index = [DESCRICAO_ESTADOS.get(s, s) for s in df_matriz_discreta.index]
    df_matriz_discreta.columns = [f"Para {s} (%)" for s in df_matriz_discreta.columns]
    df_matriz_discreta.insert(0, "Estado Atual (Origem)", df_matriz_discreta.index)
    df_matriz_discreta["Transições Concluídas"] = somas_linha.values
    df_matriz_discreta = df_matriz_discreta.reset_index(drop=True)

    # 2. Matriz Geradora Q (Cadeia de Markov em Tempo Contínuo — CTMC)
    df_saidas = df_2025[df_2025["Status_Transicao"] == "Concluída"].copy()
    contagem_saidas = pd.crosstab(df_saidas["Estado_Markov"], df_saidas["Proximo_Estado"])
    contagem_saidas = contagem_saidas.reindex(index=estados_presentes, columns=estados_presentes, fill_value=0)

    exposicao_dias = df_2025.groupby("Estado_Markov")["Dias_no_Estado"].sum().reindex(estados_presentes, fill_value=0.0)
    matriz_q = pd.DataFrame(0.0, index=estados_presentes, columns=estados_presentes)

    for st_i in estados_presentes:
        if st_i in ESTADOS_ABSORVENTES or exposicao_dias[st_i] <= 0:
            continue
        for st_j in estados_presentes:
            if st_i != st_j:
                matriz_q.loc[st_i, st_j] = round(contagem_saidas.loc[st_i, st_j] / exposicao_dias[st_i], 5)
        matriz_q.loc[st_i, st_i] = round(-matriz_q.loc[st_i, [c for c in estados_presentes if c != st_i]].sum(), 5)

    df_ctmc = matriz_q.copy()
    df_ctmc.index = [DESCRICAO_ESTADOS.get(s, s) for s in df_ctmc.index]
    df_ctmc.columns = [f"q_{s} (trans/dia)" for s in df_ctmc.columns]
    df_ctmc.insert(0, "Estado Atual (Origem)", df_ctmc.index)
    df_ctmc["Exposição Total (Dias)"] = exposicao_dias.round(2).values
    taxas_saida = [-matriz_q.loc[s, s] for s in estados_presentes]
    df_ctmc["Taxa Saída Total (1/dia)"] = [round(t, 5) for t in taxas_saida]
    df_ctmc["Tempo Médio Espera CTMC (Dias)"] = [round(1.0 / t, 2) if t > 0 else 0.0 for t in taxas_saida]
    df_ctmc = df_ctmc.reset_index(drop=True)

    # 3. Tabela Tempo_Permanencia_2025 reconciliada
    linhas_tempo = []
    for st in estados_presentes:
        sub = df_2025[df_2025["Estado_Markov"] == st]
        saidas = int((sub["Status_Transicao"] == "Concluída").sum())
        retidos = int((sub["Status_Transicao"] == "Em Aberto (Fim 2025)").sum())
        absorvidos = int((sub["Status_Transicao"] == "Estado Absorvente").sum())
        qtd_total = saidas + retidos + absorvidos

        dias_serie = pd.to_numeric(sub["Dias_no_Estado"], errors="coerce").dropna()
        horas_serie = pd.to_numeric(sub["Horas_no_Estado"], errors="coerce").dropna()

        linhas_tempo.append({
            "Estado": st,
            "Descrição": DESCRICAO_ESTADOS.get(st, st),
            "Qtd_Ocorrencias": qtd_total,
            "Transicoes_Saida": saidas,
            "Retidos_Fim_2025": retidos,
            "Absorvidos": absorvidos,
            "Tempo_Medio_Dias": round(float(dias_serie.mean()), 2) if not dias_serie.empty else 0.0,
            "Mediana_Dias": round(float(dias_serie.median()), 2) if not dias_serie.empty else 0.0,
            "Min_Dias": round(float(dias_serie.min()), 2) if not dias_serie.empty else 0.0,
            "Max_Dias": round(float(dias_serie.max()), 2) if not dias_serie.empty else 0.0,
            "Tempo_Medio_Horas": round(float(horas_serie.mean()), 2) if not horas_serie.empty else 0.0,
        })

    return df_matriz_discreta, df_ctmc, pd.DataFrame(linhas_tempo)


def estilizar_aba(ws, df: pd.DataFrame):
    ws.views.sheetView[0].showGridLines = True
    fill_header = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_zebra = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    font_body = Font(name="Calibri", size=10, color="111827")
    borda_fina = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    ws.append(list(df.columns))
    for col_idx in range(1, len(df.columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina
    ws.row_dimensions[1].height = 24

    colunas_centro = {
        "GPON / Serial", "Cod. Material", "Dt. Movimento", "Dia Dig.", "Hora Dig.",
        "Sentido Mov.", "Estado_Markov", "Proximo_Estado", "Dias_no_Estado", "Horas_no_Estado", "Estado"
    }

    for row_idx, row_data in enumerate(df.itertuples(index=False), start=2):
        ws.append(list(row_data))
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = font_body
            cell.border = borda_fina
            if row_idx % 2 == 0:
                cell.fill = fill_zebra
            if df.columns[col_idx - 1] in colunas_centro:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, col_name in enumerate(df.columns, start=1):
        max_len = max([len(str(col_name))] + [len(str(val)) for val in df[col_name].fillna("")])
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 48)

    ws.freeze_panes = "A2"


def exportar_excel_mestre(
    df_mestre: pd.DataFrame,
    df_matriz_discreta: pd.DataFrame,
    df_ctmc: pd.DataFrame,
    df_tempo: pd.DataFrame,
    caminho_saida: str,
):
    wb = Workbook()

    ws_geral = wb.active
    ws_geral.title = f"Geral_{ANO_FILTRO}"
    estilizar_aba(ws_geral, df_mestre)

    ws_disc = wb.create_sheet(title=f"Matriz_Discreta_{ANO_FILTRO}")
    estilizar_aba(ws_disc, df_matriz_discreta)

    ws_ctmc = wb.create_sheet(title=f"Matriz_Taxas_CTMC_{ANO_FILTRO}")
    estilizar_aba(ws_ctmc, df_ctmc)

    ws_tempo = wb.create_sheet(title=f"Tempo_Permanencia_{ANO_FILTRO}")
    estilizar_aba(ws_tempo, df_tempo)

    wb.save(caminho_saida)


def executar_pipeline():
    t_inicio = time.time()
    print(f"Iniciando leitura e saneamento na pasta 'entrada_materiais' (Ano: {ANO_FILTRO})...", flush=True)
    dfs = []
    arquivos = sorted(glob.glob(os.path.join(PASTA_ENTRADA, "*")))

    if not arquivos:
        print("AVISO: A pasta 'entrada_materiais' está vazia! Coloque seus arquivos lá dentro e tente novamente.")
        return

    for arquivo in arquivos:
        nome = os.path.basename(arquivo)
        if nome.startswith("~$"):
            continue
        ext = os.path.splitext(arquivo)[1].lower()
        t_arq = time.time()
        print(f" -> Processando: {nome}", flush=True)
        if ext in [".xlsx", ".xlsm", ".xls", ".ods"]:
            dfs.extend(ler_planilhas(arquivo))
        elif ext == ".pdf":
            dfs.extend(ler_pdf(arquivo))
        elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"]:
            dfs.extend(ler_imagem_ocr(arquivo))
        print(f"    Concluído em {time.time() - t_arq:.1f}s", flush=True)

    if not dfs:
        print("Nenhum registro válido encontrado nos arquivos.")
        return

    print("Consolidando registros e filtrando itens passivos/seriais inválidos...", flush=True)
    df_total = pd.concat(dfs, ignore_index=True)
    df_total = df_total[df_total["GPON / Serial"].apply(serial_valido)].copy()

    texto_concat = df_total.fillna("").astype(str).agg(" ".join, axis=1)
    df_total = df_total[~texto_concat.apply(eh_passivo)].copy()

    colunas_base = [
        c for c in COLUNAS_PADRAO
        if c not in {"Estado_Markov", "Proximo_Estado", "Dias_no_Estado", "Horas_no_Estado", "Status_Transicao"}
    ]
    for col in colunas_base:
        if col not in df_total.columns:
            df_total[col] = ""

    df_total = df_total.drop_duplicates(subset=colunas_base).reset_index(drop=True)

    print("Calculando cadeia de Markov (S0-S9, RMA exclusivo modelo 168)...", flush=True)
    df_2025 = calcular_tempos_e_transicoes(df_total, ANO_FILTRO)
    if df_2025.empty:
        print(f"Nenhum registro válido encontrado para o ano {ANO_FILTRO}.")
        return

    df_2025 = df_2025[COLUNAS_PADRAO]
    df_matriz_discreta, df_ctmc, df_tempo = gerar_tabelas_analiticas_markov(df_2025)

    print("Salvando arquivo Excel mestre consolidado...", flush=True)
    exportar_excel_mestre(df_2025, df_matriz_discreta, df_ctmc, df_tempo, ARQUIVO_SAIDA)
    print(
        f"\nSUCESSO! {len(df_2025)} eventos saneados de {ANO_FILTRO} salvos em "
        f"'{os.path.basename(ARQUIVO_SAIDA)}' ({time.time() - t_inicio:.1f}s no total)."
    )


if __name__ == "__main__":
    executar_pipeline()