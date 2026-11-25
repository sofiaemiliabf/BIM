import io
from pathlib import Path
import re
from datetime import date, datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import plotly.express as px
import streamlit as st

# ==============================================================================
# 1. CONFIGURACIÓN DE PÁGINA Y PALETA PROFESIONAL (AZUL, MORADO, ROSADO, CELESTE)
# ==============================================================================
st.set_page_config(
    page_title="PlanificaObra | Gestión de Recursos & APU",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_EXCEL_FILENAME = "7.21.5_Anexo_N°_4_Presupuesto (2).xlsx"
DEFAULT_EXCEL_PATH = Path(__file__).resolve().parent / DEFAULT_EXCEL_FILENAME

PALETTE = {
    "navy": "#1E3A8A",      # Azul profundo (títulos y estructura)
    "purple": "#6D28D9",    # Morado corporativo (primario)
    "pink": "#EC4899",      # Rosado acento (destacados y alertas)
    "sky": "#38BDF8",       # Celeste moderno (gráficos y badges)
    "indigo": "#4F46E5",    # Azul-morado intermedio
    "bg": "#F8FAFC",        # Fondo limpio
    "card": "#FFFFFF",      # Tarjetas blancas
    "text": "#1E293B",      # Texto principal
    "muted": "#64748B",     # Texto secundario
}

CUSTOM_CSS = f"""
<style>
    /* Fondo general y tipografía */
    .stApp {{
        background-color: {PALETTE["bg"]};
        color: {PALETTE["text"]};
        font-family: 'Inter', 'Segoe UI', sans-serif;
    }}

    /* Encabezado Hero con degradado Azul -> Morado -> Rosado */
    .hero-banner {{
        background: linear-gradient(125deg, {PALETTE["navy"]} 0%, {PALETTE["purple"]} 58%, {PALETTE["pink"]} 100%);
        padding: 1.6rem 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px rgba(109, 40, 217, 0.18);
    }}
    .hero-banner h1 {{
        margin: 0;
        font-size: 1.9rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        color: #FFFFFF;
    }}
    .hero-banner p {{
        margin: 0.45rem 0 0 0;
        font-size: 0.96rem;
        color: #E0F2FE;
    }}

    /* Tarjetas KPI */
    .kpi-card {{
        background: {PALETTE["card"]};
        border-radius: 14px;
        padding: 1.15rem 1.3rem;
        border: 1px solid #E2E8F0;
        border-top: 4px solid {PALETTE["purple"]};
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
        transition: transform 0.15s ease;
    }}
    .kpi-card:hover {{
        transform: translateY(-2px);
    }}
    .kpi-card.blue {{ border-top-color: {PALETTE["navy"]}; }}
    .kpi-card.purple {{ border-top-color: {PALETTE["purple"]}; }}
    .kpi-card.pink {{ border-top-color: {PALETTE["pink"]}; }}
    .kpi-card.sky {{ border-top-color: {PALETTE["sky"]}; }}

    .kpi-label {{
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: {PALETTE["muted"]};
        font-weight: 700;
        margin-bottom: 0.35rem;
    }}
    .kpi-value {{
        font-size: 1.55rem;
        font-weight: 800;
        color: {PALETTE["navy"]};
        line-height: 1.2;
    }}
    .kpi-sub {{
        font-size: 0.78rem;
        color: {PALETTE["muted"]};
        margin-top: 0.25rem;
    }}

    /* Sidebar estilizada */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #1E1B4B 0%, #312E81 65%, #4C1D95 100%);
        color: white;
    }}
    section[data-testid="stSidebar"] * {{
        color: #F8FAFC !important;
    }}
    section[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] * {{
        color: #1E293B !important;
    }}
    section[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] * {{
        color: #1E293B !important;
    }}

    /* Pestañas */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background-color: #EEF2FF;
        padding: 6px;
        border-radius: 12px;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 8px;
        padding: 8px 18px;
        font-weight: 600;
        color: {PALETTE["navy"]};
    }}
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(90deg, {PALETTE["purple"]}, {PALETTE["indigo"]}) !important;
        color: white !important;
    }}

    /* Badges de categoría */
    .section-title {{
        font-size: 1.15rem;
        font-weight: 700;
        color: {PALETTE["navy"]};
        border-left: 4px solid {PALETTE["pink"]};
        padding-left: 10px;
        margin: 1rem 0 0.75rem 0;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==============================================================================
# 2. FUNCIONES AUXILIARES DE LIMPIEZA Y NORMALIZACIÓN
# ==============================================================================
SPECIAL_SHEETS = {
    "PPTO MODIFICADO",
    "Gastos Generales",
    "MATERIALES",
    "MANO DE OBRA",
    "MAQUINARIA",
}

SUMMARY_LABELS = {"DIRECTO", "GG", "UTIL", "NETO", "IVA", "TOTAL"}


def safe_float(val, default=0.0):
    if val is None or isinstance(val, bool):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    text = str(val).strip().replace("$", "").replace(" ", "")
    if not text or text in {"-", "—", "N/A"}:
        return default
    # Manejo básico de separadores chilenos/decimales
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return default


def clean_str(val):
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val)).strip()


def normalize_item_code(val, known_apu_sheets=None):
    """Normaliza códigos de ítem y corrige casos convertidos a datetime por Excel."""
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        # Excel suele convertir "10.1" -> 2026-01-10 o "2.6" -> 2026-06-02
        candidates = [
            f"{val.day}.{val.month}",
            f"{val.month}.{val.day}",
            f"{val.day}.{val.month}.{val.year % 100}",
        ]
        if known_apu_sheets:
            for cand in candidates:
                if cand in known_apu_sheets:
                    return cand
        return candidates[0]
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return clean_str(val).rstrip(".")


def normalize_resource_code(raw_code):
    """Limpia prefijos vacíos (MAT-, MO-, MAQ-) y unifica ceros a la izquierda (MO-01 -> MO-1)."""
    code = clean_str(raw_code).upper()
    if not code or code in {"MAT-", "MO-", "MAQ-", "SUB-", "-"}:
        return "SIN CÓDIGO"
    match = re.match(r"^(MAT|MO|MAQ)-0*(\d+)$", code)
    if match:
        return f"{match.group(1)}-{int(match.group(2))}"
    return code


def normalize_unit(unit_val):
    u = clean_str(unit_val)
    if not u:
        return "-"
    mapping = {
        "u": "UN",
        "un": "UN",
        "uni": "UN",
        "unidad": "UN",
        "unidades": "UN",
        "undiad": "UN",
        "día": "Día",
        "dia": "Día",
        "dias": "Día",
        "días": "Día",
        "m2": "m²",
        "m3": "m³",
        "ml": "ml",
        "m": "m",
        "kg": "kg",
        "gl": "GL",
        "pza": "Pza",
        "saco": "Saco",
        "sacos": "Saco",
        "rollo": "Rollo",
        "plancha": "Plancha",
        "planchas": "Plancha",
        "placa": "Placa",
        "tineta": "Tineta",
        "gal": "Galón",
        "galon": "Galón",
        "bolsa": "Bolsa",
        "caja": "Caja",
        "pack": "Pack",
        "tarro": "Tarro",
        "lt": "L",
        "hr": "Hr",
        "mes": "Mes",
        "barra": "Barra",
        "perfil": "Perfil",
        "perfiles": "Perfil",
    }
    return mapping.get(u.lower(), u)


def fmt_clp(val):
    return f"$ {val:,.0f}".replace(",", ".")


def fmt_num(val, dec=2):
    return f"{val:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def item_sort_key(code):
    """Ordenamiento jerárquico natural: 1.2.1 antes de 1.10.1 y 2.1 antes de 10.1."""
    parts = re.findall(r"\d+", str(code))
    return tuple(int(p) for p in parts) if parts else (999999,)


# ==============================================================================
# 3. MOTOR DE EXTRACCIÓN CON OPENPYXL (ADAPTADO AL EXCEL)
# ==============================================================================
def parse_ppto_sheet(ws, known_apu_sheets):
    """Lee la hoja PPTO MODIFICADO: metadatos, capítulos, partidas y cierre."""
    meta = {"proyecto": "Sin especificar", "ubicacion": "Sin especificar"}
    partidas = []
    resumen = {}

    current_cap_num = ""
    current_cap_name = ""

    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_col=9, values_only=True), start=1):
        col_b = row[1] if len(row) > 1 else None
        col_c = row[2] if len(row) > 2 else None
        col_d = row[3] if len(row) > 3 else None
        col_e = row[4] if len(row) > 4 else None
        col_f = row[5] if len(row) > 5 else None
        col_g = row[6] if len(row) > 6 else None

        # Metadatos del proyecto
        b_str = clean_str(col_b).upper()
        if "PROYECTO" in b_str and col_c:
            meta["proyecto"] = clean_str(col_c).strip("“”\"")
        elif "UBICACIÓN" in b_str or "UBICACION" in b_str:
            if col_c:
                meta["ubicacion"] = clean_str(col_c)

        if r_idx <= 10:
            continue

        # Filas de resumen final (DIRECTO, GG, UTIL, NETO, IVA, TOTAL)
        e_str = clean_str(col_e).upper()
        if e_str in SUMMARY_LABELS:
            resumen[e_str] = {
                "porcentaje": safe_float(col_f, None) if col_f is not None else None,
                "monto": safe_float(col_g, 0.0),
            }
            continue

        if col_b is None and col_c is None:
            continue

        item_code = normalize_item_code(col_b, known_apu_sheets)
        desc = clean_str(col_c)
        unidad = clean_str(col_d)
        cant = safe_float(col_e, None) if col_e is not None else None
        valor_unit = safe_float(col_f, None) if col_f is not None else None
        total = safe_float(col_g, None) if col_g is not None else None

        if not item_code or not desc:
            continue

        # Detectar capítulos principales (ej: "1", "2", ... "10")
        if re.fullmatch(r"\d+", item_code) and cant is None:
            current_cap_num = item_code
            current_cap_name = desc
            continue

        # Partida presupuestaria (tiene unidad y cantidad)
        if cant is not None and (valor_unit is not None or total is not None):
            cap_id = item_code.split(".")[0] if "." in item_code else current_cap_num
            partidas.append(
                {
                    "Item": item_code,
                    "Capitulo_ID": cap_id,
                    "Capitulo": f"{cap_id}. {current_cap_name}" if current_cap_name else f"Capítulo {cap_id}",
                    "Designacion": desc,
                    "Unidad": normalize_unit(unidad),
                    "Cantidad_PPTO": cant,
                    "Precio_Unitario_PPTO": valor_unit or 0.0,
                    "Total_PPTO": total if total is not None else cant * (valor_unit or 0.0),
                    "Tiene_APU": item_code in known_apu_sheets,
                }
            )

    df_ppto = pd.DataFrame(partidas)
    if not df_ppto.empty:
        df_ppto["_sort"] = df_ppto["Item"].apply(item_sort_key)
        df_ppto = df_ppto.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    return meta, df_ppto, resumen


def parse_catalog_sheets(wb):
    """Lee las hojas maestras MATERIALES, MANO DE OBRA y MAQUINARIA."""
    catalogs = {}

    if "MATERIALES" in wb.sheetnames:
        ws = wb["MATERIALES"]
        rows = []
        for row in ws.iter_rows(min_row=5, max_col=9, values_only=True):
            code = normalize_resource_code(row[1])
            item = clean_str(row[2])
            if not item:
                continue
            rows.append(
                {
                    "Código": code,
                    "Material": item,
                    "Unidad": normalize_unit(row[3]),
                    "Cantidad_Base": safe_float(row[4], 1.0),
                    "Valor_Con_IVA": safe_float(row[5], 0.0),
                    "Precio_Neto": safe_float(row[6], 0.0),
                    "IVA": safe_float(row[7], 0.0),
                    "Fuente_Cotizacion": clean_str(row[8]) if len(row) > 8 else "",
                }
            )
        catalogs["MATERIALES"] = pd.DataFrame(rows)

    if "MANO DE OBRA" in wb.sheetnames:
        ws = wb["MANO DE OBRA"]
        rows = []
        for row in ws.iter_rows(min_row=3, max_col=8, values_only=True):
            code = normalize_resource_code(row[1])
            item = clean_str(row[2])
            if not item:
                continue
            rows.append(
                {
                    "Código": code,
                    "Cargo / Especialidad": item,
                    "Cantidad_Base": safe_float(row[3], 1.0),
                    "Unidad": normalize_unit(row[4]),
                    "Valor_Diario": safe_float(row[5], 0.0),
                    "Observación": clean_str(row[7]) if len(row) > 7 else "",
                }
            )
        catalogs["MANO DE OBRA"] = pd.DataFrame(rows)

    if "MAQUINARIA" in wb.sheetnames:
        ws = wb["MAQUINARIA"]
        rows = []
        for row in ws.iter_rows(min_row=4, max_col=10, values_only=True):
            code = normalize_resource_code(row[1])
            item = clean_str(row[2])
            if not item:
                continue
            rows.append(
                {
                    "Código": code,
                    "Equipo / Maquinaria": item,
                    "Cantidad_Base": safe_float(row[3], 1.0),
                    "Unidad": normalize_unit(row[4]),
                    "Arriendo": safe_float(row[5], 0.0),
                    "Arriendo_c_Operador": safe_float(row[6], 0.0),
                    "Compra": safe_float(row[7], 0.0),
                    "Fuente": clean_str(row[9]) if len(row) > 9 else "",
                }
            )
        catalogs["MAQUINARIA"] = pd.DataFrame(rows)

    return catalogs


def parse_apu_sheet(ws, ppto_lookup):
    """Extrae cabecera, recursos (A, B, C, D) y leyes sociales de una hoja APU."""
    item_code = clean_str(ws.title)
    contratista = ""
    actividad = ""
    unidad_apu = ""
    cantidad_apu = 0.0
    costo_unit_apu = 0.0
    total_apu = 0.0

    current_section = None
    in_table = False
    section_c_done = False
    resources = []
    leyes_sociales_rows = []
    observaciones = []

    all_rows = list(ws.iter_rows(min_row=1, max_col=10, values_only=True))

    for r_idx, row in enumerate(all_rows, start=1):
        col_b = row[1] if len(row) > 1 else None
        col_c = row[2] if len(row) > 2 else None
        col_d = row[3] if len(row) > 3 else None
        col_e = row[4] if len(row) > 4 else None
        col_f = row[5] if len(row) > 5 else None
        col_g = row[6] if len(row) > 6 else None
        col_h = row[7] if len(row) > 7 else None
        col_j = row[9] if len(row) > 9 else None

        b_str = clean_str(col_b)
        b_upper = b_str.upper()
        c_str = clean_str(col_c)
        f_str = clean_str(col_f)
        g_str = clean_str(col_g)

        # Metadatos
        if "CONTRATISTA" in b_upper and c_str:
            contratista = c_str
        elif r_idx == 9:
            actividad = c_str
            unidad_apu = normalize_unit(col_e)
            cantidad_apu = safe_float(col_f, 0.0)
            costo_unit_apu = safe_float(col_g, 0.0)
            total_apu = safe_float(col_h, 0.0)
            if col_j and clean_str(col_j):
                observaciones.append(clean_str(col_j))
            continue

        # Observaciones técnicas en columna J
        if r_idx >= 10 and col_j and isinstance(col_j, str) and len(clean_str(col_j)) > 3:
            if "criterios/observaciones" not in clean_str(col_j).lower():
                observaciones.append(clean_str(col_j))

        # Detección dinámica de secciones A, B, C, D
        if b_upper.startswith("A-") or "A- MATERIALES" in b_upper:
            current_section = "Materiales"
            in_table = False
            continue
        if b_upper.startswith("B-") or "MAQUINAS Y EQUIPOS" in b_upper:
            current_section = "Maquinaria y Equipos"
            in_table = False
            continue
        if b_upper.startswith("C-") or "MANO DE OBRA" in b_upper:
            current_section = "Mano de Obra"
            in_table = False
            section_c_done = False
            continue
        if b_upper.startswith("D-") or "SUBCONTRATOS" in b_upper:
            current_section = "Subcontratos y Otros"
            in_table = False
            continue

        # Fila de encabezado de subtabla ("Código" | "Descripción Recurso")
        if b_upper in {"CÓDIGO", "CODIGO"} and "DESCRIPCI" in c_str.upper():
            # Caso borde (ej. hoja 3.11.3): tras Subtotal C aparece la tabla D sin título "D-Subcontratos"
            if current_section == "Mano de Obra" and section_c_done and not in_table:
                current_section = "Subcontratos y Otros"
            in_table = True
            continue

        # Detección de subtotales y fin del APU
        if "SUBTOTAL" in g_str.upper() or "SUBTOTAL" in b_upper:
            if current_section == "Mano de Obra":
                section_c_done = True
            in_table = False
            continue
        if "TOTAL COSTO UNITARIO DIRECTO" in b_upper:
            if col_h is not None and safe_float(col_h, 0.0) > 0:
                costo_unit_apu = safe_float(col_h, costo_unit_apu)
            break

        # Procesar filas de recursos dentro de una tabla activa
        if not in_table or current_section is None:
            continue

        # Separar fila especial de Leyes Sociales
        if "LEYES SOCIALES" in f_str.upper() or "LEYES SOCIALES" in c_str.upper():
            leyes_sociales_rows.append(
                {
                    "Item": item_code,
                    "Actividad": actividad,
                    "Porcentaje_LS": safe_float(col_g, 0.0),
                    "Costo_Unit_LS": safe_float(col_h, 0.0),
                }
            )
            continue

        rendimiento = safe_float(col_d, 0.0)
        factor_perdida = safe_float(col_f, 1.0)
        if factor_perdida <= 0:
            factor_perdida = 1.0
        costo_unit = safe_float(col_g, 0.0)
        total_unit = safe_float(col_h, 0.0)

        # Ignorar filas vacías o sin descripción ni costo
        if not c_str and rendimiento == 0 and total_unit == 0:
            continue

        # Recalcular total unitario si falta pero existen rendimiento y costo unitario
        if total_unit == 0 and rendimiento > 0 and costo_unit > 0:
            total_unit = rendimiento * factor_perdida * costo_unit

        ppto_info = ppto_lookup.get(item_code, {})
        cant_ppto = ppto_info.get("Cantidad_PPTO", cantidad_apu)
        capitulo = ppto_info.get("Capitulo", f"Capítulo {item_code.split('.')[0]}")
        actividad_final = ppto_info.get("Designacion", actividad) or actividad

        resources.append(
            {
                "Item": item_code,
                "Capitulo": capitulo,
                "Actividad": actividad_final,
                "Unidad_Partida": ppto_info.get("Unidad", unidad_apu),
                "Cantidad_PPTO": cant_ppto,
                "Cantidad_APU": cantidad_apu,
                "Categoria": current_section,
                "Codigo": normalize_resource_code(col_b),
                "Recurso": c_str,
                "Unidad_Recurso": normalize_unit(col_e),
                "Rendimiento": rendimiento,
                "Factor_Perdida": factor_perdida,
                "Costo_Unitario": costo_unit,
                "Total_Unitario": total_unit,
            }
        )

    apu_header = {
        "Item": item_code,
        "Actividad": actividad,
        "Contratista": contratista,
        "Unidad_APU": unidad_apu,
        "Cantidad_APU": cantidad_apu,
        "Costo_Unit_APU": costo_unit_apu,
        "Total_APU": total_apu,
        "Observaciones": " | ".join(dict.fromkeys(observaciones)),
    }
    return apu_header, resources, leyes_sociales_rows


@st.cache_data(show_spinner="Analizando presupuesto y 191 hojas APU con OpenPyXL...")
def load_excel_data(file_bytes):
    """Carga completa del libro Excel con OpenPyXL en modo lectura rápida."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    apu_sheet_names = [s for s in wb.sheetnames if s not in SPECIAL_SHEETS]
    apu_sheet_set = set(apu_sheet_names)

    # 1. Presupuesto general
    meta = {"proyecto": "Proyecto de Construcción", "ubicacion": "-", "contratista": "-"}
    df_ppto = pd.DataFrame()
    resumen_ppto = {}

    if "PPTO MODIFICADO" in wb.sheetnames:
        meta, df_ppto, resumen_ppto = parse_ppto_sheet(wb["PPTO MODIFICADO"], apu_sheet_set)

    ppto_lookup = (
        df_ppto.set_index("Item").to_dict(orient="index") if not df_ppto.empty else {}
    )

    # 2. Catálogos maestros
    catalogs = parse_catalog_sheets(wb)

    # 3. Extracción de todas las hojas APU
    all_headers = []
    all_resources = []
    all_ls = []

    for sheet_name in apu_sheet_names:
        ws = wb[sheet_name]
        hdr, res, ls = parse_apu_sheet(ws, ppto_lookup)
        all_headers.append(hdr)
        all_resources.extend(res)
        all_ls.extend(ls)
        if hdr["Contratista"] and meta.get("contratista", "-") == "-":
            meta["contratista"] = hdr["Contratista"]

    wb.close()

    df_apu_headers = pd.DataFrame(all_headers)
    df_resources = pd.DataFrame(all_resources)
    df_ls = pd.DataFrame(all_ls)

    if not df_apu_headers.empty:
        df_apu_headers["_sort"] = df_apu_headers["Item"].apply(item_sort_key)
        df_apu_headers = df_apu_headers.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)

    return meta, df_ppto, resumen_ppto, catalogs, df_apu_headers, df_resources, df_ls


# ==============================================================================
# 4. EXPORTADOR A EXCEL CON OPENPYXL (CON ESTILO PROFESIONAL)
# ==============================================================================
def build_excel_export(df_consolidado, df_detalle, df_ppto, meta):
    """Genera un archivo .xlsx profesional usando OpenPyXL."""
    wb = openpyxl.Workbook()

    # Estilos en paleta Azul / Morado / Celeste
    header_fill = PatternFill("solid", fgColor="4C1D95")  # Morado oscuro
    zebra_fill = PatternFill("solid", fgColor="F8FAFC")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="1E3A8A")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    def write_sheet(ws, title, df):
        ws.title = title
        ws["A1"] = f"{meta.get('proyecto', 'Planificación')} — {title}"
        ws["A1"].font = title_font
        ws["A2"] = f"Ubicación: {meta.get('ubicacion', '-')} | Contratista: {meta.get('contratista', '-')}"
        ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="64748B")

        start_row = 4
        for c_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=start_row, column=c_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[start_row].height = 26

        for r_offset, row_vals in enumerate(df.itertuples(index=False), start=1):
            r_num = start_row + r_offset
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_num, column=c_idx, value=val)
                cell.border = thin_border
                if r_offset % 2 == 0:
                    cell.fill = zebra_fill
                col_lower = df.columns[c_idx - 1].lower()
                if isinstance(val, (int, float)):
                    if any(k in col_lower for k in ["costo", "total", "precio", "valor", "$"]):
                        cell.number_format = "$ #,##0"
                    else:
                        cell.number_format = "#,##0.00"

        # Autoajuste de columnas
        for col in ws.columns:
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 3, 12), 48)
        ws.freeze_panes = "A5"

    ws1 = wb.active
    write_sheet(ws1, "Recursos Consolidados", df_consolidado)

    ws2 = wb.create_sheet()
    write_sheet(ws2, "Detalle por Partida", df_detalle)

    if not df_ppto.empty:
        ws3 = wb.create_sheet()
        write_sheet(ws3, "Presupuesto", df_ppto)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ==============================================================================
# 5. INTERFAZ PRINCIPAL STREAMLIT
# ==============================================================================
with st.sidebar:
    st.markdown("## 🏗️ PlanificaObra")
    st.caption("Control de Presupuesto, APU y Explosión de Recursos")
    st.markdown("---")

    uploaded_file = st.file_uploader(
        "📂 Cargar otro archivo Excel (.xlsx) (opcional)",
        type=["xlsx"],
        help=f"Por defecto se carga automáticamente '{DEFAULT_EXCEL_FILENAME}' desde la carpeta del proyecto.",
    )

    if uploaded_file is not None:
        st.success(f"📄 Usando archivo subido: {uploaded_file.name}")
    elif DEFAULT_EXCEL_PATH.exists():
        st.success(f"✅ Cargado por defecto: {DEFAULT_EXCEL_FILENAME}")
    else:
        st.warning(f"No se encontró {DEFAULT_EXCEL_FILENAME} en la carpeta actual.")

    st.markdown("### ⚙️ Parámetros de Planificación")
    qty_source = st.radio(
        "Base de cubicación para cantidades:",
        options=["PPTO MODIFICADO (Oficial)", "Cantidad declarada en cada APU"],
        index=0,
        help="Permite resolver diferencias de cubicación entre el presupuesto general y las hojas APU.",
    )
    incluir_perdida = st.toggle("Aplicar factor de pérdida en cantidades", value=True)

    st.markdown("---")
    st.caption("Motor: Python • OpenPyXL • Streamlit")

# Seleccionar fuente de datos: archivo subido o archivo local por defecto
excel_bytes_input = None
active_filename = DEFAULT_EXCEL_FILENAME

if uploaded_file is not None:
    excel_bytes_input = uploaded_file.getvalue()
    active_filename = uploaded_file.name
elif DEFAULT_EXCEL_PATH.exists():
    excel_bytes_input = DEFAULT_EXCEL_PATH.read_bytes()
    active_filename = DEFAULT_EXCEL_FILENAME

# Pantalla inicial si no se encuentra el archivo por defecto ni se subió ninguno
if excel_bytes_input is None:
    st.markdown(
        """
        <div class="hero-banner">
            <h1>Plataforma de Planificación y Explosión de Recursos</h1>
            <p>Carga tu presupuesto en Excel con itemizado y Análisis de Precios Unitarios (APU) para consolidar automáticamente materiales, mano de obra, maquinaria y subcontratos.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown('<div class="kpi-card blue"><div class="kpi-label">Paso 1</div><div class="kpi-value">Sube tu Excel</div><div class="kpi-sub">Compatible con PPTO + hojas APU</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="kpi-card purple"><div class="kpi-label">Paso 2</div><div class="kpi-value">Explosión APU</div><div class="kpi-sub">Lectura automática con OpenPyXL</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="kpi-card pink"><div class="kpi-label">Paso 3</div><div class="kpi-value">Cantidades Reales</div><div class="kpi-sub">Con o sin factor de pérdida</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="kpi-card sky"><div class="kpi-label">Paso 4</div><div class="kpi-value">Exportación</div><div class="kpi-sub">Reporte .xlsx listo para compras</div></div>', unsafe_allow_html=True)
    st.info(f"👈 Utiliza el panel izquierdo para subir **{DEFAULT_EXCEL_FILENAME}** u otro presupuesto con la misma estructura.")
    st.stop()

# Cargar y procesar el Excel con OpenPyXL
meta, df_ppto, resumen_ppto, catalogs, df_apu_headers, df_resources, df_ls = load_excel_data(
    excel_bytes_input
)

# Calcular cantidades y costos totales según la base seleccionada
if not df_resources.empty:
    qty_col = "Cantidad_PPTO" if "PPTO" in qty_source else "Cantidad_APU"
    df_calc = df_resources.copy()
    df_calc["Cantidad_Base_Partida"] = df_calc[qty_col]
    df_calc["Cantidad_Neta"] = df_calc["Cantidad_Base_Partida"] * df_calc["Rendimiento"]
    df_calc["Cantidad_Con_Perdida"] = (
        df_calc["Cantidad_Base_Partida"] * df_calc["Rendimiento"] * df_calc["Factor_Perdida"]
    )
    df_calc["Cantidad_Planificada"] = (
        df_calc["Cantidad_Con_Perdida"] if incluir_perdida else df_calc["Cantidad_Neta"]
    )
    df_calc["Costo_Total_Recurso"] = df_calc["Cantidad_Base_Partida"] * df_calc["Total_Unitario"]
    # Clave limpia para consolidar recursos iguales (normalizando mayúsculas y espacios)
    df_calc["Recurso_Norm"] = df_calc["Recurso"].str.strip().str.upper()
else:
    df_calc = pd.DataFrame()

# Banner superior del proyecto cargado
st.markdown(
    f"""
    <div class="hero-banner">
        <h1>{meta.get("proyecto", "Proyecto sin nombre")}</h1>
        <p>📍 {meta.get("ubicacion", "-")} &nbsp;|&nbsp; 🏢 Contratista: <b>{meta.get("contratista", "-")}</b> &nbsp;|&nbsp; 📄 {len(df_apu_headers)} hojas APU procesadas ({active_filename})</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# KPIs ejecutivos
costo_directo = resumen_ppto.get("DIRECTO", {}).get(
    "monto", df_ppto["Total_PPTO"].sum() if not df_ppto.empty else 0.0
)
total_presupuesto = resumen_ppto.get("TOTAL", {}).get("monto", costo_directo)
total_recursos_unicos = (
    df_calc[["Categoria", "Codigo", "Recurso_Norm", "Unidad_Recurso"]].drop_duplicates().shape[0]
    if not df_calc.empty
    else 0
)

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(
        f'<div class="kpi-card blue"><div class="kpi-label">Costo Directo (PPTO)</div>'
        f'<div class="kpi-value">{fmt_clp(costo_directo)}</div>'
        f'<div class="kpi-sub">{len(df_ppto)} partidas en itemizado</div></div>',
        unsafe_allow_html=True,
    )
with k2:
    st.markdown(
        f'<div class="kpi-card purple"><div class="kpi-label">Presupuesto Total (c/ IVA)</div>'
        f'<div class="kpi-value">{fmt_clp(total_presupuesto)}</div>'
        f'<div class="kpi-sub">Incluye GG, Utilidad e IVA 19%</div></div>',
        unsafe_allow_html=True,
    )
with k3:
    st.markdown(
        f'<div class="kpi-card pink"><div class="kpi-label">Recursos Consolidados</div>'
        f'<div class="kpi-value">{total_recursos_unicos:,}</div>'
        f'<div class="kpi-sub">Extraídos desde {len(df_apu_headers)} APUs</div></div>',
        unsafe_allow_html=True,
    )
with k4:
    total_insumos_apu = df_calc["Costo_Total_Recurso"].sum() if not df_calc.empty else 0.0
    st.markdown(
        f'<div class="kpi-card sky"><div class="kpi-label">Total Insumos (s/ Leyes Soc.)</div>'
        f'<div class="kpi-value">{fmt_clp(total_insumos_apu)}</div>'
        f'<div class="kpi-sub">Base: {"PPTO" if "PPTO" in qty_source else "APU"}</div></div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ==============================================================================
# 6. PESTAÑAS DE NAVEGACIÓN
# ==============================================================================
tab_recursos, tab_ppto, tab_apu, tab_catalogos = st.tabs(
    [
        "📦 Explosión de Recursos (Planificación)",
        "📋 Presupuesto e Itemizado",
        "🔍 Detalle APU por Partida",
        "📚 Catálogos Maestros del Excel",
    ]
)

# ------------------------------------------------------------------------------
# TAB 1: EXPLOSIÓN Y CONSOLIDACIÓN DE RECURSOS
# ------------------------------------------------------------------------------
with tab_recursos:
    st.markdown('<div class="section-title">Listado Consolidado de Recursos y Cantidades</div>', unsafe_allow_html=True)

    if df_calc.empty:
        st.warning("No se encontraron recursos en las hojas APU.")
    else:
        f1, f2, f3 = st.columns([2, 2, 2])
        with f1:
            cats_avail = ["Materiales", "Mano de Obra", "Maquinaria y Equipos", "Subcontratos y Otros"]
            cats_present = [c for c in cats_avail if c in df_calc["Categoria"].unique()]
            sel_cats = st.multiselect("Filtrar por Categoría:", options=cats_present, default=cats_present)
        with f2:
            caps_avail = sorted(df_calc["Capitulo"].unique(), key=item_sort_key)
            sel_caps = st.multiselect("Filtrar por Capítulo:", options=caps_avail, default=caps_avail)
        with f3:
            search_q = st.text_input("🔎 Buscar recurso o código:", placeholder="Ej: Hormigón, MAT-3, Jornal...")

        df_filtered = df_calc[
            df_calc["Categoria"].isin(sel_cats) & df_calc["Capitulo"].isin(sel_caps)
        ].copy()

        if search_q:
            q_low = search_q.strip().lower()
            df_filtered = df_filtered[
                df_filtered["Recurso"].str.lower().str.contains(q_low, na=False)
                | df_filtered["Codigo"].str.lower().str.contains(q_low, na=False)
                | df_filtered["Actividad"].str.lower().str.contains(q_low, na=False)
            ]

        # Agrupación ordenada de recursos
        df_consolidado = (
            df_filtered.groupby(["Categoria", "Codigo", "Recurso_Norm", "Unidad_Recurso"], as_index=False)
            .agg(
                Recurso=("Recurso", "first"),
                Partidas=("Item", "nunique"),
                Cantidad_Neta=("Cantidad_Neta", "sum"),
                Cantidad_Con_Perdida=("Cantidad_Con_Perdida", "sum"),
                Cantidad_Planificada=("Cantidad_Planificada", "sum"),
                Costo_Unit_Prom=("Costo_Unitario", "mean"),
                Costo_Total=("Costo_Total_Recurso", "sum"),
            )
            .drop(columns=["Recurso_Norm"])
        )

        # Ordenamiento lógico: por categoría fija y luego mayor impacto económico
        cat_order = {c: i for i, c in enumerate(cats_avail)}
        df_consolidado["_cat_ord"] = df_consolidado["Categoria"].map(cat_order).fillna(99)
        df_consolidado = (
            df_consolidado.sort_values(["_cat_ord", "Costo_Total"], ascending=[True, False])
            .drop(columns=["_cat_ord"])
            .reset_index(drop=True)
        )

        # Gráficos interactivos
        g1, g2 = st.columns([1.1, 1.4])
        color_map = {
            "Materiales": PALETTE["purple"],
            "Mano de Obra": PALETTE["navy"],
            "Maquinaria y Equipos": PALETTE["sky"],
            "Subcontratos y Otros": PALETTE["pink"],
        }
        with g1:
            df_cat_sum = df_consolidado.groupby("Categoria", as_index=False)["Costo_Total"].sum()
            fig_pie = px.pie(
                df_cat_sum,
                names="Categoria",
                values="Costo_Total",
                hole=0.52,
                color="Categoria",
                color_discrete_map=color_map,
                title="Distribución de Costos por Tipo de Recurso",
            )
            fig_pie.update_layout(margin=dict(t=45, b=10, l=10, r=10), paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_pie, use_container_width=True)

        with g2:
            top10 = df_consolidado.nlargest(10, "Costo_Total").sort_values("Costo_Total", ascending=True)
            fig_bar = px.bar(
                top10,
                x="Costo_Total",
                y="Recurso",
                color="Categoria",
                orientation="h",
                color_discrete_map=color_map,
                title="Top 10 Recursos por Incidencia Económica ($)",
                labels={"Costo_Total": "Costo Total ($)", "Recurso": ""},
            )
            fig_bar.update_layout(
                margin=dict(t=45, b=10, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        # Tabla principal ordenada
        st.markdown(f"**Mostrando {len(df_consolidado)} recursos consolidados** (ordenados por categoría y costo total):")
        df_view = df_consolidado[
            [
                "Categoria",
                "Codigo",
                "Recurso",
                "Unidad_Recurso",
                "Cantidad_Neta",
                "Cantidad_Con_Perdida",
                "Partidas",
                "Costo_Unit_Prom",
                "Costo_Total",
            ]
        ].rename(
            columns={
                "Categoria": "Categoría",
                "Codigo": "Código",
                "Unidad_Recurso": "Unidad",
                "Cantidad_Neta": "Cant. Neta",
                "Cantidad_Con_Perdida": "Cant. Total (c/ Pérd.)",
                "Partidas": "N° Partidas",
                "Costo_Unit_Prom": "Precio Unit. Prom ($)",
                "Costo_Total": "Costo Total ($)",
            }
        )

        st.dataframe(
            df_view,
            use_container_width=True,
            height=460,
            hide_index=True,
            column_config={
                "Cant. Neta": st.column_config.NumberColumn(format="%.2f"),
                "Cant. Total (c/ Pérd.)": st.column_config.NumberColumn(format="%.2f"),
                "N° Partidas": st.column_config.NumberColumn(format="%d"),
                "Precio Unit. Prom ($)": st.column_config.NumberColumn(format="$ %,.0f"),
                "Costo Total ($)": st.column_config.NumberColumn(format="$ %,.0f"),
            },
        )

        # Trazabilidad de un recurso específico
        with st.expander("🔎 Trazabilidad: ver en qué partidas se utiliza un recurso específico"):
            opciones_rec = df_consolidado["Recurso"].unique().tolist()
            if opciones_rec:
                rec_sel = st.selectbox("Selecciona un recurso:", opciones_rec)
                df_trace = df_filtered[df_filtered["Recurso"] == rec_sel][
                    [
                        "Item",
                        "Capitulo",
                        "Actividad",
                        "Unidad_Partida",
                        "Cantidad_Base_Partida",
                        "Rendimiento",
                        "Factor_Perdida",
                        "Cantidad_Planificada",
                        "Costo_Total_Recurso",
                    ]
                ].rename(
                    columns={
                        "Capitulo": "Capítulo",
                        "Unidad_Partida": "Und. Partida",
                        "Cantidad_Base_Partida": "Cant. Partida",
                        "Factor_Perdida": "Factor Pérd.",
                        "Cantidad_Planificada": "Cant. Recurso",
                        "Costo_Total_Recurso": "Costo Total ($)",
                    }
                )
                st.dataframe(df_trace, use_container_width=True, hide_index=True)

        # Botón de exportación a Excel con OpenPyXL
        excel_bytes = build_excel_export(
            df_view,
            df_filtered[
                [
                    "Item",
                    "Capitulo",
                    "Actividad",
                    "Categoria",
                    "Codigo",
                    "Recurso",
                    "Unidad_Recurso",
                    "Cantidad_Base_Partida",
                    "Rendimiento",
                    "Factor_Perdida",
                    "Cantidad_Planificada",
                    "Costo_Unitario",
                    "Costo_Total_Recurso",
                ]
            ],
            df_ppto,
            meta,
        )
        st.download_button(
            label="📥 Descargar Planificación Consolidada (.xlsx con OpenPyXL)",
            data=excel_bytes,
            file_name="Planificacion_Recursos_Consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

# ------------------------------------------------------------------------------
# TAB 2: PRESUPUESTO E ITEMIZADO
# ------------------------------------------------------------------------------
with tab_ppto:
    st.markdown('<div class="section-title">Itemizado General — PPTO MODIFICADO</div>', unsafe_allow_html=True)

    if df_ppto.empty:
        st.warning("No se encontró la hoja 'PPTO MODIFICADO'.")
    else:
        c_chart, c_res = st.columns([1.6, 1.0])
        with c_chart:
            df_cap = df_ppto.groupby("Capitulo", as_index=False)["Total_PPTO"].sum()
            df_cap["_s"] = df_cap["Capitulo"].apply(item_sort_key)
            df_cap = df_cap.sort_values("_s")
            fig_cap = px.bar(
                df_cap,
                x="Capitulo",
                y="Total_PPTO",
                color_discrete_sequence=[PALETTE["indigo"]],
                title="Costo Directo por Capítulo ($)",
                labels={"Capitulo": "Capítulo", "Total_PPTO": "Monto ($)"},
            )
            fig_cap.update_layout(
                xaxis_tickangle=-25,
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_cap, use_container_width=True)

        with c_res:
            st.markdown("#### Cierre Presupuestario")
            res_rows = []
            for label in ["DIRECTO", "GG", "UTIL", "NETO", "IVA", "TOTAL"]:
                if label in resumen_ppto:
                    pct = resumen_ppto[label]["porcentaje"]
                    res_rows.append(
                        {
                            "Concepto": label,
                            "%": f"{pct * 100:.2f}%" if pct is not None else "—",
                            "Monto ($)": fmt_clp(resumen_ppto[label]["monto"]),
                        }
                    )
            if res_rows:
                st.table(pd.DataFrame(res_rows))

        st.dataframe(
            df_ppto.drop(columns=["Capitulo_ID"]),
            use_container_width=True,
            height=480,
            hide_index=True,
            column_config={
                "Cantidad_PPTO": st.column_config.NumberColumn("Cantidad", format="%.2f"),
                "Precio_Unitario_PPTO": st.column_config.NumberColumn("Precio Unitario", format="$ %,.0f"),
                "Total_PPTO": st.column_config.NumberColumn("Total", format="$ %,.0f"),
                "Tiene_APU": st.column_config.CheckboxColumn("APU Vinculado"),
            },
        )

# ------------------------------------------------------------------------------
# TAB 3: DETALLE APU POR PARTIDA
# ------------------------------------------------------------------------------
with tab_apu:
    st.markdown('<div class="section-title">Inspector de Análisis de Precios Unitarios (APU)</div>', unsafe_allow_html=True)

    if df_apu_headers.empty:
        st.warning("No se detectaron hojas de APU.")
    else:
        df_apu_headers["Etiqueta"] = df_apu_headers["Item"] + " — " + df_apu_headers["Actividad"]
        sel_label = st.selectbox("Selecciona una partida (APU):", df_apu_headers["Etiqueta"].tolist())
        sel_item = sel_label.split(" — ")[0]

        hdr = df_apu_headers[df_apu_headers["Item"] == sel_item].iloc[0]
        ppto_match = df_ppto[df_ppto["Item"] == sel_item] if not df_ppto.empty else pd.DataFrame()
        cant_ppto_item = float(ppto_match["Cantidad_PPTO"].iloc[0]) if not ppto_match.empty else hdr["Cantidad_APU"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ítem", sel_item)
        m2.metric("Unidad", hdr["Unidad_APU"])
        m3.metric("Cantidad (PPTO / APU)", f"{fmt_num(cant_ppto_item)} / {fmt_num(hdr['Cantidad_APU'])}")
        m4.metric("Costo Unitario Directo", fmt_clp(hdr["Costo_Unit_APU"]))

        if abs(cant_ppto_item - hdr["Cantidad_APU"]) > 0.01:
            st.warning(
                f"⚠️ Diferencia detectada: En `PPTO MODIFICADO` la cantidad es **{fmt_num(cant_ppto_item)}**, "
                f"mientras que en la hoja APU `{sel_item}` figura **{fmt_num(hdr['Cantidad_APU'])}**."
            )

        if hdr["Observaciones"]:
            st.info(f"📝 **Observaciones técnicas del Excel:** {hdr['Observaciones']}")

        df_item_res = df_calc[df_calc["Item"] == sel_item].copy() if not df_calc.empty else pd.DataFrame()
        for cat in ["Materiales", "Maquinaria y Equipos", "Mano de Obra", "Subcontratos y Otros"]:
            sub = df_item_res[df_item_res["Categoria"] == cat] if not df_item_res.empty else pd.DataFrame()
            st.markdown(f"**{cat}**")
            if sub.empty:
                st.caption("Sin recursos registrados en esta sección.")
            else:
                st.dataframe(
                    sub[
                        [
                            "Codigo",
                            "Recurso",
                            "Rendimiento",
                            "Unidad_Recurso",
                            "Factor_Perdida",
                            "Costo_Unitario",
                            "Total_Unitario",
                            "Cantidad_Planificada",
                            "Costo_Total_Recurso",
                        ]
                    ].rename(
                        columns={
                            "Codigo": "Código",
                            "Unidad_Recurso": "Unidad",
                            "Factor_Perdida": "Factor Pérd.",
                            "Costo_Unitario": "Costo Unit. ($)",
                            "Total_Unitario": "Total Unit. ($)",
                            "Cantidad_Planificada": "Cant. Total Obra",
                            "Costo_Total_Recurso": "Costo Total Obra ($)",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

        # Mostrar Leyes Sociales del ítem si existen
        ls_item = df_ls[df_ls["Item"] == sel_item] if not df_ls.empty else pd.DataFrame()
        if not ls_item.empty:
            ls_row = ls_item.iloc[0]
            st.caption(
                f"💼 **Leyes Sociales (Mano de Obra):** {ls_row['Porcentaje_LS'] * 100:.2f}% "
                f"→ Costo unitario: {fmt_clp(ls_row['Costo_Unit_LS'])} | "
                f"Total partida: {fmt_clp(ls_row['Costo_Unit_LS'] * cant_ppto_item)}"
            )

# ------------------------------------------------------------------------------
# TAB 4: CATÁLOGOS MAESTROS DEL EXCEL
# ------------------------------------------------------------------------------
with tab_catalogos:
    st.markdown('<div class="section-title">Bases de Precios Maestras del Libro Excel</div>', unsafe_allow_html=True)
    cat_choice = st.radio(
        "Seleccionar catálogo:",
        options=[k for k in ["MATERIALES", "MANO DE OBRA", "MAQUINARIA"] if k in catalogs],
        horizontal=True,
    )
    df_cat_master = catalogs.get(cat_choice, pd.DataFrame())
    if not df_cat_master.empty:
        q_cat = st.text_input("Filtrar en catálogo maestro:", key="cat_search")
        if q_cat:
            mask = df_cat_master.astype(str).apply(
                lambda col: col.str.lower().str.contains(q_cat.strip().lower(), na=False)
            ).any(axis=1)
            df_cat_master = df_cat_master[mask]
        st.dataframe(df_cat_master, use_container_width=True, height=500, hide_index=True)
