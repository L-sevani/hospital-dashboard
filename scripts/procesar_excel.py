"""
procesar_excel.py  —  HGZ 1 TLAXCALA
--------------------------------------
Descarga el Excel de IAAS desde SharePoint,
expande los microorganismos y genera el dashboard HTML.

Variables de entorno requeridas (GitHub Secrets):
    SP_TENANT_ID      → ID del tenant de Microsoft 365
    SP_CLIENT_ID      → ID de la app registrada en Azure
    SP_CLIENT_SECRET  → Secreto de la app
    SP_SITE_ID        → ID del sitio de SharePoint
    SP_FILE_PATH      → Ruta del Excel en SharePoint
                        Ej: /sites/Bacteriologia/Shared Documents/reporte_iaas.xlsx
"""

import os, sys, json, re, requests, tempfile, shutil
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════
# CATÁLOGOS DE MAPEO
# ══════════════════════════════════════════════════════

INFECCION_MAP = {
    "NEUMONIA ASOCIADA A VENTILADOR": "NEUMONÍA",
    "NEUMONIA CLINICA NO ASOCIADA A VENTILADOR": "NEUMONÍA",
    "INFECCION DE HERIDA QUIRURGICA INCISIONAL SUPERFICIAL": "HERIDAS",
    "INFECCION DE HERIDA QUIRURGICA INCISIONAL PROFUNDA": "HERIDAS",
    "INFECCION DE HERIDA QUIRURGICA DE ORGANOS Y ESPACIOS": "HERIDAS",
    "INFECCIONES DE SITIO DE INSERCION DE CATETER, TUNEL O PUERTO SUBCUTANEO": "ITS",
    "ITS RELACIONADA A CATETER VENOSO CENTRAL": "ITS",
    "ITS CONFIRMADA POR LABORATORIO": "ITS",
    "BACTERIEMIA PRIMARIA NO DEMOSTRADA": "ITS",
    "ITS SECUNDARIA A PROCEDIMIENTOS (CISTOSCOPIAS Y COLANGIOGRAFIAS)": "ITS",
    "ENDOCARDITIS": "ITS",
    "ITS RELACIONADA A CONTAMINACION DE SOLUCIONES": "ITS",
    "INFUSIONES O MEDICAMENTOS INTRAVENOSOS": "ITS",
    "ITS RELACIONADA A CONTAMINACION DE SOLUCIONES, INFUSIONES O MEDICAMENTOS INTRAVENOSOS": "ITS",
    "IVU ASOCIADA A SONDA VESICAL": "IVU",
    "IVU NO ASOCIADA A SONDA VESICAL": "IVU",
    "PERITONITIS ASOCIADA A DIALISIS": "PERITONITIS",
    "PERITONITIS ASOCIADA A LA INSTALACION DE CATETER DE DIALISIS PERITONEAL": "PERITONITIS",
    "INFECCIONES DE PIEL Y TEJIDOS BLANDOS": "RESTO DE IAAS",
    "INFECCIONES DE PIEL Y TEJIDOS BLANDOS EN PACIENTES CON QUEMADURAS": "RESTO DE IAAS",
    "FLEBITIS": "RESTO DE IAAS",
    "CONJUNTIVITIS": "RESTO DE IAAS",
    "OTRO": "RESTO DE IAAS",
    "RINOFARINGITIS Y FARINGOAMIGDALITIS (IVRA)": "RESTO DE IAAS",
    "COVID-19": "RESTO DE IAAS",
    "ENDOMETRITIS": "RESTO DE IAAS",
    "BRONQUITIS": "RESTO DE IAAS",
    "TRAQUEOBRONQUITIS": "RESTO DE IAAS",
    "TRAQUEITIS SIN DATOS DE NEUMONIA": "RESTO DE IAAS",
    "INFECCIONES DE LA BURSA O ARTICULARES": "RESTO DE IAAS",
    "OTITIS MEDIA AGUDA": "RESTO DE IAAS",
    "INFECCIONES RELACIONADAS A PROCEDIMIENTOS ODONTOLOGICOS": "RESTO DE IAAS",
    "FASCITIS NECROSANTE": "RESTO DE IAAS",
    "GANGRENA INFECCIOSA": "RESTO DE IAAS",
    "CELULITIS": "RESTO DE IAAS",
    "MIOSITIS Y LINFADENITIS": "RESTO DE IAAS",
    "GASTROENTERITIS NOSOCOMIAL": "RESTO DE IAAS",
    "MENINGITIS O VENTRICULITIS": "RESTO DE IAAS",
    "INFLUENZA VIRUS": "RESTO DE IAAS",
}

CULTIVO_MAP = {
    "HEMOCULTIVO DURANTE PICO FEBRIL": "HEMOCULTIVOS",
    "SOLO HEMOCULTIVO CENTRAL": "HEMOCULTIVOS",
    "CULTIVO DE PUNTA DE CATETER + PERIFERICO": "HEMOCULTIVOS",
    "HEMOCULTIVO SIN PICO FEBRIL": "HEMOCULTIVOS",
    "SOLO HEMOCULTIVO PERIFERICO": "HEMOCULTIVOS",
    "HEMOCULTIVO CENTRAL + PERIFERICO": "HEMOCULTIVOS",
    "HEMOCULTIVO PERIFERICO (1 BRAZO) - INADECUADO": "HEMOCULTIVOS",
    "HEMOCULTIVOS PERIFERICOS (2 BRAZOS)": "HEMOCULTIVOS",
    "HEMOCULTIVOS": "HEMOCULTIVOS",
    "SOLO CULTIVO DE PUNTA DE CATETER. - INADECUADO": "HEMOCULTIVOS",
    "UROCULTIVO POR PUNCION DE SONDA": "UROCULTIVOS",
    "UROCULTIVO POR CHORRO MEDIO": "UROCULTIVOS",
    "CULTIVO DE LA SONDA URETRAL - INADECUADO": "UROCULTIVOS",
    "UROCULTIVO POR PUNCION SUPRAPUBICA": "UROCULTIVOS",
    "CULTIVO DE ESPUTO": "NEUMO",
    "CULTIVO DE LIQUIDO PLEUREAL": "NEUMO",
    "ASPIRADO TRAQUEAL": "NEUMO",
    "CULTIVO POR PUNCION ASPIRACION": "NEUMO",
    "LAVADO BRONQUIAL": "NEUMO",
    "CULTIVO DE LIQUIDO PERITONEAL": "PERITONEAL",
    "CULTIVO POR HISOPO - INADECUADO": "OTROS",
    "CULTIVO DEL PUS - INADECUADO": "OTROS",
    "COPROCULTIVO": "OTROS",
    "PONER TODAS LAS OPCIONES": "OTROS",
    "CULTIVO DE DRENAJES - INADECUADO": "OTROS",
    "CULTIVO DE LIQUIDO ESTERIL": "OTROS",
    "EXUDADO NASOFARINGEO": "OTROS",
    "CULTIVOS TOMADOS DE LA BOLSA RECOLECTORA DE ORINA - INADECUADO": "OTROS",
    "CULTIVO DE LCR": "OTROS",
    "OTROS - INADECUADO": "OTROS",
    "CULTIVO DE BIOPSIA": "OTROS",
}

SERVICIO_MAP = {
    "GINECOLOGIA": "GINECOLOGÍA", "OBSTETRICIA": "GINECOLOGÍA",
    "GINECOLOGIA Y OBSTETRICIA": "GINECOLOGÍA",
    "ONCOLOGIA GINECOLOGICA": "GINECOLOGÍA",
    "ONCOLOGIA DE TUMORES DE MAMA": "GINECOLOGÍA",
    "MEDICINA INTERNA": "MEDICINA INTERNA", "NEFROLOGIA": "MEDICINA INTERNA",
    "HEMODIALISIS": "MEDICINA INTERNA", "MEDICINA FAMILIAR": "MEDICINA INTERNA",
    "CARDIOLOGIA": "MEDICINA INTERNA", "GERONTOLOGIA/GERIATRIA": "MEDICINA INTERNA",
    "NEUROLOGIA": "MEDICINA INTERNA", "AUDIOLOGIA": "MEDICINA INTERNA",
    "REUMATOLOGIA": "MEDICINA INTERNA", "NEUMOLOGIA": "MEDICINA INTERNA",
    "NEFROLOGIA PEDIATRICA": "MEDICINA INTERNA",
    "PSIQUIATRIA CLINICA": "MEDICINA INTERNA",
    "MEDICINA INTERNA PEDIATRICA": "MEDICINA INTERNA",
    "CIRUGIA GENERAL": "CIRUGÍA", "ORTOPEDIA Y TRAUMATOLOGIA": "CIRUGÍA",
    "TRAUMATOLOGIA": "CIRUGÍA", "UROLOGIA": "CIRUGÍA",
    "OFTALMOLOGIA": "CIRUGÍA", "ORTOPEDIA": "CIRUGÍA",
    "CIRUGIA PEDIATRICA": "CIRUGÍA", "OTORRINOLARINGOLOGIA": "CIRUGÍA",
    "CIRUGIA AMBULATORIA": "CIRUGÍA", "ANGIOLOGIA": "CIRUGÍA",
    "CIRUGIA PLASTICA RECONSTRUCTIVA": "CIRUGÍA",
    "TRUMATOLOGIA Y ORTOPEDIA DE CADERA Y PELVIS": "CIRUGÍA",
    "ONCOLOGIA QUIRURGICA": "CIRUGÍA", "CIRUGIA MAXILOFACIAL": "CIRUGÍA",
    "CIRUGIA CARDIOVASCULAR": "CIRUGÍA",
    "TRAUMATOLOGIA Y O. DE CIRUGIA TORACO ABDOMINAL": "CIRUGÍA",
    "TRAUMATOLOGIA Y ORTOPEDIA DE MIEMBRO PELVICO": "CIRUGÍA",
    "TRAUMATOLOGIA Y ORTOPEDIA DE MIEMBRO TORACICO": "CIRUGÍA",
    "NEUROCIRUGIA": "CIRUGÍA",
    "CUNERO PATOLOGICO": "PEDIATRÍA", "PEDIATRIA MEDICA": "PEDIATRÍA",
    "MED. ENFERMO EDO. CRITICO": "UCI",
    "URGENCIAS": "URGENCIAS",
}

# ══════════════════════════════════════════════════════
# SHAREPOINT — DESCARGA DEL EXCEL
# ══════════════════════════════════════════════════════

def obtener_token_sharepoint():
    tenant  = os.environ["SP_TENANT_ID"]
    client  = os.environ["SP_CLIENT_ID"]
    secret  = os.environ["SP_CLIENT_SECRET"]
    url     = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    resp    = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     client,
        "client_secret": secret,
        "scope":         "https://graph.microsoft.com/.default",
    })
    resp.raise_for_status()
    return resp.json()["access_token"]

def descargar_excel_sharepoint(destino):
    token     = obtener_token_sharepoint()
    site_id   = os.environ["SP_SITE_ID"]
    file_path = os.environ["SP_FILE_PATH"]   # Ej: /sites/Bacteriologia/Shared Documents/reporte.xlsx
    headers   = {"Authorization": f"Bearer {token}"}

    # Obtener metadata del archivo
    enc_path = requests.utils.quote(file_path)
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{enc_path}:/content"
    resp = requests.get(url, headers=headers, allow_redirects=True)
    resp.raise_for_status()
    with open(destino, "wb") as f:
        f.write(resp.content)
    print(f"✅ Excel descargado desde SharePoint → {destino}")

# ══════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════

def normalizar(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip().upper()

def mapear_infeccion(texto):
    t = normalizar(texto)
    if t in INFECCION_MAP:
        return INFECCION_MAP[t], t
    if t.startswith("ITS"):      return "ITS", t
    if t.startswith("IVU"):      return "IVU", t
    if "NEUMONIA" in t:          return "NEUMONÍA", t
    if "HERIDA QUIRURGICA" in t: return "HERIDAS", t
    if "PERITONITIS" in t:       return "PERITONITIS", t
    return "RESTO DE IAAS", t

def mapear_cultivo(texto):
    t = normalizar(texto)
    if t in CULTIVO_MAP: return CULTIVO_MAP[t]
    if "HEMOCULTIVO" in t:  return "HEMOCULTIVOS"
    if "UROCULTIVO" in t:   return "UROCULTIVOS"
    if any(x in t for x in ["ESPUTO","BRONQUIAL","TRAQUEAL","PLEUREAL"]): return "NEUMO"
    if "PERITONEAL" in t:   return "PERITONEAL"
    return "OTROS"

def mapear_servicio(texto):
    t = normalizar(texto)
    if t in SERVICIO_MAP: return SERVICIO_MAP[t]
    for k, v in SERVICIO_MAP.items():
        if k in t: return v
    return t

def extraer_periodo(fecha):
    if pd.isna(fecha): return None
    try:
        if hasattr(fecha, "strftime"): return fecha.strftime("%Y-%m")
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try: return datetime.strptime(str(fecha).strip(), fmt).strftime("%Y-%m")
            except: continue
    except: pass
    return None

# ══════════════════════════════════════════════════════
# LECTURA Y EXPANSIÓN
# ══════════════════════════════════════════════════════

def leer_excel(path):
    raw = pd.read_excel(path, header=3)
    real_cols = raw.iloc[0].tolist()
    raw.columns = real_cols
    raw = raw.iloc[1:].reset_index(drop=True)
    return raw, real_cols

def expandir_microorganismos(raw, real_cols):
    blocks = [
        (41, 42, 43,  44, 112, 113),
        (114,115,116, 117, 185, 186),
        (187,188,189, 190, 258, 259),
        (260,261,262, 263, 331, 332),
    ]
    abx_names = [real_cols[i] for i in range(44, 113)]
    records = []

    for _, row in raw.iterrows():
        base = {real_cols[i]: row.iloc[i] for i in range(41)}
        fecha = row.iloc[16]
        if pd.isna(fecha) or extraer_periodo(fecha) is None:
            fecha = row.iloc[0]
        base["_periodo"] = extraer_periodo(fecha)

        sg_inf, tipo_norm = mapear_infeccion(row.iloc[34])
        base["_subgrupo_infeccion"]  = sg_inf
        base["_tipo_infeccion_norm"] = tipo_norm
        base["_subgrupo_cultivo"]    = mapear_cultivo(row.iloc[38])
        base["_subgrupo_servicio"]   = mapear_servicio(row.iloc[27])

        for num, (ri, rc, rm, as_, ae, mec) in enumerate(blocks, 1):
            micro = row.iloc[rm]
            if pd.isna(micro) or str(micro).strip() in ("","0","nan"):
                continue
            rec = dict(base)
            rec["N_Micro"]     = num
            rec["Resistencia"] = row.iloc[ri]
            rec["Clave_Grupo"] = row.iloc[rc]
            rec["Microorganismo"] = str(micro).strip().upper()
            for j, abx in enumerate(abx_names):
                v = str(row.iloc[as_ + j]).strip().upper() if pd.notna(row.iloc[as_ + j]) else ""
                rec[abx] = v if v in ("S","R","I") else ""
            rec["Mecanismo_Resistencia"] = row.iloc[mec]
            records.append(rec)

    return pd.DataFrame(records), abx_names

# ══════════════════════════════════════════════════════
# GENERACIÓN DEL HTML (inserta JSON en el dashboard)
# ══════════════════════════════════════════════════════

def sri_por_micro(sub, useful_abx):
    result = {}
    for micro, grp in sub.groupby("Microorganismo"):
        abx_data = []
        for abx in useful_abx:
            if abx not in grp.columns: continue
            vals = grp[abx].dropna().str.strip().str.upper()
            s, r, i = int((vals=="S").sum()), int((vals=="R").sum()), int((vals=="I").sum())
            if s + r + i > 0:
                abx_data.append({"a": abx, "S": s, "R": r, "I": i, "t": s+r+i})
        if abx_data:
            result[micro] = {"total": len(grp), "antibioticos": abx_data}
    return result

def construir_raw(df, abx_names):
    useful = [a for a in abx_names if a in df.columns and df[a].isin(["S","R","I"]).any()]
    periodos = sorted(df["_periodo"].dropna().unique())

    INF_SGS  = ["HERIDAS","ITS","IVU","NEUMONÍA","PERITONITIS","RESTO DE IAAS"]
    CULT_SGS = ["HEMOCULTIVOS","UROCULTIVOS","NEUMO","PERITONEAL","OTROS"]
    SERV_SGS = ["GINECOLOGÍA","MEDICINA INTERNA","CIRUGÍA","PEDIATRÍA","UCI","URGENCIAS"]

    data_p = {}
    for p in periodos:
        pf = df[df["_periodo"] == p]
        def section(col, sgs):
            d = {}
            for sg in sgs:
                sub = pf[pf[col] == sg]
                if len(sub): d[sg] = sri_por_micro(sub, useful)
            return {"subgrupos": [sg for sg in sgs if sg in d], "data": d}

        data_p[p] = {
            "infecciones": section("_subgrupo_infeccion", INF_SGS),
            "cultivos":    section("_subgrupo_cultivo",   CULT_SGS),
            "servicios":   section("_subgrupo_servicio",  SERV_SGS),
        }
    return {"periodos": periodos, "data": data_p}

def inyectar_json_en_html(raw_json, html_in, html_out):
    with open(html_in, "r", encoding="utf-8") as f:
        html = f.read()

    nuevo_bloque = "const RAW_PERIODOS = " + json.dumps(raw_json, ensure_ascii=False) + ";"

    # Reemplaza bloque existente si ya existe
    patron = r"const RAW_PERIODOS\s*=\s*\{.*?\};"
    if re.search(patron, html, flags=re.DOTALL):
        html = re.sub(patron, nuevo_bloque, html, flags=re.DOTALL)
    else:
        # Inserta antes del cierre del último script
        html = html.replace("</script>", f"{nuevo_bloque}\n</script>", 1)

    # Actualiza fecha de última actualización si existe el marcador
    fecha_hoy = datetime.now().strftime("%d/%m/%Y %H:%M")
    html = re.sub(
        r'id="ultima-actualizacion">[^<]*<',
        f'id="ultima-actualizacion">{fecha_hoy}<',
        html
    )

    with open(html_out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Dashboard actualizado → {html_out}")

# ══════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════

def main():
    # Determinar fuente del Excel
    modo_sharepoint = all(k in os.environ for k in
                          ["SP_TENANT_ID","SP_CLIENT_ID","SP_CLIENT_SECRET","SP_SITE_ID","SP_FILE_PATH"])

    if modo_sharepoint:
        tmp = tempfile.mktemp(suffix=".xlsx")
        descargar_excel_sharepoint(tmp)
        xlsx_path = tmp
    elif len(sys.argv) > 1:
        xlsx_path = sys.argv[1]
    else:
        print("❌ Indica el archivo Excel: python procesar_excel.py reporte.xlsx")
        sys.exit(1)

    print(f"📂 Procesando: {xlsx_path}")
    raw, real_cols = leer_excel(xlsx_path)
    print(f"   Filas: {len(raw)}")

    df, abx_names = expandir_microorganismos(raw, real_cols)
    print(f"   Microorganismos expandidos: {len(df)}")
    if len(df) == 0:
        print("⚠️  Sin microorganismos. Verifica el archivo.")
        sys.exit(1)

    periodos = sorted(df["_periodo"].dropna().unique())
    print(f"   Periodos: {periodos}")

    raw_json = construir_raw(df, abx_names)

    # Guardar JSON de respaldo
    json_out = os.path.join(os.path.dirname(__file__), "..", "dashboard_data.json")
    with open(json_out, "w", encoding="utf-8") as f:
        json.dump(raw_json, f, ensure_ascii=False, indent=2)
    print(f"✅ JSON guardado: {json_out}")

    # Inyectar en HTML
    base = os.path.join(os.path.dirname(__file__), "..")
    html_in  = os.path.join(base, "dashboard_completo.html")
    html_out = os.path.join(base, "index.html")
    if os.path.exists(html_in):
        inyectar_json_en_html(raw_json, html_in, html_out)
    else:
        print("⚠️  No se encontró dashboard_completo.html — solo se generó el JSON.")

    # Limpiar temporal
    if modo_sharepoint and os.path.exists(xlsx_path):
        os.remove(xlsx_path)

if __name__ == "__main__":
    main()
