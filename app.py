"""
Sistema de Ficha Médica Ocupacional - CASISO S.A.S.
Dr. Martín Lino Zambrano
Versión Cloud — Render.com + Supabase
"""
import os, io, json
from copy import copy
from datetime import date
from flask import Flask, request, jsonify, send_file, Response
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

BASE = os.path.dirname(os.path.abspath(__file__))
PLANTILLAS = os.path.join(BASE, "plantillas")

# ── SUPABASE (base de datos en la nube) ──────────────────────
# En producción (Render) se usan variables de entorno.
# En local funciona con lista en memoria como respaldo.
SUPABASE_URL  = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY  = os.environ.get("SUPABASE_KEY", "")
TABLE_NAME    = "matriz_evaluaciones"

supa = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        from supabase import create_client
        supa = create_client(SUPABASE_URL, SUPABASE_KEY)
        print("✓ Conectado a Supabase")
    except Exception as e:
        print(f"⚠ Supabase no disponible: {e}")
        supa = None

# Fallback local (cuando no hay Supabase configurado)
_matriz_local = []

def db_insertar(fila):
    if supa:
        try:
            supa.table(TABLE_NAME).insert({"data": fila}).execute()
            return
        except Exception as e:
            print(f"⚠ Supabase insert error: {e}")
    _matriz_local.append(fila)

def db_listar():
    if supa:
        try:
            res = supa.table(TABLE_NAME).select("*").order("id").execute()
            supabase_data = [r["data"] for r in res.data]
            # Combinar con locales si hay
            return supabase_data + _matriz_local
        except Exception as e:
            print(f"⚠ Supabase list error: {e}")
    return _matriz_local

def db_contar():
    if supa:
        try:
            res = supa.table(TABLE_NAME).select("id", count="exact").execute()
            return (res.count or 0) + len(_matriz_local)
        except Exception as e:
            print(f"⚠ Supabase count error: {e}")
    return len(_matriz_local)

def db_limpiar():
    global _matriz_local
    if supa:
        try:
            supa.table(TABLE_NAME).delete().neq("id", 0).execute()
        except Exception as e:
            print(f"⚠ Supabase delete error: {e}")
    _matriz_local = []
# ─────────────────────────────────────────────────────────────

def sw(ws, addr, value):
    """safe_write: only write to top-left cells of merged ranges"""
    try:
        ws[addr] = value
    except AttributeError:
        pass  # MergedCell - skip silently

app = Flask(__name__)

# ─────────────────────────────────────────────
#  LLENADO DE PLANTILLAS
# ─────────────────────────────────────────────
def chk(v): return "X" if v else ""
def val(v): return str(v) if v else ""

def llenar_hoja1(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "PRIMERA_HOJA.xlsx"))
    ws = wb.active

    # A. Datos establecimiento (fila 5)
    sw(ws, "A5", d.get("institucion",""))
    sw(ws, "P5", d.get("ruc",""))
    sw(ws, "Y5", d.get("ciiu",""))
    sw(ws, "AB5", d.get("establecimiento",""))
    sw(ws, "AO5", d.get("nroHistoria",""))
    sw(ws, "AT5", d.get("nroArchivo",""))

    # A. Nombres (fila 8)
    sw(ws, "A8", d.get("primerApellido",""))
    sw(ws, "R8", d.get("segundoApellido",""))
    sw(ws, "AG8", d.get("primerNombre",""))
    sw(ws, "AQ8", d.get("segundoNombre",""))

    # A. Grupo atención prioritaria + sexo (fila 13)
    sw(ws, "A13", chk(d.get("embarazada")))
    sw(ws, "F13", chk(d.get("discapacidad")))
    sw(ws, "J13", chk(d.get("enfCatastrofica")))
    sw(ws, "M13", chk(d.get("adultoMayor")))
    sexo = d.get("sexo","")
    sw(ws, "T13", "X" if sexo == "Hombre" else "")
    sw(ws, "W13", "X" if sexo == "Mujer" else "")

    # Fecha nacimiento
    fn = d.get("fechaNacimiento","")
    if fn and len(fn) >= 10:
        sw(ws, "Y13", fn[0:4])
        sw(ws, "AB13", fn[5:7])
        sw(ws, "AD13", fn[8:10])
    sw(ws, "AE13", d.get("edad",""))
    sw(ws, "AG13", d.get("grupoSanguineo",""))
    sw(ws, "AO13", d.get("lateralidad",""))

    # B. Motivo de consulta
    sw(ws, "I16", d.get("puestoCIUO",""))
    sw(ws, "AG16", d.get("fechaAtencion",""))
    sw(ws, "A18", d.get("fechaIngreso",""))
    sw(ws, "AC18", d.get("fechaReintegro",""))
    sw(ws, "AL18", d.get("fechaUltimoDia",""))

    # Tipo evaluación
    tipo = d.get("tipoEvaluacion","")
    sw(ws, "F20", "X" if tipo == "INGRESO" else "")
    sw(ws, "Y20", "X" if tipo == "PERIÓDICO" else "")
    sw(ws, "AH20", "X" if tipo == "REINTEGRO" else "")
    sw(ws, "AZ20", "X" if tipo == "RETIRO" else "")
    sw(ws, "A22", d.get("motivoObs",""))

    # C. Antecedentes personales
    texto_ant = d.get("antClinico","")
    sw(ws, "A25", texto_ant)
    if texto_ant:
        lineas = max(2, len(texto_ant) // 120 + 1)
        ws.row_dimensions[25].height = max(30, lineas * 15)

    texto_fam = d.get("antFamiliares","")
    sw(ws, "A27", texto_fam)
    if texto_fam:
        lineas = max(2, len(texto_fam) // 120 + 1)
        ws.row_dimensions[27].height = max(30, lineas * 15)

    # Transfusiones y tratamiento hormonal
    transfusion = d.get("autTransfusion","")
    sw(ws, "N30", "X" if transfusion == "SI" else "")
    sw(ws, "R30", "X" if transfusion == "NO" else "")
    tto = d.get("ttoHormonal","")
    sw(ws, "AG30", "X" if tto == "SI" else "")
    sw(ws, "BD30", "X" if tto == "NO" else "")
    sw(ws, "AK30", d.get("ttoHormonalCual",""))

    # Gineco obstétricos
    fum = d.get("fum","")
    sw(ws, "A34", fum)
    sw(ws, "W34", d.get("gestas",""))
    sw(ws, "AA34", d.get("partos",""))
    sw(ws, "AD34", d.get("cesareas",""))
    sw(ws, "AG34", d.get("abortos",""))
    sw(ws, "AM34", d.get("metPlanFem",""))

    # Reproductivos masculinos
    sw(ws, "A41", d.get("examRepMasc",""))
    sw(ws, "N41", d.get("tiempoRepMasc",""))
    sw(ws, "AM41", d.get("metPlanMasc",""))

    # Consumo de sustancias
    sw(ws, "A45", "TABACO")
    sw(ws, "M45", d.get("tabacoT",""))
    sw(ws, "R45", chk(d.get("tabacoEx")))
    sw(ws, "W45", d.get("tabacoAbs",""))
    sw(ws, "AA45", chk(d.get("tabacoNo")))

    sw(ws, "A46", "ALCOHOL")
    sw(ws, "M46", d.get("alcoholT",""))
    sw(ws, "R46", chk(d.get("alcoholEx")))
    sw(ws, "W46", d.get("alcoholAbs",""))
    sw(ws, "AA46", chk(d.get("alcoholNo")))

    sw(ws, "A47", d.get("otraSustancia","OTRAS"))
    sw(ws, "M47", d.get("otraSustanciaT",""))
    sw(ws, "R47", chk(d.get("otraSustanciaEx")))
    sw(ws, "W47", d.get("otraSustanciaAbs",""))
    sw(ws, "AA47", chk(d.get("otraSustanciaNo")))

    sw(ws, "A48", d.get("condEspecial",""))

    # Estilo de vida - Actividad física
    sw(ws, "AD44", "ACTIVIDAD FÍSICA")
    sw(ws, "AH45", chk(d.get("actFisica")))
    sw(ws, "AK45", d.get("actFisicaT",""))
    sw(ws, "AT45", d.get("actFisicaCual",""))

    # Medicación habitual — AN44:AS47 es la celda merged del label
    # Datos van en: AT45=¿cuál?, BB45=cantidad
    sw(ws, "AN44", "MEDICACIÓN HABITUAL")
    if d.get("medicHabitual"):
        sw(ws, "AT45", d.get("medicHabitualCual",""))
        sw(ws, "BB45", d.get("medicHabitualCantidad",""))

    # No Aplica Gineco — si está marcado, escribir en celda FUM
    if d.get("ginecoNoAplica"):
        sw(ws, "A34", "NO APLICA")

    # D. Enfermedad actual
    sw(ws, "A51", d.get("enfermedadActual",""))

    # E. Constantes vitales
    sw(ws, "A56", d.get("temperatura",""))
    sw(ws, "H56", d.get("pa",""))
    sw(ws, "O56", d.get("fc",""))
    sw(ws, "V56", d.get("fr",""))
    sw(ws, "AB56", d.get("spo2",""))
    sw(ws, "AF56", d.get("peso",""))
    sw(ws, "AH56", d.get("talla",""))
    sw(ws, "AL56", d.get("imc",""))
    sw(ws, "AR56", d.get("perAbdominal",""))

    # F. Examen físico — construir texto desde regiones marcadas
    regiones = d.get("examenFisicoRegiones", {})
    obs_adicional = d.get("examenFisicoObs","")
    
    if regiones:
        lineas = []
        for region, data in regiones.items():
            if isinstance(data, dict):
                if data.get("checked") and data.get("hallazgo"):
                    lineas.append(f"{region}: {data['hallazgo']}")
                elif data.get("checked"):
                    lineas.append(f"{region}: PATOLOGÍA (pendiente descripción)")
                else:
                    lineas.append(f"{region}: NORMAL")
        texto_examen = "; ".join(lineas)
        if obs_adicional:
            texto_examen += f"\nObservaciones: {obs_adicional}"
    else:
        texto_examen = d.get("examenFisicoObs","")
    
    sw(ws, "A72", texto_examen)
    if texto_examen:
        chars_por_linea = 120
        lineas_count = max(2, len(texto_examen) // chars_por_linea + texto_examen.count('\n') + 1)
        altura = max(30, lineas_count * 15)
        ws.row_dimensions[72].height = altura
        ws.row_dimensions[73].height = altura

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def llenar_hoja2(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "SEGUNDA_HOJA.xlsx"))
    ws = wb.active

    # Puesto de trabajo
    sw(ws, "G2", d.get("puestoRiesgo",""))

    # Actividades
    actividades = d.get("actividades", ["","","","","","",""])
    act_cols = ["G5","I5","K5","M5","N5","O5","P5"]
    for i, col in enumerate(act_cols):
        if i < len(actividades):
            sw(ws, col, actividades[i])

    # Riesgos - columnas G..P para actividades 1-7
    RIESGOS_FILAS = {
        # FÍSICO
        "Temperaturas altas": 6,
        "Temperaturas bajas": 7,
        "Radiación Ionizante": 8,
        "Radiación No Ionizante": 9,
        "Ruido": 10,
        "Vibración": 11,
        "Iluminación": 12,
        "Ventilación": 13,
        "Fluido eléctrico": 14,
        "Otros (Físico)": 15,
        # DE SEGURIDAD
        "Falta de señalización, aseo, desorden": 16,
        "Atrapamiento entre Máquinas y o superficies": 17,
        "Atrapamiento entre objetos": 18,
        "Caída de objetos": 19,
        "Caídas al mismo nivel": 20,
        "Caídas a diferente nivel": 21,
        "Pinchazos": 22,
        "Cortes": 23,
        "Choques /colisión vehicular": 24,
        "Atropellamientos por vehículos": 25,
        "Proyección de fluidos": 26,
        "Proyección de partículas – fragmentos": 27,
        "Contacto con superficies de trabajos": 28,
        "Contacto eléctrico": 29,
        # QUÍMICO
        "Polvos ": 31,
        "Sólidos": 32,
        "Humos": 33,
        "líquidos ": 34,
        "vapores": 35,
        "Aerosoles": 36,
        "Neblinas ": 37,
        "Gaseosos": 38,
        # BIOLÓGICO
        "Virus ": 40,
        "Hongos": 41,
        "Bacterias ": 42,
        "Parásitos ": 43,
        "Exposición a vectores": 44,
        "Exposición a animales selváticos ": 45,
        # ERGONÓMICO
        "Manejo manual de cargas": 47,
        "Movimiento repetitivos": 48,
        "Posturas forzadas": 49,
        "Trabajos con PVD": 50,
        "Diseño Inadecuado del puesto": 51,
        # PSICOSOCIAL
        "Monotonía del trabajo": 53,
        "Sobrecarga laboral": 54,
        "Minuciosidad de la tarea ": 55,
        "Alta responsabilidad": 56,
        "Autonomía  en la toma de decisiones": 57,
        "Supervisión y estilos de dirección deficiente": 58,
        "Conflicto de rol": 59,
        "Falta de Claridad en las funciones": 60,
        "Incorrecta distribución del trabajo ": 61,
        "Turnos rotativos": 62,
        "Relaciones interpersonales ": 63,
        "inestabilidad laboral": 64,
        "Amenaza Delincuencial": 65,
    }
    # Columnas para actividades 1-7
    ACT_COLS = ["G","H","I","J","K","L","M","N","O","P"]  # G=act1, I=act2, K=act3, M=act4, N=act5, O=act6, P=act7
    ACT_LETTER = {0:"G",1:"I",2:"K",3:"M",4:"N",5:"O",6:"P"}

    riesgos_data = d.get("riesgos", {})
    for nombre_riesgo, fila in RIESGOS_FILAS.items():
        # Find best match in riesgos_data
        for key, vals in riesgos_data.items():
            if key.strip().lower() == nombre_riesgo.strip().lower() or key.strip() in nombre_riesgo or nombre_riesgo.strip() in key:
                for i, checked in enumerate(vals[:7]):
                    col = ACT_LETTER.get(i,"G")
                    if checked:
                        sw(ws, f"{col}{fila}", "X")
                break

    # Medidas preventivas — A67:F71 es la celda merged del label
    # El texto va en G67 (primera celda disponible a la derecha del label)
    sw(ws, "G67", d.get("medidasPreventivas",""))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def llenar_hoja3(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "TERCERA_HOJA.xlsx"))
    ws = wb.active

    # H. Antecedentes laborales (primera fila de datos)
    al = d.get("antLaborales", [{}])
    if al:
        a = al[0]
        sw(ws, "B7", a.get("centro",""))
        sw(ws, "J7", a.get("actividades",""))
        sw(ws, "W7", chk(a.get("anterior")))
        sw(ws, "Y7", chk(a.get("actual")))
        sw(ws, "AA7", a.get("tiempo",""))
        sw(ws, "AC7", chk(a.get("incidente")))
        sw(ws, "AE7", chk(a.get("accidente")))
        sw(ws, "AH7", chk(a.get("ep")))
        sw(ws, "AK7", chk(a.get("iesssi")))
        sw(ws, "AM7", chk(a.get("iessno")))
        sw(ws, "AO7", a.get("fecha",""))
        sw(ws, "AR7", a.get("especificar",""))
    # Segunda fila si existe
    if len(al) > 1:
        a = al[1]
        sw(ws, "B8", a.get("centro",""))
        sw(ws, "J8", a.get("actividades",""))
        sw(ws, "AA8", a.get("tiempo",""))
        sw(ws, "AR8", a.get("especificar",""))

    # I. Extra laborales
    sw(ws, "B29", d.get("actExtraLaborales",""))

    # J. Exámenes
    examenes = d.get("examenes", [])
    filas_exam = [36, 37, 38, 39, 40]
    for i, ex in enumerate(examenes[:5]):
        if i < len(filas_exam):
            f = filas_exam[i]
            sw(ws, f"B{f}", ex.get("nombre",""))
            sw(ws, f"M{f}", ex.get("fecha",""))
            sw(ws, f"T{f}", ex.get("resultado",""))
    sw(ws, "B41", "OBSERVACIONES: " + d.get("examenesObs",""))

    # K. Diagnósticos
    diagnosticos = d.get("diagnosticos", [])
    filas_dx = [45,46,47,48,49,50]
    for i, dx in enumerate(diagnosticos[:6]):
        if i < len(filas_dx):
            f = filas_dx[i]
            sw(ws, f"B{f}", dx.get("cie10",""))
            sw(ws, f"Q{f}", dx.get("desc",""))
            sw(ws, f"AO{f}", dx.get("tipo","PRE"))

    # L. Aptitud — checkboxes en P53, AE53, AS53, BB53 (celdas junto a los labels)
    aptitud = d.get("aptitud","")
    sw(ws, "P53",  "X" if aptitud == "APTO" else "")
    sw(ws, "AE53", "X" if aptitud == "APTO_OBS" else "")
    sw(ws, "AS53", "X" if aptitud == "APTO_LIM" else "")
    sw(ws, "BB53", "X" if aptitud == "NO_APTO" else "")
    # Observaciones van en B57 (fila debajo del label OBSERVACIONES: en B56)
    sw(ws, "B57", d.get("aptitudObs",""))

    # M. Recomendaciones
    sw(ws, "B59", d.get("recomendaciones",""))

    # N. Retiro — checkboxes en AG65/AW65 y AG66/AW66 (celdas junto a SI/NO)
    ret_eval = d.get("retiroEval","")
    sw(ws, "AG65", "X" if ret_eval == "SI" else "")
    sw(ws, "AW65", "X" if ret_eval == "NO" else "")
    ret_rel = d.get("retiroRelacionado","")
    sw(ws, "AG66", "X" if ret_rel == "SI" else "")
    sw(ws, "AW66", "X" if ret_rel == "NO" else "")
    # Observación retiro — B68 (fila debajo del label Observación: en B67)
    sw(ws, "B68", d.get("retiroObs",""))

    # O. Datos profesional
    sw(ws, "B73", d.get("nombreProf",""))
    sw(ws, "U73", d.get("codigoMedico",""))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def llenar_certificado(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "CERTIFICADO_DE_APTITUP_LABORAL.xlsx"))
    ws = wb.active

    # A. Datos establecimiento
    sw(ws, "A4", d.get("institucion",""))
    sw(ws, "L4", d.get("ruc",""))
    sw(ws, "R4", d.get("ciiu",""))
    sw(ws, "V4", d.get("establecimiento",""))
    sw(ws, "AC4", d.get("nroHistoria",""))
    sw(ws, "AH4", d.get("nroArchivo",""))

    # A. Nombres
    sw(ws, "A6", d.get("primerApellido",""))
    sw(ws, "J6", d.get("segundoApellido",""))
    sw(ws, "Q6", d.get("primerNombre",""))
    sw(ws, "X6", d.get("segundoNombre",""))
    sw(ws, "AD6", d.get("sexo",""))
    sw(ws, "AG6", d.get("puestoCIUO",""))

    # B. Datos generales — fecha en K10/M10/O10 (K11/M11/O11 son labels aaaa/mm/dd)
    fa = d.get("fechaAtencion","")
    if fa and len(fa) >= 10:
        sw(ws, "K10", fa[0:4])   # año
        sw(ws, "M10", fa[5:7])   # mes
        sw(ws, "O10", fa[8:10])  # día

    # Tipo evaluación — checkboxes en L12/U12/AC12/AI12 (I12/P12/W12/AF12 son labels)
    tipo = d.get("tipoEvaluacion","")
    sw(ws, "L12",  "X" if tipo == "INGRESO" else "")
    sw(ws, "U12",  "X" if tipo == "PERIÓDICO" else "")
    sw(ws, "AC12", "X" if tipo == "REINTEGRO" else "")
    sw(ws, "AI12", "X" if tipo == "RETIRO" else "")

    # C. Aptitud — checkboxes en I17/S17/AC17/AK17 (A17/J17/U17/AE17 son labels)
    aptitud = d.get("aptitud","")
    sw(ws, "I17",  "X" if aptitud == "APTO" else "")
    sw(ws, "S17",  "X" if aptitud == "APTO_OBS" else "")
    sw(ws, "AC17", "X" if aptitud == "APTO_LIM" else "")
    sw(ws, "AK17", "X" if aptitud == "NO_APTO" else "")
    # Observaciones en A19 (A18 es label "DETALLE DE OBSERVACIONES:")
    sw(ws, "A19", d.get("aptitudObs",""))

    # D. Recomendaciones
    sw(ws, "A24", d.get("recomendaciones",""))
    sw(ws, "A25", "Observación: " + d.get("aptitudObs",""))

    # E. Datos profesional
    sw(ws, "A33", d.get("nombreProf",""))
    sw(ws, "L33", d.get("codigoMedico",""))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
#  MAPA DE COLORES INDEXED → RGB (paleta Excel)
# ─────────────────────────────────────────────
INDEXED_TO_RGB = {
    9:  'FFFFFFFF',  # Blanco
    27: 'FFCCFFFF',  # Celeste claro
    31: 'FFCCCCFF',  # Lavanda/azul claro
    42: 'FFCCFFCC',  # Verde claro
    0:  'FF000000', 1:  'FFFFFFFF', 2:  'FFFF0000', 3:  'FF00FF00',
    4:  'FF0000FF', 5:  'FFFFFF00', 6:  'FFFF00FF', 7:  'FF00FFFF',
    22: 'FFC0C0C0', 26: 'FFFFFFCC', 43: 'FFFFFF99', 44: 'FF99CCFF',
    45: 'FFFF99CC', 47: 'FFFFCC99', 49: 'FF33CCCC', 51: 'FFFFCC00',
    53: 'FFFF6600', 62: 'FF333399', 63: 'FF000000', 64: 'FFFFFFFF',
}

def get_fill_rgb(fill):
    """Obtiene el color RGB de un fill, convirtiendo indexed si es necesario."""
    if not fill or fill.fill_type != 'solid':
        return None
    fg = fill.fgColor
    if fg.type == 'rgb' and fg.rgb not in ('00000000', 'FF000000'):
        return fg.rgb
    if fg.type == 'indexed':
        return INDEXED_TO_RGB.get(fg.indexed)
    return None

def normalize_border_side(side):
    """Convierte un Side con color theme/indexed a RGB explícito (negro)."""
    if not side or not side.style:
        return side
    try:
        ct = side.color.type if side.color else 'rgb'
        if ct in ('theme', 'indexed'):
            color = 'FF000000'
        elif ct == 'rgb':
            color = side.color.rgb if side.color.rgb != '00000000' else 'FF000000'
        else:
            color = 'FF000000'
    except:
        color = 'FF000000'
    return Side(style=side.style, color=color)

def normalize_border(border):
    """Normaliza todos los lados de un borde a colores RGB explícitos."""
    if not border:
        return border
    return Border(
        left=normalize_border_side(border.left),
        right=normalize_border_side(border.right),
        top=normalize_border_side(border.top),
        bottom=normalize_border_side(border.bottom),
        diagonal=border.diagonal,
        diagonal_direction=border.diagonal_direction,
    )


# ─────────────────────────────────────────────
#  HELPER: COPIAR HOJA ENTRE WORKBOOKS
# ─────────────────────────────────────────────
def copiar_hoja(ws_origen, wb_destino, nombre_hoja):
    """Copia una hoja completa (celdas, formato, merges, dimensiones) a otro workbook."""
    ws_dest = wb_destino.create_sheet(nombre_hoja)

    # Dimensiones de columnas y filas
    for col, cd in ws_origen.column_dimensions.items():
        ws_dest.column_dimensions[col].width = cd.width
        ws_dest.column_dimensions[col].hidden = cd.hidden
    for row, rd in ws_origen.row_dimensions.items():
        ws_dest.row_dimensions[row].height = rd.height
        ws_dest.row_dimensions[row].hidden = rd.hidden

    # Celdas combinadas (merged) — aplicar PRIMERO
    for rng in ws_origen.merged_cells.ranges:
        ws_dest.merge_cells(str(rng))

    # Copiar celdas con valor y estilo
    # IMPORTANTE: las celdas secundarias de rangos merged (MergedCell)
    # también tienen bordes — deben copiarse para que el formato sea correcto
    for row in ws_origen.iter_rows():
        for cell in row:
            dest_cell = ws_dest.cell(row=cell.row, column=cell.column)

            # Valor: solo funciona en celdas normales, no en MergedCell secundarias
            try:
                dest_cell.value = cell.value
            except AttributeError:
                pass  # MergedCell secundaria — sin valor

            # Estilos: intentar copiar en todas las celdas (incluidas MergedCell)
            if cell.has_style:
                # Borde: fundamental para celdas secundarias de merged ranges
                try:
                    dest_cell.border = normalize_border(cell.border)
                except Exception:
                    pass
                # Resto de estilos solo para celdas normales
                try:
                    dest_cell.font      = copy(cell.font)
                    dest_cell.alignment = copy(cell.alignment)
                    dest_cell.number_format = cell.number_format
                    rgb = get_fill_rgb(cell.fill)
                    if rgb:
                        dest_cell.fill = PatternFill(fill_type='solid', fgColor=rgb)
                    else:
                        dest_cell.fill = copy(cell.fill)
                except AttributeError:
                    pass  # MergedCell secundaria — solo border importa
    return ws_dest


def generar_ficha_completa(d):
    """Genera ZIP con las 3 hojas separadas — funciona en plan gratuito."""
    import gc, zipfile
    nombre = f"{d.get('primerApellido','SN')}_{d.get('primerNombre','SN')}".replace(" ","_")
    buf_zip = io.BytesIO()
    with zipfile.ZipFile(buf_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        b1 = llenar_hoja1(d); zf.writestr(f"FichaOcupacional_1-3_{nombre}.xlsx", b1.read()); del b1; gc.collect()
        b2 = llenar_hoja2(d); zf.writestr(f"FichaOcupacional_2-3_{nombre}.xlsx", b2.read()); del b2; gc.collect()
        b3 = llenar_hoja3(d); zf.writestr(f"FichaOcupacional_3-3_{nombre}.xlsx", b3.read()); del b3; gc.collect()
    buf_zip.seek(0)
    return buf_zip


# ─────────────────────────────────────────────
#  RUTAS API
# ─────────────────────────────────────────────
@app.route("/api/descargar-ficha", methods=["POST"])
def api_ficha():
    d = request.json
    nombre = f"{d.get('primerApellido','SN')}_{d.get('primerNombre','SN')}".replace(" ","_")
    fecha  = d.get("fechaAtencion","SFecha")
    buf = generar_ficha_completa(d)
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"FichaOcupacional_{nombre}_{fecha}.zip"
    )

@app.route("/api/descargar-certificado", methods=["POST"])
def api_cert():
    d = request.json
    nombre = f"{d.get('primerApellido','SN')}_{d.get('primerNombre','SN')}".replace(" ","_")
    fecha  = d.get("fechaAtencion","SFecha")
    buf = llenar_certificado(d)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Certificado_Aptitud_{nombre}_{fecha}.xlsx"
    )

@app.route("/api/agregar-matriz", methods=["POST"])
def api_agregar():
    d = request.json
    nombre = f"{d.get('primerApellido','')} {d.get('segundoApellido','')} {d.get('primerNombre','')} {d.get('segundoNombre','')}".strip()
    APTITUD_TEXTO = {"APTO":"APTO","APTO_OBS":"APTO EN OBSERVACIÓN","APTO_LIM":"APTO CON LIMITACIONES","NO_APTO":"NO APTO"}
    fila = {
        "N°": db_contar() + 1,
        "NOMBRE COMPLETO": nombre,
        "N° HISTORIA CLÍNICA": d.get("nroHistoria",""),
        "EMPRESA/INSTITUCIÓN": d.get("institucion",""),
        "RUC": d.get("ruc",""),
        "ESTABLECIMIENTO": d.get("establecimiento",""),
        "PUESTO/CARGO (CIUO)": d.get("puestoCIUO",""),
        "SEXO": d.get("sexo",""),
        "FECHA NACIMIENTO": d.get("fechaNacimiento",""),
        "EDAD (años)": d.get("edad",""),
        "GRUPO SANGUÍNEO": d.get("grupoSanguineo",""),
        "GRUPO AT. PRIORITARIA": ", ".join(filter(None,[
            "Embarazada" if d.get("embarazada") else "",
            "Discapacidad" if d.get("discapacidad") else "",
            "Enf.Catastrófica" if d.get("enfCatastrofica") else "",
            "Adulto Mayor" if d.get("adultoMayor") else "",
        ])) or "Ninguno",
        "TIPO EVALUACIÓN": d.get("tipoEvaluacion",""),
        "FECHA ATENCIÓN": d.get("fechaAtencion",""),
        "FECHA INGRESO": d.get("fechaIngreso",""),
        "TEMPERATURA (°C)": d.get("temperatura",""),
        "PRESIÓN ARTERIAL": d.get("pa",""),
        "FC (lat/min)": d.get("fc",""),
        "FR (fr/min)": d.get("fr",""),
        "SpO2 (%)": d.get("spo2",""),
        "PESO (kg)": d.get("peso",""),
        "TALLA (cm)": d.get("talla",""),
        "IMC (kg/m²)": d.get("imc",""),
        "PERÍMETRO ABDOMINAL (cm)": d.get("perAbdominal",""),
        "APTITUD": APTITUD_TEXTO.get(d.get("aptitud",""),""),
        "OBSERVACIONES APTITUD": d.get("aptitudObs",""),
        "DIAGNÓSTICO 1 CIE-10": (d.get("diagnosticos",[{}])[0] or {}).get("cie10",""),
        "DESCRIPCIÓN DX1": (d.get("diagnosticos",[{}])[0] or {}).get("desc",""),
        "DIAGNÓSTICO 2 CIE-10": (d.get("diagnosticos",[{},{}])[1] if len(d.get("diagnosticos",[])) > 1 else {}).get("cie10",""),
        "DESCRIPCIÓN DX2": (d.get("diagnosticos",[{},{}])[1] if len(d.get("diagnosticos",[])) > 1 else {}).get("desc",""),
        "RECOMENDACIONES": d.get("recomendaciones",""),
        "MÉDICO RESPONSABLE": d.get("nombreProf",""),
        "CÓDIGO MÉDICO": d.get("codigoMedico",""),
    }
    try:
        db_insertar(fila)
        total = db_contar()
        return jsonify({"ok": True, "total": total, "nombre": nombre})
    except Exception as e:
        # Si Supabase falla, guardar en memoria local como respaldo
        _matriz_local.append(fila)
        total = len(_matriz_local)
        print(f"⚠ Supabase error, guardado local: {e}")
        return jsonify({"ok": True, "total": total, "nombre": nombre})

@app.route("/api/descargar-matriz", methods=["GET"])
def api_matriz():
    try:
        filas = db_listar()
    except Exception as e:
        filas = []
    # Si Supabase está vacío pero hay datos locales, usar locales
    if not filas and _matriz_local:
        filas = _matriz_local
    if not filas:
        return jsonify({"error": "Matriz vacía"}), 400
    wb = Workbook()
    ws = wb.active
    ws.title = "MATRIZ DE SEGUIMIENTO"
    headers = list(filas[0].keys())
    fill_hdr = PatternFill("solid", fgColor="FF1556A0")
    font_hdr = Font(bold=True, color="FFFFFFFF", size=9)
    thin = Side(style='thin', color='FF000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_i, value=h)
        cell.fill = fill_hdr
        cell.font = font_hdr
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = border
    fill_even = PatternFill("solid", fgColor="FFE8EEF8")
    font_data = Font(size=9)
    for row_i, fila in enumerate(filas, 2):
        fill = fill_even if row_i % 2 == 0 else PatternFill("solid", fgColor="FFFFFFFF")
        for col_i, h in enumerate(headers, 1):
            cell = ws.cell(row=row_i, column=col_i, value=fila.get(h,""))
            cell.fill = fill
            cell.font = font_data
            cell.alignment = Alignment(wrap_text=True)
            cell.border = border
    widths = {"N°":5,"NOMBRE COMPLETO":30,"N° HISTORIA CLÍNICA":15,"EMPRESA/INSTITUCIÓN":25,"RUC":14,
              "ESTABLECIMIENTO":22,"PUESTO/CARGO (CIUO)":22,"SEXO":8,"FECHA NACIMIENTO":14,"EDAD (años)":8,
              "GRUPO SANGUÍNEO":10,"GRUPO AT. PRIORITARIA":20,"TIPO EVALUACIÓN":14,"FECHA ATENCIÓN":14,
              "FECHA INGRESO":14,"TEMPERATURA (°C)":10,"PRESIÓN ARTERIAL":10,"FC (lat/min)":8,"FR (fr/min)":8,
              "SpO2 (%)":8,"PESO (kg)":8,"TALLA (cm)":8,"IMC (kg/m²)":10,"PERÍMETRO ABDOMINAL (cm)":14,
              "APTITUD":22,"OBSERVACIONES APTITUD":30,"DIAGNÓSTICO 1 CIE-10":12,"DESCRIPCIÓN DX1":30,
              "DIAGNÓSTICO 2 CIE-10":12,"DESCRIPCIÓN DX2":30,"RECOMENDACIONES":40,
              "MÉDICO RESPONSABLE":25,"CÓDIGO MÉDICO":12}
    from openpyxl.utils import get_column_letter
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 15)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=f"Matriz_Fichas_Medicas_{date.today().isoformat()}.xlsx")

@app.route("/api/limpiar-matriz", methods=["POST"])
def api_limpiar():
    n = db_contar()
    db_limpiar()
    return jsonify({"ok": True, "eliminados": n})

@app.route("/api/conteo-matriz")
def api_conteo():
    filas = db_listar()
    return jsonify({"total": len(filas),
                    "nombres": [f"{m.get('N°',i+1)}. {m.get('NOMBRE COMPLETO','')}" for i,m in enumerate(filas)]})

@app.route("/")
def index():
    html_path = os.path.join(BASE, "index.html")
    with open(html_path, encoding="utf-8") as f:
        return f.read()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print("=" * 55)
    print("  FICHA MÉDICA OCUPACIONAL – CASISO S.A.S.")
    print(f"  Servidor en puerto {port}")
    print("=" * 55)
    app.run(host="0.0.0.0", port=port, debug=debug)
