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
    wb = load_workbook(os.path.join(PLANTILLAS, "Formulario_evaluaciones_ocupacionales.xlsx"))
    ws = wb['-EVALUACION OCUPACIONAL 1-3']

    # A. Datos establecimiento (fila 4)
    sw(ws, "B4",  d.get("institucion",""))
    sw(ws, "Q4",  d.get("ruc",""))
    sw(ws, "Z4",  d.get("ciiu",""))
    sw(ws, "AC4", d.get("establecimiento",""))
    sw(ws, "AP4", d.get("nroHistoria",""))
    sw(ws, "BB4", d.get("nroArchivo",""))

    # A. Nombres (fila 6)
    sw(ws, "B6",  d.get("primerApellido",""))
    sw(ws, "S6",  d.get("segundoApellido",""))
    sw(ws, "AH6", d.get("primerNombre",""))
    sw(ws, "AR6", d.get("segundoNombre",""))

    # Grupo atención prioritaria (fila 10)
    sw(ws, "B10", chk(d.get("embarazada")))
    sw(ws, "G10", chk(d.get("discapacidad")))
    sw(ws, "K10", chk(d.get("enfCatastrofica")))
    sw(ws, "N10", chk(d.get("adultoMayor")))
    sexo = d.get("sexo","")
    sw(ws, "U11", "X" if sexo == "Hombre" else "")
    sw(ws, "X11", "X" if sexo == "Mujer" else "")

    # Fecha nacimiento fila 12
    fn = d.get("fechaNacimiento","")
    if fn and len(fn) >= 10:
        sw(ws, "Z12",  fn[0:4])
        sw(ws, "AC12", fn[5:7])
        sw(ws, "AE12", fn[8:10])
    sw(ws, "AF9",  d.get("edad",""))
    sw(ws, "AH9",  d.get("grupoSanguineo",""))
    sw(ws, "AP9",  d.get("lateralidad",""))

    # B. Motivo de consulta fila 16-20
    sw(ws, "J16",  d.get("puestoCIUO",""))
    sw(ws, "Y16",  d.get("fechaAtencion",""))
    sw(ws, "B17",  d.get("fechaIngreso",""))
    sw(ws, "AD17", d.get("fechaReintegro",""))
    sw(ws, "AM17", d.get("fechaUltimoDia",""))

    tipo = d.get("tipoEvaluacion","")
    sw(ws, "B20",  "X" if tipo == "INGRESO"   else "")
    sw(ws, "O20",  "X" if tipo == "PERIÓDICO" else "")
    sw(ws, "AD20", "X" if tipo == "REINTEGRO" else "")
    sw(ws, "AM20", "X" if tipo == "RETIRO"    else "")
    sw(ws, "B21",  d.get("motivoObs",""))

    # C. Antecedentes personales fila 25/27
    sw(ws, "B25", d.get("antClinico",""))
    sw(ws, "B27", d.get("antFamiliares",""))

    # Transfusiones / tto hormonal fila 29
    transfusion = d.get("autTransfusion","")
    sw(ws, "O29",  "X" if transfusion == "SI" else "")
    sw(ws, "S29",  "X" if transfusion == "NO" else "")
    tto = d.get("ttoHormonal","")
    sw(ws, "AH29", "X" if tto == "SI" else "")
    sw(ws, "AL29", d.get("ttoHormonalCual",""))
    sw(ws, "BE29", "X" if tto == "NO" else "")

    # Gineco obstétricos fila 32-33
    if d.get("ginecoNoAplica"):
        sw(ws, "B32", "NO APLICA")
    else:
        sw(ws, "B32",  d.get("fum",""))
        sw(ws, "X32",  d.get("gestas",""))
        sw(ws, "AB32", d.get("partos",""))
        sw(ws, "AE32", d.get("cesareas",""))
        sw(ws, "AH32", d.get("abortos",""))
        sw(ws, "AL32", d.get("metPlanFem",""))

    # Reproductivos masculinos fila 39-40
    sw(ws, "B39",  d.get("examRepMasc",""))
    sw(ws, "O39",  d.get("tiempoRepMasc",""))
    sw(ws, "AE39", d.get("metPlanMasc",""))

    # Consumo sustancias filas 45-48
    sw(ws, "N45",  d.get("tabacoT",""))
    sw(ws, "S45",  chk(d.get("tabacoEx")))
    sw(ws, "X45",  d.get("tabacoAbs",""))
    sw(ws, "AB45", chk(d.get("tabacoNo")))

    sw(ws, "N46",  d.get("alcoholT",""))
    sw(ws, "S46",  chk(d.get("alcoholEx")))
    sw(ws, "X46",  d.get("alcoholAbs",""))
    sw(ws, "AB46", chk(d.get("alcoholNo")))

    sw(ws, "N47",  d.get("otraSustanciaT",""))
    sw(ws, "S47",  chk(d.get("otraSustanciaEx")))
    sw(ws, "X47",  d.get("otraSustanciaAbs",""))
    sw(ws, "AB47", chk(d.get("otraSustanciaNo")))
    sw(ws, "B48",  d.get("condEspecial",""))

    # Estilo de vida fila 44-45
    sw(ws, "AI44", d.get("actFisicaCual",""))
    sw(ws, "AL44", d.get("actFisicaT",""))

    # Medicación habitual fila 44-47
    if d.get("medicHabitual"):
        sw(ws, "AU44", d.get("medicHabitualCual",""))
        sw(ws, "BC44", d.get("medicHabitualCantidad",""))

    # D. Enfermedad actual fila 51
    sw(ws, "B51", d.get("enfermedadActual",""))

    # E. Constantes vitales fila 55-56
    sw(ws, "B55",  d.get("temperatura",""))
    sw(ws, "I55",  d.get("pa",""))
    sw(ws, "P55",  d.get("fc",""))
    sw(ws, "W55",  d.get("fr",""))
    sw(ws, "AC55", d.get("spo2",""))
    sw(ws, "AG55", d.get("peso",""))
    sw(ws, "AI55", d.get("talla",""))
    sw(ws, "AM55", d.get("imc",""))
    sw(ws, "AS55", d.get("perAbdominal",""))

    # F. Examen físico fila 75
    regiones  = d.get("examenFisicoRegiones", {})
    obs_adicional = d.get("examenFisicoObs","")
    if regiones:
        con_patologia, normales = [], []
        for region, data in regiones.items():
            if isinstance(data, dict):
                if data.get("checked"):
                    h = data.get("hallazgo","").strip()
                    con_patologia.append(f"{region}: {h}" if h else f"{region}: PATOLOGÍA")
                else:
                    normales.append(region.split(".")[0].strip())
        lineas = []
        if con_patologia: lineas.append("CON HALLAZGOS: " + "; ".join(con_patologia))
        if normales:       lineas.append("NORMAL: regiones " + ", ".join(normales))
        if obs_adicional:  lineas.append("Obs: " + obs_adicional)
        texto_examen = "\n".join(lineas) if lineas else "Sin hallazgos patológicos."
    else:
        texto_examen = obs_adicional

    sw(ws, "B75", texto_examen)

    print_setup(ws, "landscape")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def llenar_hoja2(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "Formulario_evaluaciones_ocupacionales.xlsx"))
    ws = wb['EVALUACION OCUPACIONAL2-3']

    # Puesto de trabajo A3:F5
    sw(ws, "A2", d.get("puestoRiesgo",""))

    # Actividades columnas G-P fila 3-4
    actividades = d.get("actividades", ["","","","","","",""])
    act_cols = ["G3","I3","K3","M3","N3","O3","P3"]
    for i, col in enumerate(act_cols[:len(actividades)]):
        sw(ws, col, actividades[i])

    # Riesgos — mismo dict que antes, mismas filas
    RIESGOS_FILAS = {
        "Temperaturas altas": 6, "Temperaturas bajas": 7,
        "Radiación Ionizante": 8, "Radiación No Ionizante": 9,
        "Ruido": 10, "Vibración": 11, "Iluminación": 12,
        "Ventilación": 13, "Fluido eléctrico": 14, "Otros (Físico)": 15,
        "Falta de señalización, aseo, desorden": 16,
        "Atrapamiento entre Máquinas y o superficies": 17,
        "Atrapamiento entre objetos": 18, "Caída de objetos": 19,
        "Caídas al mismo nivel": 20, "Caídas a diferente nivel": 21,
        "Pinchazos": 22, "Cortes": 23, "Choques /colisión vehicular": 24,
        "Atropellamientos por vehículos": 25, "Proyección de fluidos": 26,
        "Proyección de partículas – fragmentos": 27,
        "Contacto con superficies de trabajos": 28, "Contacto eléctrico": 29,
        "Polvos ": 31, "Sólidos": 32, "Humos": 33, "líquidos ": 34,
        "vapores": 35, "Aerosoles": 36, "Neblinas ": 37, "Gaseosos": 38,
        "Virus ": 40, "Hongos": 41, "Bacterias ": 42, "Parásitos ": 43,
        "Exposición a vectores": 44, "Exposición a animales selváticos ": 45,
        "Manejo manual de cargas": 47, "Movimiento repetitivos": 48,
        "Posturas forzadas": 49, "Trabajos con PVD": 50,
        "Diseño Inadecuado del puesto": 51,
        "Monotonía del trabajo": 53, "Sobrecarga laboral": 54,
        "Minuciosidad de la tarea ": 55, "Alta responsabilidad": 56,
        "Autonomía  en la toma de decisiones": 57,
        "Supervisión y estilos de dirección deficiente": 58,
        "Conflicto de rol": 59, "Falta de Claridad en las funciones": 60,
        "Incorrecta distribución del trabajo ": 61,
        "Turnos rotativos": 62, "Relaciones interpersonales ": 63,
        "inestabilidad laboral": 64, "Amenaza Delincuencial": 65,
    }
    ACT_LETTER = {0:"G",1:"I",2:"K",3:"M",4:"N",5:"O",6:"P"}
    riesgos_data = d.get("riesgos", {})
    for nombre_riesgo, fila in RIESGOS_FILAS.items():
        for key, vals in riesgos_data.items():
            if key.strip().lower() == nombre_riesgo.strip().lower() or key.strip() in nombre_riesgo or nombre_riesgo.strip() in key:
                for i, checked in enumerate(vals[:7]):
                    if checked:
                        sw(ws, f"{ACT_LETTER.get(i,'G')}{fila}", "X")
                break

    # Medidas preventivas — G67
    sw(ws, "G67", d.get("medidasPreventivas",""))

    print_setup(ws, "landscape")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def llenar_hoja3(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "Formulario_evaluaciones_ocupacionales.xlsx"))
    ws = wb['EVALUACION OCUPACIONAL 3-3']

    # H. Antecedentes laborales (filas 7-25)
    al = d.get("antLaborales", [{}])
    filas_al = [7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25]
    for i, a in enumerate(al[:len(filas_al)]):
        f = filas_al[i]
        sw(ws, f"B{f}",  a.get("centro",""))
        sw(ws, f"J{f}",  a.get("actividades",""))
        sw(ws, f"W{f}",  chk(a.get("anterior")))
        sw(ws, f"Y{f}",  chk(a.get("actual")))
        sw(ws, f"AA{f}", a.get("tiempo",""))
        sw(ws, f"AC{f}", chk(a.get("incidente")))
        sw(ws, f"AE{f}", chk(a.get("accidente")))
        sw(ws, f"AH{f}", chk(a.get("ep")))
        sw(ws, f"AK{f}", chk(a.get("iesssi")))
        sw(ws, f"AM{f}", chk(a.get("iessno")))
        sw(ws, f"AO{f}", a.get("fecha",""))
        sw(ws, f"AR{f}", a.get("especificar",""))

    # I. Extra laborales fila 28-29
    sw(ws, "A28", d.get("actExtraLaborales",""))

    # J. Exámenes filas 35-40
    examenes = d.get("examenes", [])
    for i, ex in enumerate(examenes[:6]):
        f = 35 + i
        sw(ws, f"B{f}",  ex.get("nombre",""))
        sw(ws, f"M{f}",  ex.get("fecha",""))
        sw(ws, f"T{f}",  ex.get("resultado",""))
    sw(ws, "B41", "OBSERVACIONES: " + d.get("examenesObs",""))

    # K. Diagnósticos filas 45-50
    diagnosticos = d.get("diagnosticos", [])
    for i, dx in enumerate(diagnosticos[:6]):
        f = 45 + i
        sw(ws, f"B{f}",  dx.get("cie10",""))
        sw(ws, f"Q{f}",  dx.get("desc",""))
        pre = dx.get("tipo","") == "PRE"
        def_ = dx.get("tipo","") == "DEF"
        sw(ws, f"BE{f}", "X" if pre  else "")
        sw(ws, f"BJ{f}", "X" if def_ else "")

    # L. Aptitud filas 53-54
    aptitud = d.get("aptitud","")
    sw(ws, "B53",  "X" if aptitud == "APTO"     else "")
    sw(ws, "R53",  "X" if aptitud == "APTO_OBS" else "")
    sw(ws, "AG53", "X" if aptitud == "APTO_LIM" else "")
    sw(ws, "AU53", "X" if aptitud == "NO_APTO"  else "")
    sw(ws, "B56",  d.get("aptitudObs",""))

    # M. Recomendaciones fila 59
    sw(ws, "B59", d.get("recomendaciones",""))

    # N. Retiro filas 65-67
    ret_eval = d.get("retiroEval","")
    sw(ws, "W65",  "X" if ret_eval == "SI" else "")
    sw(ws, "AI65", "X" if ret_eval == "NO" else "")
    ret_rel = d.get("retiroRelacionado","")
    sw(ws, "W66",  "X" if ret_rel == "SI" else "")
    sw(ws, "AI66", "X" if ret_rel == "NO" else "")
    sw(ws, "B67",  d.get("retiroObs",""))

    # O. Datos profesional filas 73-74
    sw(ws, "K73",  d.get("nombreProf",""))
    sw(ws, "AB73", d.get("codigoMedico",""))

    print_setup(ws, "landscape")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def llenar_certificado(d):
    wb = load_workbook(os.path.join(PLANTILLAS, "Certificado-de-Evaluacion-Medica-Ocupacional.xlsx"))
    ws = wb['CERTIFICADO']

    # A. Datos establecimiento fila 4
    sw(ws, "A4",  d.get("institucion",""))
    sw(ws, "L4",  d.get("ruc",""))
    sw(ws, "R4",  d.get("ciiu",""))
    sw(ws, "V4",  d.get("establecimiento",""))
    sw(ws, "AC4", d.get("nroHistoria",""))
    sw(ws, "AI4", d.get("nroArchivo",""))

    # Nombres fila 6
    sw(ws, "A6",  d.get("primerApellido",""))
    sw(ws, "J6",  d.get("segundoApellido",""))
    sw(ws, "Q6",  d.get("primerNombre",""))
    sw(ws, "X6",  d.get("segundoNombre",""))
    sw(ws, "AD6", d.get("sexo",""))
    sw(ws, "AG6", d.get("puestoCIUO",""))

    # B. Fecha emisión fila 10 (K10=año, M10=mes, O10=día)
    fa = d.get("fechaAtencion","")
    if fa and len(fa) >= 10:
        sw(ws, "K10", fa[0:4])
        sw(ws, "M10", fa[5:7])
        sw(ws, "O10", fa[8:10])

    # Tipo evaluación fila 12
    tipo = d.get("tipoEvaluacion","")
    sw(ws, "L12",  "X" if tipo == "INGRESO"   else "")
    sw(ws, "U12",  "X" if tipo == "PERIÓDICO" else "")
    sw(ws, "AC12", "X" if tipo == "REINTEGRO" else "")
    sw(ws, "AI12", "X" if tipo == "RETIRO"    else "")

    # C. Aptitud fila 17 — celdas maestras exactas
    aptitud = d.get("aptitud","")
    sw(ws, "A17",  "X" if aptitud == "APTO"     else "")
    sw(ws, "J17",  "X" if aptitud == "APTO_OBS" else "")
    sw(ws, "U17",  "X" if aptitud == "APTO_LIM" else "")
    sw(ws, "AE17", "X" if aptitud == "NO_APTO"  else "")
    sw(ws, "A19",  d.get("aptitudObs",""))

    # D. Recomendaciones fila 24-25
    sw(ws, "A24", d.get("recomendaciones",""))
    sw(ws, "A25", d.get("aptitudObs",""))

    # E. Datos profesional fila 33
    sw(ws, "A33", d.get("nombreProf",""))
    sw(ws, "L33", d.get("codigoMedico",""))

    print_setup(ws, "portrait")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

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


def print_setup(ws, orientacion="portrait"):
    """Configura impresión A4, ajuste al ancho de la página, márgenes mínimos."""
    from openpyxl.worksheet.page import PageMargins
    from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
    # Inicializar sheet_properties si no existe (algunas plantillas lo omiten)
    if ws.sheet_properties is None:
        ws.sheet_properties = WorksheetProperties()
    if ws.sheet_properties.pageSetUpPr is None:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.orientation = orientacion
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.scale       = 100
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered   = False
    ws.page_margins = PageMargins(
        left=0.2, right=0.2, top=0.3, bottom=0.3,
        header=0.0, footer=0.0
    )


def generar_ficha_completa(d):
    """Genera Excel con las 3 hojas de evaluación."""
    import gc
    wb_final = Workbook()
    wb_final.remove(wb_final.active)

    for fn, nombre_hoja in [
        (llenar_hoja1, "EVALUACION OCUPACIONAL 1-3"),
        (llenar_hoja2, "EVALUACION OCUPACIONAL 2-3"),
        (llenar_hoja3, "EVALUACION OCUPACIONAL 3-3"),
    ]:
        buf = fn(d)
        wb_src = load_workbook(io.BytesIO(buf.read()))
        copiar_hoja(wb_src.active, wb_final, nombre_hoja)
        wb_src.close()
        del wb_src, buf
        gc.collect()

    out = io.BytesIO()
    wb_final.save(out)
    out.seek(0)
    return out


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
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"FichaOcupacional_{nombre}_{fecha}.xlsx"
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

@app.route("/api/descargar-historia", methods=["POST"])
def api_historia():
    d = request.json
    nombre = f"{d.get('primerApellido','SN')}_{d.get('primerNombre','SN')}".replace(" ","_")
    fecha  = d.get("fechaAtencion","SFecha")
    buf = generar_ficha_completa(d)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"FichaOcupacional_{nombre}_{fecha}.xlsx"
    )

@app.route("/api/descargar-consentimiento", methods=["POST"])
def api_consentimiento():
    from docx import Document
    from docx.shared import Pt
    d = request.json or {}
    nombre = f"{d.get('primerNombre','')} {d.get('primerApellido','')}".strip() or "Paciente"
    cedula = d.get("cedula","") or d.get("nroHistoria","")
    empresa = d.get("institucion","")
    fecha = d.get("fechaAtencion","")
    plantilla = os.path.join(PLANTILLAS, "CONSENTIMIENTO_INFORMADO.docx")
    if not os.path.exists(plantilla):
        return jsonify({"ok": False, "error": "Plantilla no encontrada"}), 404
    doc = Document(plantilla)
    reemplazos = {
        "{{NOMBRE}}": nombre,
        "{{CEDULA}}": cedula,
        "{{EMPRESA}}": empresa,
        "{{FECHA}}": fecha,
    }
    for p in doc.paragraphs:
        for run in p.runs:
            for k, v in reemplazos.items():
                if k in run.text:
                    run.text = run.text.replace(k, v)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=f"Consentimiento_{nombre}.docx"
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

# ─────────────────────────────────────────────
#  BORRADORES — guardar / cargar / eliminar
# ─────────────────────────────────────────────
@app.route("/api/borrador/guardar", methods=["POST"])
def borrador_guardar():
    if not supa:
        return jsonify({"ok": False, "error": "Supabase no configurado"}), 503
    d = request.json or {}
    cedula = (d.get("cedula") or d.get("nroHistoria") or "").strip()
    if not cedula:
        return jsonify({"ok": False, "error": "Cédula requerida"}), 400
    try:
        # Upsert: si ya existe ese borrador lo sobreescribe
        supa.table("borradores_femo").upsert(
            {"cedula": cedula, "data": d},
            on_conflict="cedula"
        ).execute()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/borrador/<cedula>", methods=["GET"])
def borrador_cargar(cedula):
    if not supa:
        return jsonify({"ok": False, "error": "Supabase no configurado"}), 503
    try:
        r = supa.table("borradores_femo").select("data").eq("cedula", cedula).execute()
        if r.data:
            return jsonify({"ok": True, "data": r.data[0]["data"]})
        return jsonify({"ok": False, "error": "no_encontrado"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/borrador/<cedula>", methods=["DELETE"])
def borrador_eliminar(cedula):
    if not supa:
        return jsonify({"ok": False, "error": "Supabase no configurado"}), 503
    try:
        supa.table("borradores_femo").delete().eq("cedula", cedula).execute()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print("=" * 55)
    print("  FICHA MÉDICA OCUPACIONAL – CASISO S.A.S.")
    print(f"  Servidor en puerto {port}")
    print("=" * 55)
    app.run(host="0.0.0.0", port=port, debug=debug)
