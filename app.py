import streamlit as st
import pandas as pd
from datetime import datetime, date, time, timedelta
import pytz
import requests
import json
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from fpdf import FPDF

st.set_page_config(page_title="Control de Combustible", layout="wide", page_icon="⛽")

# ==========================================================
# CONFIGURACIONES GENERALES Y HORARIO
# ==========================================================
PRESUPUESTO_GLOBAL = 3800.00
BOLSA_COMODIN_TOTAL = 200.00
HORA_LIMITE = time(15, 10)  # ⏰ 3:10 PM
ZONA_HORARIA = pytz.timezone("America/Merida")

CONFIG_FILE = "config_sistema.json"
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzOjgha2Zjyog01t6LmA_R--EB4Ecqv2ifO_i2YJbLRLbXGShbu5uzFVi85FUTGplM8/exec"

PRESUPUESTO_BASE_POR_SOLICITANTE = {
    "COB CHAVEZ NARCISO DEL JESUS": 200.00,
    "PEREZ MAZIN CARLOS EDUARDO": 200.00,
    "DE LA CRUZ PEREZ WILLIAN ARLEY": 200.00,
    "NOEL CHAN": 850.00,
    "LIAN": 150.00,
    "QUEVEDO": 1500.00,
    "RENAN/HELDER": 500.00,
}

OPERADORES_POR_SOLICITANTE = {
    "COB CHAVEZ NARCISO DEL JESUS": ["JESUS COB"],
    "PEREZ MAZIN CARLOS EDUARDO": ["EDUARDO PEREZ"],
    "DE LA CRUZ PEREZ WILLIAN ARLEY": ["WILLIAN PEREZ", "FRANCISCO ALONZO"],
    "NOEL CHAN": ["AXEL SARAVIA", "NOEL CHAN", "ROMAN DZUL", "ROGER DUARTE", "LUIS CAHUICH", "FRANCISCO CONTRERAS"],
    "LIAN": ["FRANCISCO ALONZO"],
    "QUEVEDO": ["FARID PAVON", "WALDEMAR SAGUNDO", "JORGE MELIK"],
    "RENAN/HELDER": ["HELDER PACHECO", "RENAN CETINA"],
}

TXT_DESARROLLO_URBANO = "LLEVAR A CABO ACTIVIDADES DE INSPECCIONES, VERIFICACIONES Y SUPERVICIONES DE OBRAS Y OBSTRUCCIONES A LA VIA PÚBLICA CORRESPONDIENTES A LA SUBDIRECCION DE DESARROLLO URBANO"
TXT_MEDIO_AMBIENTE = "PARA LLEVAR A CABO INSPECCIONES A CARGO DE LA SUBDIRECCION DE MEDIO AMBIENTE, COMO LO SON ATENDER REPORTES POR TIRADERO DE AGUAS JABONOSAS, MALTRATO ANIMAL Y CONTAMINACION AUDITIVA, ASI COMO DIVERSOS TIPOS DE CONTAMINACION"
TXT_RAM_AMBIENTAL = "PARA LLEVAR A CABO ACTIVIDADES DE ESTERILIZACIONES DE PERROS Y GATOS, RECOLECCION DE MERMA DE FRUTAS Y VERDURAS EN SUPERMERCADOS Y REFORESTACIONES"

MAPEO_SOLICITANTES = {
    12: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "85GWU7", "actividad": TXT_DESARROLLO_URBANO, "base_sug": 100.0},
    13: {"solicita": "PEREZ MAZIN CARLOS EDUARDO", "vehiculo": "MOTO SUSUKI", "placa": "86GWU7", "actividad": TXT_DESARROLLO_URBANO, "base_sug": 200.0},
    14: {"solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "vehiculo": "MOTO SUSUKI", "placa": "86GWU8", "actividad": TXT_DESARROLLO_URBANO, "base_sug": 200.0},
    15: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "87GWU8", "actividad": TXT_DESARROLLO_URBANO, "base_sug": 100.0},
    16: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "88GWU7", "actividad": TXT_DESARROLLO_URBANO, "base_sug": 140.0},
    17: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "88GWU8", "actividad": TXT_MEDIO_AMBIENTE, "base_sug": 140.0},
    18: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU7", "actividad": TXT_MEDIO_AMBIENTE, "base_sug": 140.0},
    19: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU8", "actividad": TXT_MEDIO_AMBIENTE, "base_sug": 140.0},
    20: {"solicita": "NOEL CHAN", "vehiculo": "MOTO DINAMO", "placa": "90GWU7", "actividad": TXT_MEDIO_AMBIENTE, "base_sug": 145.0},
    21: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "90GWU8", "actividad": TXT_MEDIO_AMBIENTE, "base_sug": 145.0},
    22: {"solicita": "LIAN", "vehiculo": "MOTO SUSUKI", "placa": "91GWU7", "actividad": TXT_MEDIO_AMBIENTE, "base_sug": 150.0},
    23: {"solicita": "RENAN/HELDER", "vehiculo": "AUTOMOVIL JETTA", "placa": "DFT565C", "actividad": TXT_DESARROLLO_URBANO, "base_sug": 500.0},
    24: {"solicita": "QUEVEDO", "vehiculo": "CAMIONETA RAM 701", "placa": "CN2633B", "actividad": TXT_RAM_AMBIENTAL, "base_sug": 1500.0},
}

PASSWORDS_DEFAULT = {
    "LIAN": "admin123",
    "VERO": "distribucion123",
    "QUEVEDO": "ambiental2026"
}

MAPA_DIA_CARGA = {
    "Lunes": "MARTES",
    "Jueves": "VIERNES"
}

def obtener_nombre_archivo_oficial(prefijo, turno, f_prog, extension):
    dia_nombre = MAPA_DIA_CARGA.get(turno, turno.upper())
    fecha_str = f_prog.strftime("%d_%m_%Y")
    return f"{prefijo}_{dia_nombre}_{fecha_str}.{extension}"

# --- PERSISTENCIA LOCAL ---
def leer_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "desbloqueo_horario": False,
        "asignacion_comodin": {},
        "cesion_lian": {},
        "dia_activo": "Lunes",
        "passwords": PASSWORDS_DEFAULT.copy()
    }

def guardar_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f)
    except Exception:
        pass

def obtener_passwords():
    cfg = leer_config()
    pwds = cfg.get("passwords", {})
    actualizado = False
    for u, p in PASSWORDS_DEFAULT.items():
        if u not in pwds:
            pwds[u] = p
            actualizado = True
    if actualizado:
        cfg["passwords"] = pwds
        guardar_config(cfg)
    return pwds

def calcular_presupuesto_efectivo():
    cfg = leer_config()
    asig_comodin = cfg.get("asignacion_comodin", {})
    cesion_lian = cfg.get("cesion_lian", {})
    
    presupuestos = PRESUPUESTO_BASE_POR_SOLICITANTE.copy()
    for solicitante, extra in asig_comodin.items():
        if solicitante in presupuestos:
            presupuestos[solicitante] += float(extra)
            
    total_cedido_lian = 0.0
    for solicitante, monto in cesion_lian.items():
        if solicitante in presupuestos and solicitante != "LIAN":
            presupuestos[solicitante] += float(monto)
            total_cedido_lian += float(monto)
            
    presupuestos["LIAN"] = max(0.0, PRESUPUESTO_BASE_POR_SOLICITANTE["LIAN"] - total_cedido_lian)
    return presupuestos

def limpiar_texto_operador(val):
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ["none", "null", "nan", ""] else s

# --- CONSULTA Y ENVÍO A GOOGLE SHEETS ---
def obtener_datos_dos_hojas(forzar=False):
    if "df_lunes" in st.session_state and "df_jueves" in st.session_state and not forzar:
        df_chk = st.session_state.df_lunes
        if len(df_chk) == 13 and (df_chk[df_chk["row"] == 24]["Solicitante"].iloc[0] == "QUEVEDO"):
            return st.session_state.df_lunes.copy(), st.session_state.df_jueves.copy()

    filas_lunes = []
    filas_jueves = []
    
    try:
        res = requests.get(WEBHOOK_URL, timeout=8, allow_redirects=True)
        if res.status_code == 200:
            datos_json = res.json()
            raw_lunes = datos_json.get("lunes", [])
            raw_jueves = datos_json.get("jueves", [])
            
            dict_l = {int(it["row"]): it for it in raw_lunes if "row" in it}
            dict_j = {int(it["row"]): it for it in raw_jueves if "row" in it}
            
            for r in sorted(MAPEO_SOLICITANTES.keys()):
                info = MAPEO_SOLICITANTES[r]
                sol = info["solicita"]
                
                item_l = dict_l.get(r, {})
                op_l = limpiar_texto_operador(item_l.get("encargado", ""))
                imp_l = float(item_l.get("importe", 0.0)) if item_l.get("importe") else 0.0
                real_l = float(item_l.get("real", 0.0)) if item_l.get("real") else 0.0
                
                filas_lunes.append({
                    "row": r, "Solicitante": sol, "Vehículo": info["vehiculo"],
                    "Placa": info["placa"], "Actividad": info["actividad"],
                    "Base": info.get("base_sug", 0.0),
                    "Operador": op_l, "Importe": imp_l, "Real": real_l
                })
                
                item_j = dict_j.get(r, {})
                op_j = limpiar_texto_operador(item_j.get("encargado", ""))
                imp_j = float(item_j.get("importe", 0.0)) if item_j.get("importe") else 0.0
                real_j = float(item_j.get("real", 0.0)) if item_j.get("real") else 0.0
                
                filas_jueves.append({
                    "row": r, "Solicitante": sol, "Vehículo": info["vehiculo"],
                    "Placa": info["placa"], "Actividad": info["actividad"],
                    "Base": info.get("base_sug", 0.0),
                    "Operador": op_j, "Importe": imp_j, "Real": real_j
                })
    except Exception:
        pass

    if not filas_lunes:
        for r in sorted(MAPEO_SOLICITANTES.keys()):
            info = MAPEO_SOLICITANTES[r]
            filas_lunes.append({
                "row": r, "Solicitante": info["solicita"], "Vehículo": info["vehiculo"],
                "Placa": info["placa"], "Actividad": info["actividad"],
                "Base": info.get("base_sug", 0.0),
                "Operador": "", "Importe": 0.0, "Real": 0.0
            })
            filas_jueves.append({
                "row": r, "Solicitante": info["solicita"], "Vehículo": info["vehiculo"],
                "Placa": info["placa"], "Actividad": info["actividad"],
                "Base": info.get("base_sug", 0.0),
                "Operador": "", "Importe": 0.0, "Real": 0.0
            })

    df_l = pd.DataFrame(filas_lunes)
    df_j = pd.DataFrame(filas_jueves)
    
    st.session_state.df_lunes = df_l.copy()
    st.session_state.df_jueves = df_j.copy()
    return df_l, df_j

def enviar_datos_hoja(df_a_enviar, hoja="lunes", tipo="solicitado", f_elab=None, f_prog=None, historico_obj=None):
    payload = {
        "hoja": hoja,
        "tipo": tipo,
        "registros": []
    }
    if f_elab:
        payload["fecha_elaboro"] = f_elab.strftime("%d/%m/%Y")
    if f_prog:
        payload["fecha_prog"] = f_prog.strftime("%d/%m/%Y")
    if historico_obj:
        payload["historico"] = historico_obj

    for _, fila in df_a_enviar.iterrows():
        r = int(fila["row"])
        enc = limpiar_texto_operador(fila["Operador"])
        imp = fila["Importe"] if tipo == "solicitado" else fila["Real"]
        payload["registros"].append({
            "row": r,
            "encargado": enc,
            "importe": float(imp) if pd.notna(imp) else 0.0
        })

    key_state = "df_lunes" if hoja == "lunes" else "df_jueves"
    if key_state in st.session_state:
        df_mem = st.session_state[key_state]
        for _, r_env in df_a_enviar.iterrows():
            r = int(r_env["row"])
            mask = df_mem["row"] == r
            if tipo == "solicitado":
                df_mem.loc[mask, "Operador"] = limpiar_texto_operador(r_env["Operador"])
                df_mem.loc[mask, "Importe"] = r_env["Importe"]
            else:
                df_mem.loc[mask, "Real"] = r_env["Real"]
        st.session_state[key_state] = df_mem

    try:
        res = requests.post(WEBHOOK_URL, json=payload, timeout=12, allow_redirects=True)
        if res.status_code == 200:
            return res.json().get("status") == "success"
        return False
    except Exception:
        return False

# ==========================================
# GENERADORES DE ARCHIVOS OFICIALES
# ==========================================
def generar_excel_oficial_formato(df_datos, dia_nombre, f_elab, f_prog):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    ws = wb.active
    ws.title = dia_nombre.lower()
    ws.views.sheetView[0].showGridLines = True
    
    fuente_titulo = Font(name="Calibri", size=10, bold=True)
    fuente_sub = Font(name="Calibri", size=9, bold=True)
    fuente_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fuente_bold = Font(name="Calibri", size=9, bold=True)
    fuente_datos = Font(name="Calibri", size=8)
    
    fill_header_azul = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    border_cuadricula = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='
