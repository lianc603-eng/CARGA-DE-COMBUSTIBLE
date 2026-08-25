import streamlit as st
import pandas as pd
from datetime import datetime, date, time
import pytz
import requests
import json
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
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
    "DE LA CRUZ PEREZ WILLIAN ARLEY": ["WILLIAN PEREZ"],
    "NOEL CHAN": ["AXEL SARAVIA", "NOEL CHAN", "ROMAN DZUL", "ROGER DUARTE", "LUIS CAHUICH"],
    "LIAN": ["FRANCISCO ALONZO"],
    "QUEVEDO": ["FARID PAVON", "WALDEMAR SAGUNDO", "JORGE MELIK"],
    "RENAN/HELDER": ["HELDER PACHECO", "RENAN CETINA"],
}

TXT_DESARROLLO_URBANO = "LLEVAR A CABO ACTIVIDADES DE INSPECCIONES, VERIFICACIONES Y SUPERVICIONES DE OBRAS Y OBSTRUCCIONES A LA VIA PÚBLICA CORRESPONDIENTES A LA SUBDIRECCION DE DESARROLLO URBANO"
TXT_MEDIO_AMBIENTE = "PARA LLEVAR A CABO INSPECCIONES A CARGO DE LA SUBDIRECCION DE MEDIO AMBIENTE, COMO LO SON ATENDER REPORTES POR TIRADERO DE AGUAS JABONOSAS, MALTRATO ANIMAL Y CONTAMINACION AUDITIVA, ASI COMO DIVERSOS TIPOS DE CONTAMINACION"
TXT_RAM_AMBIENTAL = "PARA LLEVAR A CABO ACTIVIDADES DE ESTERILIZACIONES DE PERROS Y GATOS, RECOLECCION DE MERMA DE FRUTAS Y VERDURAS EN SUPERMERCADOS Y REFORESTACIONES"

MAPEO_SOLICITANTES = {
    12: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "85GWU7", "actividad": TXT_DESARROLLO_URBANO},
    13: {"solicita": "PEREZ MAZIN CARLOS EDUARDO", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "86GWU7", "actividad": TXT_DESARROLLO_URBANO},
    14: {"solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "86GWU8", "actividad": TXT_DESARROLLO_URBANO},
    15: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTOCICLETA HONDA", "placa": "88GWU7", "actividad": TXT_DESARROLLO_URBANO},
    16: {"solicita": "NOEL CHAN", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "88GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    17: {"solicita": "NOEL CHAN", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "89GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    18: {"solicita": "NOEL CHAN", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "89GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    19: {"solicita": "NOEL CHAN", "vehiculo": "MOTOCICLETA HONDA", "placa": "90GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    20: {"solicita": "NOEL CHAN", "vehiculo": "MOTOCICLETA HONDA", "placa": "90GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    21: {"solicita": "LIAN", "vehiculo": "MOTOCICLETA SUZUKI", "placa": "91GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    22: {"solicita": "QUEVEDO", "vehiculo": "CAMIONETA RAM", "placa": "CN2633B", "actividad": TXT_RAM_AMBIENTAL},
    23: {"solicita": "RENAN/HELDER", "vehiculo": "AUTOMOVIL JETTA", "placa": "DFT565C", "actividad": TXT_DESARROLLO_URBANO},
}

USUARIOS_PASSWORD = {
    "LIAN": "admin123",
    "NOEL CHAN": "inspeccion2026",
    "QUEVEDO": "ambiental2026",
    "RENAN/HELDER": "urbano2026",
    "COB CHAVEZ NARCISO DEL JESUS": "notif123",
    "PEREZ MAZIN CARLOS EDUARDO": "notif123",
    "DE LA CRUZ PEREZ WILLIAN ARLEY": "notif123",
}

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
        "dia_activo": "Lunes"
    }

def guardar_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f)
    except Exception:
        pass

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

# --- GESTIÓN DE DATOS ---
def obtener_datos_sheets(forzar=False):
    if "df_datos_persistentes" in st.session_state and not forzar:
        return st.session_state.df_datos_persistentes.copy()

    filas = []
    try:
        res = requests.get(WEBHOOK_URL, timeout=8, allow_redirects=True)
        if res.status_code == 200:
            datos_raw = res.json().get("data", [])
            for item in datos_raw:
                r = int(item.get("row", 0))
                if r in MAPEO_SOLICITANTES:
                    filas.append({
                        "row": r,
                        "Solicitante": MAPEO_SOLICITANTES[r]["solicita"],
                        "Vehículo": MAPEO_SOLICITANTES[r]["vehiculo"],
                        "Placa": MAPEO_SOLICITANTES[r]["placa"],
                        "Actividad": MAPEO_SOLICITANTES[r]["actividad"],
                        "Operador_Lunes": str(item.get("encargado_lunes", "")).strip(),
                        "Importe_Lunes": float(item.get("importe_lunes", 0.0)) if item.get("importe_lunes") else 0.0,
                        "Operador_Jueves": str(item.get("encargado_jueves", "")).strip(),
                        "Importe_Jueves": float(item.get("importe_jueves", 0.0)) if item.get("importe_jueves") else 0.0,
                        "Importe_Real": float(item.get("importe_real", 0.0)) if item.get("importe_real") else 0.0
                    })
    except Exception:
        pass

    if not filas:
        if "df_datos_persistentes" in st.session_state:
            return st.session_state.df_datos_persistentes.copy()
            
        for r, info in MAPEO_SOLICITANTES.items():
            filas.append({
                "row": r,
                "Solicitante": info["solicita"],
                "Vehículo": info["vehiculo"],
                "Placa": info["placa"],
                "Actividad": info["actividad"],
                "Operador_Lunes": "",
                "Importe_Lunes": 0.0,
                "Operador_Jueves": "",
                "Importe_Jueves": 0.0,
                "Importe_Real": 0.0
            })
            
    df_res = pd.DataFrame(filas)
    st.session_state.df_datos_persistentes = df_res.copy()
    return df_res

def enviar_datos_sheets(registros_a_enviar, turno="lunes", f_elab=None, f_prog=None, historico_obj=None):
    payload = {"tipo": turno, "turno": turno, "registros": []}
    if f_elab:
        payload["fecha_elaboro"] = f_elab.strftime("%d/%m/%Y")
    if f_prog:
        payload["fecha_prog"] = f_prog.strftime("%d/%m/%Y")
    if historico_obj:
        payload["historico"] = historico_obj
        
    for _, fila in registros_a_enviar.iterrows():
        if turno == "lunes":
            enc = fila["Operador_Lunes"]
            imp = fila["Importe_Lunes"]
        elif turno == "jueves":
            enc = fila["Operador_Jueves"]
            imp = fila["Importe_Jueves"]
        else:
            enc = fila.get("Operador_Lunes", "")
            imp = fila["Importe_Real"]

        payload["registros"].append({
            "row": int(fila["row"]),
            "encargado": str(enc).strip() if pd.notna(enc) else "",
            "importe": float(imp) if pd.notna(imp) else 0.0
        })
        
    if "df_datos_persistentes" in st.session_state:
        df_mem = st.session_state.df_datos_persistentes
        for _, r_env in registros_a_enviar.iterrows():
            mask = df_mem["row"] == r_env["row"]
            if turno == "lunes":
                df_mem.loc[mask, "Operador_Lunes"] = r_env["Operador_Lunes"]
                df_mem.loc[mask, "Importe_Lunes"] = r_env["Importe_Lunes"]
            elif turno == "jueves":
                df_mem.loc[mask, "Operador_Jueves"] = r_env["Operador_Jueves"]
                df_mem.loc[mask, "Importe_Jueves"] = r_env["Importe_Jueves"]
            else:
                df_mem.loc[mask, "Importe_Real"] = r_env["Importe_Real"]
        st.session_state.df_datos_persistentes = df_mem

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
def generar_excel_oficial_formato(df_datos, dia_reporte, f_elab, f_prog):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    ws = wb.active
    ws.title = "carga"
    ws.views.sheetView[0].showGridLines = True
    
    fuente_titulo = Font(name="Calibri", size=10, bold=True)
    fuente_sub = Font(name="Calibri", size=9, bold=True)
    fuente_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fuente_bold = Font(name="Calibri", size=9, bold=True)
    fuente_datos = Font(name="Calibri", size=8)
    
    fill_header_azul = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    border_cuadricula = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )

    ws["D2"] = "H. AYUNTAMIENTO DE CAMPECHE"
    ws["D2"].font = fuente_titulo
    ws["D2"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws["B4"] = "Unidad:"
    ws["B4"].font = fuente_sub
    ws["C4"] = "DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE"
    ws["C4"].font = fuente_sub
    
    ws["B5"] = "Subdireccion:"
    ws["B5"].font = fuente_sub
    
    ws["H8"] = "Elaboro:"
    ws["H8"].font = fuente_sub
    ws["H8"].alignment = Alignment(horizontal="right")
    ws["I8"] = f_elab.strftime("%d/%m/%Y")
    ws["I8"].font = fuente_sub
    
    ws["H9"] = "Programacion para el dia:"
    ws["H9"].font = fuente_sub
    ws["H9"].alignment = Alignment(horizontal="right")
    ws["I9"] = f_prog.strftime("%d/%m/%Y")
    ws["I9"].font = fuente_sub

    headers_oficiales = [
        (2, "NOMBRE DEL\nENCARGADO"), (3, "VEHÍCULO"), (4, "PLACA"),
        (5, "OFICIAL /\nCOMODATO"), (6, "LITROS"), (7, "IMPORTE"),
        (8, "MAGNA / DIESEL"), (9, "ACTIVIDAD")
    ]
    
    for col_num, h_text in headers_oficiales:
        cell = ws.cell(row=11, column=col_num, value=h_text)
        cell.font = fuente_header
        cell.fill = fill_header_azul
        cell.border = border_cuadricula
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[11].height = 28

    col_op = "Operador_Lunes" if dia_reporte == "Lunes" else "Operador_Jueves"
    col_imp = "Importe_Lunes" if dia_reporte == "Lunes" else "Importe_Jueves"

    for r_num in sorted(MAPEO_SOLICITANTES.keys()):
        row_info = df_datos[df_datos["row"] == r_num]
        
        if not row_info.empty:
            item = row_info.iloc[0]
            op_nombre = str(item[col_op]).strip()
            veh = str(item["Vehículo"]).strip()
            plc = str(item["Placa"]).strip()
            imp_val = float(item[col_imp])
            act_text = str(item["Actividad"]).strip()
        else:
            op_nombre = ""
            veh = MAPEO_SOLICITANTES[r_num]["vehiculo"]
            plc = MAPEO_SOLICITANTES[r_num]["placa"]
            imp_val = 0.0
            act_text = MAPEO_SOLICITANTES[r_num]["actividad"]

        ws.cell(row=r_num, column=2, value=op_nombre).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r_num, column=3, value=veh).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r_num, column=4, value=plc).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_num, column=5, value="OFICIAL").alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_num, column=6, value="").alignment = Alignment(horizontal="center", vertical="center")
        
        c_imp = ws.cell(row=r_num, column=7, value=imp_val)
        c_imp.number_format = '$#,##0.00'
        c_imp.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.cell(row=r_num, column=8, value="MAGNA").alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_num, column=9, value=act_text).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col_i in range(2, 10):
            c_est = ws.cell(row=r_num, column=col_i)
            c_est.font = fuente_datos
            c_est.border = border_cuadricula
            
        ws.row_dimensions[r_num].height = 24
        if imp_val == 0.0:
            ws.row_dimensions[r_num].hidden = True

    c_lbl_tot = ws.cell(row=24, column=6, value="TOTAL")
    c_lbl_tot.font = fuente_bold
    c_lbl_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_lbl_tot.border = border_cuadricula
    
    c_val_tot = ws.cell(row=24, column=7, value="=SUM(G12:G23)")
    c_val_tot.font = fuente_bold
    c_val_tot.number_format = '$#,##0.00'
    c_val_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_val_tot.border = border_cuadricula
    
    for c_rest in [2, 3, 4, 5, 8, 9]:
        ws.cell(row=24, column=c_rest).border = border_cuadricula

    anchos_cols = {'A': 3, 'B': 24, 'C': 18, 'D': 12, 'E': 14, 'F': 10, 'G': 14, 'H': 14, 'I': 52}
    for col_letra, ancho in anchos_cols.items():
        ws.column_dimensions[col_letra].width = ancho

    wb.save(output)
    return output.getvalue()

def generar_pdf_oficial(df_cargas, dia_reporte, f_elab, f_prog):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 6, 'H. AYUNTAMIENTO DE CAMPECHE', 0, 1, 'C')
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(0, 5, f'DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE - PROGRAMACION {dia_reporte.upper()}', 0, 1, 'C')
    pdf.ln(3)
    
    pdf.set_font('Helvetica', '', 9)
    pdf.cell(140, 5, 'Unidad: DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE', 0, 0, 'L')
    pdf.cell(130, 5, f'Elaboro: {f_elab.strftime("%d/%m/%Y")}', 0, 1, 'R')
    pdf.cell(140, 5, '', 0, 0, 'L')
    pdf.cell(130, 5, f'Programacion para el dia: {f_prog.strftime("%d/%m/%Y")}', 0, 1, 'R')
    pdf.ln(4)
    
    col_widths = (45, 32, 18, 18, 22, 16, 117)
    col_op = "Operador_Lunes" if dia_reporte == "Lunes" else "Operador_Jueves"
    col_imp = "Importe_Lunes" if dia_reporte == "Lunes" else "Importe_Jueves"
    
    with pdf.table(
        col_widths=col_widths, 
        text_align=("LEFT", "LEFT", "CENTER", "CENTER", "RIGHT", "CENTER", "LEFT"),
        line_height=4.5
    ) as table:
        pdf.set_font('Helvetica', 'B', 8)
        header_row = table.row()
        for h in ['NOMBRE DEL ENCARGADO', 'VEHICULO', 'PLACA', 'REGIMEN', 'IMPORTE', 'TIPO', 'ACTIVIDAD']:
            header_row.cell(h)
            
        pdf.set_font('Helvetica', '', 7.5)
        total = 0.0
        for _, row in df_cargas.iterrows():
            data_row = table.row()
            data_row.cell(str(row[col_op]).strip())
            data_row.cell(str(row["Vehículo"]).strip())
            data_row.cell(str(row["Placa"]).strip())
            data_row.cell('OFICIAL')
            data_row.cell(f"${float(row[col_imp]):,.2f}")
            data_row.cell('MAGNA')
            data_row.cell(str(row["Actividad"]).strip())
            total += float(row[col_imp])
            
        pdf.set_font('Helvetica', 'B', 8)
        tot_row = table.row()
        tot_row.cell(f'TOTAL AUTORIZADO {dia_reporte.upper()}:', colspan=4, align="RIGHT")
        tot_row.cell(f"${total:,.2f}", align="RIGHT")
        tot_row.cell('', colspan=2)
        
    return bytes(pdf.output())

# ==========================================
# 1. INICIO DE SESIÓN
# ==========================================
if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None
if "vista_simulada" not in st.session_state:
    st.session_state.vista_simulada = None

if st.session_state.usuario_logueado is None:
    st.title("⛽ Solicitud de Combustible")
    st.caption("Dirección de Desarrollo Urbano y Medio Ambiente")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("form_login"):
            st.subheader("🔐 Iniciar Sesión")
            usr = st.selectbox("Selecciona tu Usuario", list(USUARIOS_PASSWORD.keys()))
            pwd = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar", use_container_width=True):
                if pwd == USUARIOS_PASSWORD.get(usr):
                    st.session_state.usuario_logueado = usr
                    st.session_state.vista_simulada = None
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
    st.stop()

# ==========================================
# 2. ENCABEZADO Y SELECTOR DE TURNO (LUNES / JUEVES)
# ==========================================
usuario_real = st.session_state.usuario_logueado
es_admin_real = (usuario_real == "LIAN")
usuario_efectivo = st.session_state.vista_simulada if (es_admin_real and st.session_state.vista_simulada) else usuario_real
es_admin = (usuario_efectivo == "LIAN")

ahora_local = datetime.now(ZONA_HORARIA)
hora_actual = ahora_local.time()

cfg_actual = leer_config()
desbloqueo_activo = cfg_actual.get("desbloqueo_horario", False)
sistema_bloqueado = (hora_actual >= HORA_LIMITE) and not es_admin_real and not desbloqueo_activo

presupuestos_actuales = calcular_presupuesto_efectivo()

c1, c2, c3 = st.columns([2.5, 1.2, 0.8])
with c1:
    st.title("⛽ Control de Combustible")
    if es_admin_real and st.session_state.vista_simulada:
        st.warning(f"🧪 **MODO DE PRUEBA ACTIVO**: Simulando vista como **{usuario_efectivo}**")
    else:
        st.markdown("👑 **ADMINISTRADOR GENERAL**" if es_admin else f"👤 Solicitante: **{usuario_efectivo}**")
        
    estado_horario = "🟢 Horario Abierto" if not sistema_bloqueado else "🔴 Horario Cerrado (Límite 3:10 PM)"
    if desbloqueo_activo:
        estado_horario += " (⚡ Desbloqueo temporal activo)"
    st.caption(f"🕒 {ahora_local.strftime('%I:%M %p')} | {estado_horario}")

with c2:
    dia_guardado = cfg_actual.get("dia_activo", "Lunes")
    if es_admin:
        dia_seleccionado = st.radio("📅 Turno de Carga Activo:", ["Lunes", "Jueves"], horizontal=True, index=0 if dia_guardado == "Lunes" else 1)
        if dia_seleccionado != dia_guardado:
            cfg_actual["dia_activo"] = dia_seleccionado
            guardar_config(cfg_actual)
            st.rerun()
    else:
        dia_seleccionado = dia_guardado
        st.info(f"📅 Solicitud para: **{dia_seleccionado}**")

with c3:
    st.write("")
    if st.button("🔄 Sincronizar", use_container_width=True):
        obtener_datos_sheets(forzar=True)
        st.toast("Datos sincronizados.", icon="✅")
        st.rerun()
    if es_admin_real and st.session_state.vista_simulada:
        if st.button("⬅️ Volver", use_container_width=True):
            st.session_state.vista_simulada = None
            st.rerun()
    else:
        if st.button("🚪 Salir", use_container_width=True):
            st.session_state.usuario_logueado = None
            st.session_state.vista_simulada = None
            st.rerun()

df_actual = obtener_datos_sheets()

# ==========================================
# 3. VISTA SOLICITANTE (CON CONTROL SEMANAL LUNES / JUEVES)
# ==========================================
if not es_admin:
    presupuesto_semanal = presupuestos_actuales.get(usuario_efectivo, 0.00)
    df_solicitante = df_actual[df_actual["Solicitante"] == usuario_efectivo].copy()
    
    total_lunes_sol = df_solicitante["Importe_Lunes"].sum()
    total_jueves_sol = df_solicitante["Importe_Jueves"].sum()
    
    # Si estamos en Jueves, el saldo disponible es: Presupuesto Semanal - Lo que cargó el Lunes
    if dia_seleccionado == "Lunes":
        disponible_para_hoy = presupuesto_semanal
        cargado_anterior = 0.0
    else:
        disponible_para_hoy = max(0.0, presupuesto_semanal - total_lunes_sol)
        cargado_anterior = total_lunes_sol
        
    lista_operadores_autorizados = OPERADORES_POR_SOLICITANTE.get(usuario_efectivo, [])
    
    if sistema_bloqueado:
        st.error("🔒 **SISTEMA CERRADO POR HORARIO (3:10 PM)**. La captura ha finalizado.")
    else:
        if dia_seleccionado == "Jueves":
            st.info(f"💡 Ya registraste **${cargado_anterior:,.2f}** el Lunes. Tu saldo restante para este Jueves es **${disponible_para_hoy:,.2f}**.")
        st.caption(f"📱 Selecciona los datos para la carga del **{dia_seleccionado}**:")
    
    with st.form("form_solicitante_movil"):
        nuevos_valores = []
        
        col_op_target = "Operador_Lunes" if dia_seleccionado == "Lunes" else "Operador_Jueves"
        col_imp_target = "Importe_Lunes" if dia_seleccionado == "Lunes" else "Importe_Jueves"
        
        for idx, row in df_solicitante.iterrows():
            with st.container(border=True):
                st.markdown(f"🛵 **{row['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row['Placa']}`**")
                st.caption(f"📋 **Actividad:** {row['Actividad']}")
                
                c_op, c_imp = st.columns([1.5, 1])
                opciones_operadores = [""] + lista_operadores_autorizados
                val_actual = str(row[col_op_target]).strip()
                
                if val_actual and val_actual not in opciones_operadores:
                    opciones_operadores.append(val_actual)
                    
                idx_sel = opciones_operadores.index(val_actual) if val_actual in opciones_operadores else 0
                
                with c_op:
                    val_encargado = st.selectbox(
                        "Operador / Conductor",
                        options=opciones_operadores,
                        index=idx_sel,
                        key=f"op_{dia_seleccionado}_{row['row']}",
                        disabled=sistema_bloqueado
                    )
                    
                with c_imp:
                    val_importe = st.number_input(
                        f"Monto {dia_seleccionado} ($)",
                        value=float(row[col_imp_target]),
                        step=50.0,
                        min_value=0.0,
                        key=f"imp_{dia_seleccionado}_{row['row']}",
                        disabled=sistema_bloqueado,
                        format="%.2f"
                    )
                
                nuevos_valores.append({
                    "row": row["row"],
                    "Solicitante": row["Solicitante"],
                    "Vehículo": row["Vehículo"],
                    "Placa": row["Placa"],
                    "Actividad": row["Actividad"],
                    "Operador_Lunes": val_encargado if dia_seleccionado == "Lunes" else row["Operador_Lunes"],
                    "Importe_Lunes": val_importe if dia_seleccionado == "Lunes" else row["Importe_Lunes"],
                    "Operador_Jueves": val_encargado if dia_seleccionado == "Jueves" else row["Operador_Jueves"],
                    "Importe_Jueves": val_importe if dia_seleccionado == "Jueves" else row["Importe_Jueves"],
                    "Importe_Real": row["Importe_Real"]
                })
        
        df_edit_movil = pd.DataFrame(nuevos_valores)
        total_capturado_hoy = df_edit_movil[col_imp_target].sum()
        saldo_restante_hoy = disponible_para_hoy - total_capturado_hoy
        
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Presupuesto Semanal", f"${presupuesto_semanal:,.2f}")
        m2.metric("Cargado el Lunes", f"${cargado_anterior:,.2f}")
        m3.metric(f"Capturado {dia_seleccionado}", f"${total_capturado_hoy:,.2f}")
        m4.metric("Saldo Disponible", f"${saldo_restante_hoy:,.2f}", delta_color="normal" if saldo_restante_hoy >= 0 else "off")
        
        if total_capturado_hoy > disponible_para_hoy:
            st.error(f"⚠️ Excedes el saldo disponible para este {dia_seleccionado} por **${abs(saldo_restante_hoy):,.2f} MXN**.")
            
        btn_guardar = st.form_submit_button(f"💾 Guardar Solicitud de {dia_seleccionado}", type="primary", use_container_width=True, disabled=sistema_bloqueado)
        
        if btn_guardar:
            if total_capturado_hoy > disponible_para_hoy:
                st.error("No puedes guardar si excedes el saldo disponible.")
            else:
                with st.spinner(f"Guardando solicitud del {dia_seleccionado}..."):
                    exito = enviar_datos_sheets(df_edit_movil, turno=dia_seleccionado.lower())
                    if exito:
                        st.success(f"✅ ¡Solicitud del {dia_seleccionado} guardada con éxito en Google Sheets!")
                    else:
                        st.error("Error al comunicarse con Google Sheets.")

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Panel de Consolidación, Edición y Descarga Oficial")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración", value=date.today())
    with col_f2:
        f_prog = st.date_input(f"Programación para el {dia_seleccionado}", value=date.today())

    tab_saldos, tab_solicitud_final, tab_mi_carga, tab_auditoria, tab_mantenimiento = st.tabs([
        "📊 Monitoreo Semanal",
        f"✏️ Solicitud Final ({dia_seleccionado})",
        "🛵 Mi Carga (LIAN)",
        "✅ Cierre Semanal y Auditoría",
        "🛠️ Modo Pruebas y Cierre de Viernes"
    ])

    # 1. MONITOREO SEMANAL COMPLETO (LUNES + JUEVES)
    with tab_saldos:
        st.markdown(f"##### 💵 Balance Semanal de Presupuestos (Viendo: Turno {dia_seleccionado})")
        
        total_lunes_global = df_actual["Importe_Lunes"].sum()
        total_jueves_global = df_actual["Importe_Jueves"].sum()
        total_semana_global = total_lunes_global + total_jueves_global
        saldo_global_disponible = PRESUPUESTO_GLOBAL - total_semana_global
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Presupuesto Semanal Total", f"${PRESUPUESTO_GLOBAL:,.2f}")
        c2.metric("Total Lunes", f"${total_lunes_global:,.2f}")
        c3.metric("Total Jueves", f"${total_jueves_global:,.2f}")
        c4.metric(
            "Saldo Libre Semana", 
            f"${saldo_global_disponible:,.2f}", 
            delta=f"${total_semana_global:,.2f} total pedido",
            delta_color="normal" if saldo_global_disponible >= 0 else "inverse"
        )
        
        st.write("")
        st.markdown("##### 📋 Resumen por Solicitante (Lunes + Jueves)")
        filas_saldos_area = []
        for sol, p_efectivo in presupuestos_actuales.items():
            sub_df = df_actual[df_actual["Solicitante"] == sol]
            sol_lunes = sub_df["Importe_Lunes"].sum()
            sol_jueves = sub_df["Importe_Jueves"].sum()
            sol_total = sol_lunes + sol_jueves
            disp_monto = p_efectivo - sol_total
            
            if disp_monto == 0:
                estatus = "✅ 100% Ejercido"
            elif disp_monto < 0:
                estatus = "⚠️ Excedido"
            elif sol_total > 0:
                estatus = "🟢 Con Saldo"
            else:
                estatus = "⚪ Sin Carga"
                
            filas_saldos_area.append({
                "Solicitante / Área": sol,
                "Presupuesto Semanal": p_efectivo,
                "Carga Lunes": sol_lunes,
                "Carga Jueves": sol_jueves,
                "Total Solicitado": sol_total,
                "Saldo Restante": disp_monto,
                "Estatus": estatus
            })
            
        df_saldos_area = pd.DataFrame(filas_saldos_area)
        st.dataframe(
            df_saldos_area,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Presupuesto Semanal": st.column_config.NumberColumn(format="$%.2f"),
                "Carga Lunes": st.column_config.NumberColumn(format="$%.2f"),
                "Carga Jueves": st.column_config.NumberColumn(format="$%.2f"),
                "Total Solicitado": st.column_config.NumberColumn(format="$%.2f"),
                "Saldo Restante": st.column_config.NumberColumn(format="$%.2f"),
            }
        )

    # 2. SOLICITUD FINAL Y EDITOR (POR DÍA LUNES / JUEVES)
    with tab_solicitud_final:
        st.markdown(f"##### 🚗 Solicitud Oficial para el día: **{dia_seleccionado}**")
        st.caption(f"Edita nombres e importes correspondientes a la carga del **{dia_seleccionado}**:")
        
        todos_los_operadores = [""]
        for l_ops in OPERADORES_POR_SOLICITANTE.values():
            for o in l_ops:
                if o not in todos_los_operadores:
                    todos_los_operadores.append(o)
                    
        col_op_target = "Operador_Lunes" if dia_seleccionado == "Lunes" else "Operador_Jueves"
        col_imp_target = "Importe_Lunes" if dia_seleccionado == "Lunes" else "Importe_Jueves"
        
        df_admin_edit = st.data_editor(
            df_actual.copy(),
            use_container_width=True,
            disabled=["row", "Solicitante", "Vehículo", "Placa", "Actividad", "Importe_Real"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Área"),
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placa"),
                col_op_target: st.column_config.SelectboxColumn(
                    f"Operador {dia_seleccionado}",
                    options=todos_los_operadores,
                    width="medium"
                ),
                col_imp_target: st.column_config.NumberColumn(f"Importe {dia_seleccionado} ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "Actividad": st.column_config.TextColumn("Actividad", width="medium"),
                "row": None, "Operador_Lunes": None if dia_seleccionado == "Lunes" else "disabled",
                "Importe_Lunes": None if dia_seleccionado == "Lunes" else "disabled",
                "Operador_Jueves": None if dia_seleccionado == "Jueves" else "disabled",
                "Importe_Jueves": None if dia_seleccionado == "Jueves" else "disabled",
                "Importe_Real": None
            },
            hide_index=True,
            key=f"admin_solicitudes_editor_{dia_seleccionado}"
        )
        
        if st.button(f"💾 Guardar Cargas del {dia_seleccionado} en Google Sheets", type="primary", use_container_width=True):
            with st.spinner(f"Guardando {dia_seleccionado} en Google Sheets..."):
                exito = enviar_datos_sheets(df_admin_edit, turno=dia_seleccionado.lower(), f_elab=f_elab, f_prog=f_prog)
                if exito:
                    st.success(f"✅ ¡Cargas del {dia_seleccionado} guardadas exitosamente!")
                else:
                    st.error("Error al guardar en Google Sheets.")

        st.markdown("---")
        df_solo_cargas_dia = df_admin_edit[df_admin_edit[col_imp_target] > 0].copy()
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            excel_bytes = generar_excel_oficial_formato(df_admin_edit, dia_seleccionado, f_elab, f_prog)
            st.download_button(
                label=f"📥 Descargar Formato Oficial Excel ({dia_seleccionado})",
                data=excel_bytes,
                file_name=f"SOLICITUD_{dia_seleccionado.upper()}_{f_prog.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_d2:
            pdf_bytes = generar_pdf_oficial(df_solo_cargas_dia, dia_seleccionado, f_elab, f_prog)
            st.download_button(
                label=f"📄 Descargar Oficio Oficial en PDF ({dia_seleccionado})",
                data=pdf_bytes,
                file_name=f"OFICIO_{dia_seleccionado.upper()}_{f_prog.strftime('%d%m%Y')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    # 3. MI CARGA (LIAN)
    with tab_mi_carga:
        st.markdown(f"##### 🛵 Mi Carga (LIAN) - Turno: **{dia_seleccionado}**")
        df_lian = df_actual[df_actual["Solicitante"] == "LIAN"].copy()
        presupuesto_lian_efectivo = presupuestos_actuales["LIAN"]
        
        if not df_lian.empty:
            row_lian = df_lian.iloc[0]
            col_op_lian = "Operador_Lunes" if dia_seleccionado == "Lunes" else "Operador_Jueves"
            col_imp_lian = "Importe_Lunes" if dia_seleccionado == "Lunes" else "Importe_Jueves"
            
            cargado_anterior_lian = float(row_lian["Importe_Lunes"]) if dia_seleccionado == "Jueves" else 0.0
            disponible_lian_hoy = max(0.0, presupuesto_lian_efectivo - cargado_anterior_lian)
            
            with st.container(border=True):
                st.markdown(f"🛵 **{row_lian['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row_lian['Placa']}`** &nbsp;|&nbsp; Saldo Disponible Hoy: **${disponible_lian_hoy:,.2f}**")
                
                c_op_l, c_imp_l = st.columns([1.5, 1])
                with c_op_l:
                    opciones_lian = ["", "FRANCISCO ALONZO"]
                    val_op_lian = st.selectbox(
                        f"Operador {dia_seleccionado}",
                        options=opciones_lian,
                        index=1 if row_lian[col_op_lian] == "FRANCISCO ALONZO" else 0,
                        key=f"admin_op_lian_{dia_seleccionado}"
                    )
                with c_imp_l:
                    val_imp_lian = st.number_input(
                        f"Importe {dia_seleccionado} ($)", 
                        value=float(row_lian[col_imp_lian]),
                        step=50.0,
                        min_value=0.0,
                        max_value=float(disponible_lian_hoy),
                        key=f"admin_imp_lian_{dia_seleccionado}",
                        format="%.2f"
                    )
                
                if st.button(f"💾 Guardar Mi Carga de {dia_seleccionado}", type="primary", use_container_width=True):
                    df_mi_carga = df_actual[df_actual["Solicitante"] == "LIAN"].copy()
                    df_mi_carga[col_op_lian] = val_op_lian
                    df_mi_carga[col_imp_lian] = val_imp_lian
                    exito = enviar_datos_sheets(df_mi_carga, turno=dia_seleccionado.lower(), f_elab=f_elab, f_prog=f_prog)
                    if exito:
                        st.success(f"✅ Tu carga del {dia_seleccionado} fue registrada correctamente.")

    # 4. CIERRE SEMANAL Y AUDITORÍA
    with tab_auditoria:
        st.markdown("##### 🔍 Cierre Semanal: Conciliación Total (Lunes + Jueves vs Real)")
        st.caption("Captura el monto total ejercido en la semana completa para registrar el ahorro definitivo:")
        
        df_audit = df_actual.copy()
        df_audit["Total_Semanal_Solicitado"] = df_audit["Importe_Lunes"] + df_audit["Importe_Jueves"]
        
        df_real_edit = st.data_editor(
            df_audit,
            use_container_width=True,
            disabled=["row", "Solicitante", "Vehículo", "Placa", "Actividad", "Total_Semanal_Solicitado"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Área"),
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placa"),
                "Total_Semanal_Solicitado": st.column_config.NumberColumn("Total Solicitado Semana ($)", format="$%.2f"),
                "Importe_Real": st.column_config.NumberColumn("Total Realmente Cargado ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "row": None, "Actividad": None, "Operador_Lunes": None, "Importe_Lunes": None,
                "Operador_Jueves": None, "Importe_Jueves": None
            },
            hide_index=True,
            key="editor_cierre_semanal"
        )
        
        tot_sol_semana = df_real_edit["Total_Semanal_Solicitado"].sum()
        tot_real_semana = df_real_edit["Importe_Real"].sum()
        ahorro_semana = tot_sol_semana - tot_real_semana
        saldo_global_libre = PRESUPUESTO_GLOBAL - tot_real_semana
        
        st.divider()
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Total Solicitado Semana", f"${tot_sol_semana:,.2f}")
        r2.metric("Total Real Ejercido", f"${tot_real_semana:,.2f}")
        r3.metric("Ahorro / Remanente", f"${ahorro_semana:,.2f}", delta_color="normal")
        r4.metric("Saldo Disponible Total", f"${saldo_global_libre:,.2f}", delta_color="normal")
        
        st.write("")
        if st.button("💾 Guardar Cierre Semanal y Archivar en Histórico", type="primary", use_container_width=True):
            with st.spinner("Guardando cierre semanal..."):
                cargas_efectuadas = df_real_edit[df_real_edit["Importe_Real"] > 0]
                detalles_txt_lista = []
                for _, r_c in cargas_efectuadas.iterrows():
                    detalles_txt_lista.append(f"{r_c['Vehículo']} - ${r_c['Importe_Real']:,.2f}")
                detalles_txt = "; ".join(detalles_txt_lista) if detalles_txt_lista else "Sin cargas reales"
                
                folio_generado = f"SEMANA-{f_prog.strftime('%Y%m%d')}"
                
                historico_payload = {
                    "folio": folio_generado,
                    "fecha_elaboro": f_elab.strftime("%d/%m/%Y"),
                    "fecha_prog": f_prog.strftime("%d/%m/%Y"),
                    "fecha_registro": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p"),
                    "total_solicitado": float(tot_sol_semana),
                    "total_ejercido": float(tot_real_semana),
                    "ahorro": float(ahorro_semana),
                    "detalle_unidades": detalles_txt
                }
                
                exito = enviar_datos_sheets(
                    df_real_edit, 
                    turno="real", 
                    f_elab=f_elab, 
                    f_prog=f_prog, 
                    historico_obj=historico_payload
                )
                
                if exito:
                    st.success(f"✅ ¡Cierre semanal archivado con éxito con el Folio **{folio_generado}**!")

    # 5. REESTABLECIMIENTO DE VIERNES
    with tab_mantenimiento:
        st.markdown("##### 🛠️ Cierre de Semana (Viernes) y Mantenimiento")
        
        with st.container(border=True):
            st.subheader("🧹 Reestablecer Semana a $0.00 (Cada Viernes)")
            st.write(
                "Usa este botón los **viernes** tras guardar el cierre semanal. Limpiará las cargas de Lunes y Jueves para que la siguiente semana arranque en blanco y con presupuestos completos."
            )
            
            if st.button("🗑️ Reestablecer Todo para la Próxima Semana", type="secondary", use_container_width=True):
                df_limpio = df_actual.copy()
                df_limpio["Operador_Lunes"] = ""
                df_limpio["Importe_Lunes"] = 0.0
                df_limpio["Operador_Jueves"] = ""
                df_limpio["Importe_Jueves"] = 0.0
                df_limpio["Importe_Real"] = 0.0
                
                enviar_datos_sheets(df_limpio, turno="lunes", f_elab=f_elab, f_prog=f_prog)
                enviar_datos_sheets(df_limpio, turno="jueves", f_elab=f_elab, f_prog=f_prog)
                enviar_datos_sheets(df_limpio, turno="real", f_elab=f_elab, f_prog=f_prog)
                
                cfg_actual["dia_activo"] = "Lunes"
                guardar_config(cfg_actual)
                st.success("✅ ¡Semana reestablecida a $0.00! El sistema quedó listo para el próximo Lunes.")
                st.rerun()

        with st.container(border=True):
            st.subheader("🧪 Simular Vista de Solicitante")
            usuarios_para_test = [u for u in USUARIOS_PASSWORD.keys() if u != "LIAN"]
            solicitante_a_testear = st.selectbox("Selecciona solicitante a simular:", usuarios_para_test)
            if st.button("👁️ Entrar a Modo Simulación", type="secondary", use_container_width=True):
                st.session_state.vista_simulada = solicitante_a_testear
                st.rerun()
