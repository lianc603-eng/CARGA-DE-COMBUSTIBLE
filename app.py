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
    12: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "85GWU7", "actividad": TXT_DESARROLLO_URBANO},
    13: {"solicita": "PEREZ MAZIN CARLOS EDUARDO", "vehiculo": "MOTO SUSUKI", "placa": "86GWU7", "actividad": TXT_DESARROLLO_URBANO},
    14: {"solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "vehiculo": "MOTO SUSUKI", "placa": "86GWU8", "actividad": TXT_DESARROLLO_URBANO},
    15: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "87GWU8", "actividad": TXT_DESARROLLO_URBANO},
    16: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "88GWU7", "actividad": TXT_DESARROLLO_URBANO},
    17: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "88GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    18: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    19: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    20: {"solicita": "NOEL CHAN", "vehiculo": "MOTO DINAMO", "placa": "90GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    21: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "90GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    22: {"solicita": "LIAN", "vehiculo": "MOTO SUSUKI", "placa": "91GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    23: {"solicita": "RENAN/HELDER", "vehiculo": "AUTOMOVIL JETTA", "placa": "DFT565C", "actividad": TXT_DESARROLLO_URBANO},
    24: {"solicita": "QUEVEDO", "vehiculo": "CAMIONETA RAM 701", "placa": "CN2633B", "actividad": TXT_RAM_AMBIENTAL},
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
        "cesion_lian": {}
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

# --- CONSULTA Y ENVÍO A GOOGLE SHEETS ---
def obtener_datos_sheets(forzar=False):
    # Si ya tenemos datos en la sesión y no se forzó recarga, usarlos
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
                        "Operador_Sol": str(item.get("encargado_solicitado", "")).strip(),
                        "Importe_Sol": float(item.get("importe_solicitado", 0.0)) if item.get("importe_solicitado") else 0.0,
                        "Operador_Real": str(item.get("encargado_real", "")).strip(),
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
                "Operador_Sol": "",
                "Importe_Sol": 0.0,
                "Operador_Real": "",
                "Importe_Real": 0.0
            })
            
    df_res = pd.DataFrame(filas)
    st.session_state.df_datos_persistentes = df_res.copy()
    return df_res

def enviar_datos_sheets(registros_a_enviar, tipo="solicitado", f_elab=None, f_prog=None):
    payload = {"tipo": tipo, "registros": []}
    if f_elab:
        payload["fecha_elaboro"] = f_elab.strftime("%d/%m/%Y")
    if f_prog:
        payload["fecha_prog"] = f_prog.strftime("%d/%m/%Y")
        
    for _, fila in registros_a_enviar.iterrows():
        enc = fila["Operador_Sol"] if tipo == "solicitado" else fila.get("Operador_Real", fila["Operador_Sol"])
        imp = fila["Importe_Sol"] if tipo == "solicitado" else fila["Importe_Real"]
        payload["registros"].append({
            "row": int(fila["row"]),
            "encargado": str(enc).strip() if pd.notna(enc) else "",
            "importe": float(imp) if pd.notna(imp) else 0.0
        })
        
    # Actualizar inmediatamente la copia persistente en memoria para evitar desfases
    if "df_datos_persistentes" in st.session_state:
        df_mem = st.session_state.df_datos_persistentes
        for _, r_env in registros_a_enviar.iterrows():
            mask = df_mem["row"] == r_env["row"]
            if tipo == "solicitado":
                df_mem.loc[mask, "Operador_Sol"] = r_env["Operador_Sol"]
                df_mem.loc[mask, "Importe_Sol"] = r_env["Importe_Sol"]
            else:
                df_mem.loc[mask, "Operador_Real"] = r_env.get("Operador_Real", r_env["Operador_Sol"])
                df_mem.loc[mask, "Importe_Real"] = r_env["Importe_Real"]
        st.session_state.df_datos_persistentes = df_mem

    try:
        res = requests.post(WEBHOOK_URL, json=payload, timeout=10, allow_redirects=True)
        return res.status_code == 200
    except Exception:
        return False

# ==========================================
# GENERADORES DE ARCHIVOS OFICIALES
# ==========================================
def generar_excel_oficial_formato(df_datos, f_elab, f_prog):
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
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
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
        (2, "NOMBRE DEL\nENCARGADO"),
        (3, "VEHÍCULO"),
        (4, "PLACA"),
        (5, "OFICIAL /\nCOMODATO"),
        (6, "LITROS"),
        (7, "IMPORTE"),
        (8, "MAGNA / DIESEL"),
        (9, "ACTIVIDAD")
    ]
    
    for col_num, h_text in headers_oficiales:
        cell = ws.cell(row=11, column=col_num, value=h_text)
        cell.font = fuente_header
        cell.fill = fill_header_azul
        cell.border = border_cuadricula
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[11].height = 28

    for r_num in range(12, 25):
        row_info = df_datos[df_datos["row"] == r_num]
        
        if not row_info.empty:
            item = row_info.iloc[0]
            op_nombre = str(item["Operador_Sol"]).strip()
            veh = str(item["Vehículo"]).strip()
            plc = str(item["Placa"]).strip()
            imp_val = float(item["Importe_Sol"])
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

    c_lbl_tot = ws.cell(row=25, column=6, value="TOTAL")
    c_lbl_tot.font = fuente_bold
    c_lbl_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_lbl_tot.border = border_cuadricula
    
    c_val_tot = ws.cell(row=25, column=7, value="=SUM(G12:G24)")
    c_val_tot.font = fuente_bold
    c_val_tot.number_format = '$#,##0.00'
    c_val_tot.alignment = Alignment(horizontal="right", vertical="center")
    c_val_tot.border = border_cuadricula
    
    for c_rest in [2, 3, 4, 5, 8, 9]:
        ws.cell(row=25, column=c_rest).border = border_cuadricula

    anchos_cols = {'A': 3, 'B': 24, 'C': 16, 'D': 12, 'E': 14, 'F': 10, 'G': 14, 'H': 14, 'I': 52}
    for col_letra, ancho in anchos_cols.items():
        ws.column_dimensions[col_letra].width = ancho

    wb.save(output)
    return output.getvalue()

def generar_pdf_oficial(df_cargas, f_elab, f_prog):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 6, 'H. AYUNTAMIENTO DE CAMPECHE', 0, 1, 'C')
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(0, 5, 'DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE', 0, 1, 'C')
    pdf.ln(3)
    
    pdf.set_font('Helvetica', '', 9)
    pdf.cell(140, 5, 'Unidad: DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE', 0, 0, 'L')
    pdf.cell(130, 5, f'Elaboro: {f_elab.strftime("%d/%m/%Y")}', 0, 1, 'R')
    pdf.cell(140, 5, '', 0, 0, 'L')
    pdf.cell(130, 5, f'Programacion para el dia: {f_prog.strftime("%d/%m/%Y")}', 0, 1, 'R')
    pdf.ln(4)
    
    col_widths = (45, 32, 18, 18, 22, 16, 117)
    
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
            data_row.cell(str(row["Operador_Sol"]).strip())
            data_row.cell(str(row["Vehículo"]).strip())
            data_row.cell(str(row["Placa"]).strip())
            data_row.cell('OFICIAL')
            data_row.cell(f"${float(row['Importe_Sol']):,.2f}")
            data_row.cell('MAGNA')
            data_row.cell(str(row["Actividad"]).strip())
            total += float(row["Importe_Sol"])
            
        pdf.set_font('Helvetica', 'B', 8)
        tot_row = table.row()
        tot_row.cell('TOTAL AUTORIZADO:', colspan=4, align="RIGHT")
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
# 2. ENCABEZADO Y EVALUACIÓN DE HORARIO
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

c1, c2, c3 = st.columns([2.5, 1, 0.8])
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
    st.write("")
    if st.button("🔄 Sincronizar Sheets", use_container_width=True):
        obtener_datos_sheets(forzar=True)
        st.toast("Datos sincronizados con Google Sheets.", icon="✅")
        st.rerun()

with c3:
    st.write("")
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
# 3. VISTA SOLICITANTE (CON LISTA DE OPERADORES)
# ==========================================
if not es_admin:
    presupuesto_propio = presupuestos_actuales.get(usuario_efectivo, 0.00)
    df_solicitante = df_actual[df_actual["Solicitante"] == usuario_efectivo].copy()
    
    lista_operadores_autorizados = OPERADORES_POR_SOLICITANTE.get(usuario_efectivo, [])
    
    if sistema_bloqueado:
        st.error("🔒 **SISTEMA CERRADO POR HORARIO (3:10 PM)**. La captura semanal ha finalizado.")
    else:
        st.caption("📱 Selecciona quién conducirá y el importe de cada vehículo para esta semana:")
    
    with st.form("form_solicitante_movil"):
        nuevos_valores = []
        
        for idx, row in df_solicitante.iterrows():
            with st.container(border=True):
                st.markdown(f"🛵 **{row['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row['Placa']}`**")
                st.caption(f"📋 **Actividad:** {row['Actividad']}")
                
                c_op, c_imp = st.columns([1.5, 1])
                
                opciones_operadores = [""] + lista_operadores_autorizados
                val_actual = str(row["Operador_Sol"]).strip()
                
                if val_actual and val_actual not in opciones_operadores:
                    opciones_operadores.append(val_actual)
                    
                idx_sel = opciones_operadores.index(val_actual) if val_actual in opciones_operadores else 0
                
                with c_op:
                    val_encargado = st.selectbox(
                        "Operador / Conductor",
                        options=opciones_operadores,
                        index=idx_sel,
                        key=f"op_{row['row']}",
                        disabled=sistema_bloqueado
                    )
                    
                with c_imp:
                    val_importe = st.number_input(
                        "Monto ($)",
                        value=float(row["Importe_Sol"]),
                        step=50.0,
                        min_value=0.0,
                        key=f"imp_{row['row']}",
                        disabled=sistema_bloqueado,
                        format="%.2f"
                    )
                
                nuevos_valores.append({
                    "row": row["row"],
                    "Solicitante": row["Solicitante"],
                    "Vehículo": row["Vehículo"],
                    "Placa": row["Placa"],
                    "Actividad": row["Actividad"],
                    "Operador_Sol": val_encargado,
                    "Importe_Sol": val_importe,
                    "Operador_Real": row["Operador_Real"],
                    "Importe_Real": row["Importe_Real"]
                })
        
        df_edit_movil = pd.DataFrame(nuevos_valores)
        total_capturado = df_edit_movil["Importe_Sol"].sum()
        saldo_restante = presupuesto_propio - total_capturado
        
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("Presupuesto Asignado", f"${presupuesto_propio:,.2f}")
        m2.metric("Total a Cargar", f"${total_capturado:,.2f}")
        m3.metric("Saldo Disponible", f"${saldo_restante:,.2f}", delta_color="normal" if saldo_restante >= 0 else "off")
        
        if total_capturado > presupuesto_propio:
            st.error(f"⚠️ Excedes tu presupuesto autorizado por **${abs(saldo_restante):,.2f} MXN**.")
            
        btn_guardar = st.form_submit_button("💾 Guardar Solicitud", type="primary", use_container_width=True, disabled=sistema_bloqueado)
        
        if btn_guardar:
            if total_capturado > presupuesto_propio:
                st.error("No puedes guardar si excedes el presupuesto autorizado.")
            else:
                with st.spinner("Guardando en Google Sheets..."):
                    # 👉 SOLO SE ENVÍAN LAS FILAS DE ESTE SOLICITANTE (SIN AFECTAR A LOS DEMÁS)
                    enviar_datos_sheets(df_edit_movil, tipo="solicitado")
                    st.success("✅ ¡Solicitud guardada con éxito en Google Sheets!")
                    st.rerun()

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Panel de Consolidación, Edición y Descarga Oficial")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración", value=date.today())
    with col_f2:
        f_prog = st.date_input("Programación para el día", value=date.today())

    tab_saldos, tab_solicitud_final, tab_mi_carga, tab_auditoria, tab_mantenimiento = st.tabs([
        "📊 Monitoreo y Asignación de Presupuestos",
        "✏️ Solicitud Final y Modificación",
        "🛵 Mi Carga (LIAN)",
        "✅ Auditoría y Carga Real",
        "🛠️ Modo Pruebas"
    ])

    # 1. MONITOREO Y TRANSFERENCIA DE PRESUPUESTOS
    with tab_saldos:
        st.markdown("##### 💵 Balance de Presupuestos en Tiempo Real")
        
        total_global_sol = df_actual["Importe_Sol"].sum()
        saldo_global_sol = PRESUPUESTO_GLOBAL - total_global_sol
        
        asig_comodin_dict = cfg_actual.get("asignacion_comodin", {})
        total_comodin_usado = sum(asig_comodin_dict.values())
        comodin_disponible = max(0.0, BOLSA_COMODIN_TOTAL - total_comodin_usado)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Monto Base Áreas", f"${sum(PRESUPUESTO_BASE_POR_SOLICITANTE.values()):,.2f}")
        c2.metric("Bolsa Comodín Libre", f"${comodin_disponible:,.2f}", delta=f"${total_comodin_usado:,.2f} asignados" if total_comodin_usado > 0 else "Disponible")
        c3.metric("Presupuesto Global", f"${PRESUPUESTO_GLOBAL:,.2f}")
        c4.metric(
            "Total Ya Solicitado", 
            f"${total_global_sol:,.2f}", 
            delta=f"${saldo_global_sol:,.2f} disponible",
            delta_color="normal" if saldo_global_sol >= 0 else "inverse"
        )
        
        st.write("")
        st.markdown("##### 📋 Resumen Financiero por Solicitante")
        filas_saldos_area = []
        for sol, p_efectivo in presupuestos_actuales.items():
            sub_df = df_actual[df_actual["Solicitante"] == sol]
            sol_monto = sub_df["Importe_Sol"].sum()
            disp_monto = p_efectivo - sol_monto
            pct_ejercido = (sol_monto / p_efectivo * 100) if p_efectivo > 0 else 0
            
            if disp_monto == 0:
                estatus = "✅ 100% Ejercido"
            elif disp_monto < 0:
                estatus = "⚠️ Excedido"
            elif sol_monto > 0:
                estatus = "🟢 Con Saldo"
            else:
                estatus = "⚪ Sin Carga"
                
            filas_saldos_area.append({
                "Solicitante / Área": sol,
                "Presupuesto Base": PRESUPUESTO_BASE_POR_SOLICITANTE.get(sol, 0.0),
                "Presupuesto Autorizado": p_efectivo,
                "Monto Solicitado": sol_monto,
                "Saldo Disponible": disp_monto,
                "% Usado": f"{pct_ejercido:.1f}%",
                "Estatus": estatus
            })
            
        df_saldos_area = pd.DataFrame(filas_saldos_area)
        st.dataframe(
            df_saldos_area,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Presupuesto Base": st.column_config.NumberColumn(format="$%.2f"),
                "Presupuesto Autorizado": st.column_config.NumberColumn(format="$%.2f"),
                "Monto Solicitado": st.column_config.NumberColumn(format="$%.2f"),
                "Saldo Disponible": st.column_config.NumberColumn(format="$%.2f"),
            }
        )

        st.divider()
        st.markdown("##### 🔀 Asignar Comodín y Ceder Presupuesto Propio")
        
        col_trans1, col_trans2 = st.columns(2)
        
        with col_trans1:
            with st.container(border=True):
                st.markdown(f"🎁 **Asignar Bolsa Comodín (Libre: ${comodin_disponible:,.2f})**")
                areas_comodin = [u for u in PRESUPUESTO_BASE_POR_SOLICITANTE.keys() if u != "LIAN"]
                destinatario_comodin = st.selectbox("Asignar Comodín a:", areas_comodin, key="sel_comodin")
                monto_comodin_add = st.number_input("Monto Extra a Asignar ($)", min_value=0.0, max_value=float(comodin_disponible), step=50.0, value=0.0, key="inp_comodin")
                
                if st.button("➕ Asignar Extra de Comodín", use_container_width=True):
                    if monto_comodin_add > 0:
                        asig_act = cfg_actual.get("asignacion_comodin", {})
                        asig_act[destinatario_comodin] = asig_act.get(destinatario_comodin, 0.0) + monto_comodin_add
                        cfg_actual["asignacion_comodin"] = asig_act
                        guardar_config(cfg_actual)
                        st.toast(f"Se asignaron ${monto_comodin_add:,.2f} a {destinatario_comodin}", icon="🎁")
                        st.rerun()

        with col_trans2:
            with st.container(border=True):
                presupuesto_lian_actual = presupuestos_actuales["LIAN"]
                st.markdown(f"🤝 **Ceder Presupuesto de LIAN (Disponible: ${presupuesto_lian_actual:,.2f})**")
                areas_ceder = [u for u in PRESUPUESTO_BASE_POR_SOLICITANTE.keys() if u != "LIAN"]
                destinatario_ceder = st.selectbox("Ceder mi presupuesto a:", areas_ceder, key="sel_ceder")
                monto_ceder = st.number_input("Monto a Ceder ($)", min_value=0.0, max_value=float(presupuesto_lian_actual), step=50.0, value=0.0, key="inp_ceder")
                
                if st.button("Transferir Presupuesto", use_container_width=True):
                    if monto_ceder > 0:
                        ces_act = cfg_actual.get("cesion_lian", {})
                        ces_act[destinatario_ceder] = ces_act.get(destinatario_ceder, 0.0) + monto_ceder
                        cfg_actual["cesion_lian"] = ces_act
                        guardar_config(cfg_actual)
                        st.toast(f"Has transferido ${monto_ceder:,.2f} a {destinatario_ceder}", icon="🤝")
                        st.rerun()

        if cfg_actual.get("asignacion_comodin") or cfg_actual.get("cesion_lian"):
            if st.button("🔄 Restablecer Transferencias a Valores Originales", type="secondary", use_container_width=True):
                cfg_actual["asignacion_comodin"] = {}
                cfg_actual["cesion_lian"] = {}
                guardar_config(cfg_actual)
                st.toast("Transferencias restablecidas a presupuestos base.", icon="🔄")
                st.rerun()

    # 2. SOLICITUD FINAL Y EDITOR CON LISTA DESPLEGABLE PARA EL ADMINISTRADOR
    with tab_solicitud_final:
        st.markdown("##### 🚗 Solicitud de Carga Oficial y Editor Administrativo")
        st.caption("Como Administrador, selecciona el operador de la lista y modifica los montos antes de descargar:")
        
        todos_los_operadores = [""]
        for l_ops in OPERADORES_POR_SOLICITANTE.values():
            for o in l_ops:
                if o not in todos_los_operadores:
                    todos_los_operadores.append(o)
                    
        for op_existente in df_actual["Operador_Sol"].dropna().unique():
            op_limpio = str(op_existente).strip()
            if op_limpio and op_limpio not in todos_los_operadores:
                todos_los_operadores.append(op_limpio)
        
        df_admin_edit = st.data_editor(
            df_actual.copy(),
            use_container_width=True,
            disabled=["row", "Solicitante", "Vehículo", "Placa", "Actividad", "Operador_Real", "Importe_Real"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Área"),
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placa"),
                "Operador_Sol": st.column_config.SelectboxColumn(
                    "Operador / Encargado (Elegir)",
                    options=todos_los_operadores,
                    required=False,
                    width="medium"
                ),
                "Importe_Sol": st.column_config.NumberColumn("Importe Solicitado ($) (Editable)", min_value=0.0, step=50.0, format="$%.2f"),
                "Actividad": st.column_config.TextColumn("Actividad", width="medium"),
                "row": None, "Operador_Real": None, "Importe_Real": None
            },
            hide_index=True,
            key="admin_solicitudes_editor"
        )
        
        if st.button("💾 Guardar Cambios Realizados por Admin en Google Sheets", type="primary", use_container_width=True):
            with st.spinner("Guardando modificaciones en Google Sheets..."):
                enviar_datos_sheets(df_admin_edit, tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                st.success("✅ ¡Cambios administrativos guardados y sincronizados!")
                st.rerun()
        
        st.markdown("---")
        
        df_solo_cargas_sol = df_admin_edit[df_admin_edit["Importe_Sol"] > 0].copy()
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            excel_bytes = generar_excel_oficial_formato(df_admin_edit, f_elab, f_prog)
            st.download_button(
                label="📥 Descargar Formato Oficial Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"SOLICITUD_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with col_d2:
            pdf_bytes = generar_pdf_oficial(df_solo_cargas_sol, f_elab, f_prog)
            st.download_button(
                label="📄 Descargar Oficio Oficial en PDF",
                data=pdf_bytes,
                file_name=f"OFICIO_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    # 3. MI CARGA (LIAN)
    with tab_mi_carga:
        st.markdown("##### 🛵 Registro de Solicitud para tu Unidad")
        df_lian = df_actual[df_actual["Solicitante"] == "LIAN"].copy()
        presupuesto_lian_efectivo = presupuestos_actuales["LIAN"]
        
        if not df_lian.empty:
            row_lian = df_lian.iloc[0]
            val_guardado = float(row_lian["Importe_Sol"])
            op_guardado = str(row_lian["Operador_Sol"]).strip()
            
            with st.container(border=True):
                st.markdown(f"🛵 **{row_lian['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row_lian['Placa']}`** &nbsp;|&nbsp; Presupuesto Disponible: **${presupuesto_lian_efectivo:,.2f}**")
                st.caption(f"📋 **Actividad:** {row_lian['Actividad']}")
                
                c_op_lian, c_imp_lian = st.columns([1.5, 1])
                with c_op_lian:
                    opciones_lian = ["", "FRANCISCO ALONZO"]
                    idx_op_lian = opciones_lian.index(op_guardado) if op_guardado in opciones_lian else 0
                    val_op_lian = st.selectbox(
                        "Operador Encargado",
                        options=opciones_lian,
                        index=idx_op_lian,
                        key="admin_op_lian_tab"
                    )
                with c_imp_lian:
                    val_imp_lian = st.number_input(
                        "Importe Solicitado ($)", 
                        value=val_guardado,
                        step=50.0,
                        min_value=0.0,
                        max_value=float(presupuesto_lian_efectivo),
                        key="admin_imp_lian_tab",
                        format="%.2f"
                    )
                
                if st.button("💾 Guardar Mi Carga", type="primary", use_container_width=True):
                    # Guardar solo la fila de LIAN
                    df_mi_carga = df_actual[df_actual["Solicitante"] == "LIAN"].copy()
                    df_mi_carga["Operador_Sol"] = val_op_lian
                    df_mi_carga["Importe_Sol"] = val_imp_lian
                    enviar_datos_sheets(df_mi_carga, tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                    st.success("✅ Tu carga fue registrada correctamente.")
                    st.rerun()

    # 4. AUDITORÍA Y CARGA REAL
    with tab_auditoria:
        st.markdown("##### 🔍 Conciliación y Registro de Cargas Reales Comprobadas")
        st.caption("Captura el monto realmente cargado para calcular diferencias:")
        
        df_real_edit = st.data_editor(
            df_actual.copy(),
            use_container_width=True,
            disabled=["row", "Solicitante", "Vehículo", "Placa", "Actividad", "Importe_Sol"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Área"),
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placas"),
                "Importe_Sol": st.column_config.NumberColumn("Solicitado ($)", format="$%.2f"),
                "Importe_Real": st.column_config.NumberColumn("Monto Realmente Cargado ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "row": None, "Actividad": None, "Operador_Sol": None, "Operador_Real": None
            },
            hide_index=True,
            key="editor_seccion_real_apartado"
        )
        
        total_global_sol = df_actual["Importe_Sol"].sum()
        total_global_real = df_real_edit["Importe_Real"].sum()
        ahorro_vs_sol = total_global_sol - total_global_real
        saldo_global_disponible = PRESUPUESTO_GLOBAL - total_global_real
        
        st.divider()
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Total Solicitado", f"${total_global_sol:,.2f}")
        r2.metric("Total Real Ejercido", f"${total_global_real:,.2f}")
        r3.metric("Ahorro vs Solicitado", f"${ahorro_vs_sol:,.2f}", delta_color="normal")
        r4.metric("Saldo Disponible Restante", f"${saldo_global_disponible:,.2f}", delta_color="normal")
        
        st.write("")
        if st.button("💾 Guardar Cargas Reales en Sheets", type="primary", use_container_width=True):
            enviar_datos_sheets(df_real_edit, tipo="real", f_elab=f_elab, f_prog=f_prog)
            st.success("✅ Cargas reales sincronizadas correctamente en Google Sheets.")
            st.rerun()

    # 5. MODO PRUEBAS Y MANTENIMIENTO
    with tab_mantenimiento:
        st.markdown("##### 🛠️ Control de Horarios, Simulación y Limpieza")
        
        with st.container(border=True):
            st.subheader("⏰ Control de Bloqueo a las 3:10 PM")
            st.write(
                "Por regla general, **todos los solicitantes quedan bloqueados automáticamente a las 3:10 PM** (tú como Administrador siempre tienes acceso)."
            )
            
            estado_desbloqueo = cfg_actual.get("desbloqueo_horario", False)
            toggle_horario = st.toggle(
                "⚡ Desbloquear a todos los usuarios (Permitir captura 24/7 fuera de las 3:10 PM)", 
                value=estado_desbloqueo
            )
            
            if toggle_horario != estado_desbloqueo:
                cfg_actual["desbloqueo_horario"] = toggle_horario
                guardar_config(cfg_actual)
                if toggle_horario:
                    st.toast("Captura 24/7 HABILITADA para todos.", icon="🟢")
                else:
                    st.toast("Bloqueo de 3:10 PM ACTIVADO para usuarios.", icon="🔒")
                st.rerun()

        with st.container(border=True):
            st.subheader("🧹 Reiniciar / Limpiar Todas las Cargas a $0.00")
            st.write("Esta acción borra los nombres y regresa a **$0.00** los importes tanto de la sección solicitada como de la real en Google Sheets.")
            
            if st.button("🗑️ Limpiar Todo y Restablecer a $0.00", type="secondary", use_container_width=True):
                df_limpio = df_actual.copy()
                df_limpio["Operador_Sol"] = ""
                df_limpio["Importe_Sol"] = 0.0
                df_limpio["Operador_Real"] = ""
                df_limpio["Importe_Real"] = 0.0
                
                enviar_datos_sheets(df_limpio, tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                enviar_datos_sheets(df_limpio, tipo="real", f_elab=f_elab, f_prog=f_prog)
                st.success("✅ Todas las unidades restablecidas a $0.00.")
                st.rerun()

        with st.container(border=True):
            st.subheader("🧪 Probar Vista Móvil de Solicitante")
            usuarios_para_test = [u for u in USUARIOS_PASSWORD.keys() if u != "LIAN"]
            solicitante_a_testear = st.selectbox("Selecciona al solicitante a simular:", usuarios_para_test)
            
            if st.button("👁️ Entrar a Modo Simulación", type="secondary", use_container_width=True):
                st.session_state.vista_simulada = solicitante_a_testear
                st.rerun()
