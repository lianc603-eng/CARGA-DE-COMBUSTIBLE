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
# CONFIGURACIONES GENERALES
# ==========================================================
PRESUPUESTO_GLOBAL = 3800.00
HORA_LIMITE = time(15, 0)  # 3:00 PM
ZONA_HORARIA = pytz.timezone("America/Merida")

HABILITAR_CAPTURA_24_7 = True  

CONFIG_FILE = "config_sistema.json"
HISTORICO_FILE = "historico_cargas.json"

WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzOjgha2Zjyog01t6LmA_R--EB4Ecqv2ifO_i2YJbLRLbXGShbu5uzFVi85FUTGplM8/exec"

PRESUPUESTO_POR_SOLICITANTE = {
    "COB CHAVEZ NARCISO DEL JESUS": 200.00,
    "PEREZ MAZIN CARLOS EDUARDO": 200.00,
    "DE LA CRUZ PEREZ WILLIAN ARLEY": 200.00,
    "NOEL CHAN": 850.00,
    "LIAN": 150.00,
    "QUEVEDO": 1500.00,
    "RENAN/HELDER": 500.00,
}

TXT_DESARROLLO_URBANO = "LLEVAR A CABO ACTIVIDADES DE INSPECCIONES, VERIFICACIONES Y SUPERVICIONES DE OBRAS Y OBSTRUCCIONES A LA VIA PÚBLICA CORRESPONDIENTES A LA SUBDIRECCION DE DESARROLLO URBANO"
TXT_MEDIO_AMBIENTE = "PARA LLEVAR A CABO INSPECCIONES A CARGO DE LA SUBDIRECCION DE MEDIO AMBIENTE, COMO LO SON ATENDER REPORTES POR TIRADERO DE AGUAS JABONOSAS, MALTRATO ANIMAL Y CONTAMINACION AUDITIVA, ASI COMO DIVERSOS TIPOS DE CONTAMINACION"
TXT_RAM_AMBIENTAL = "PARA LLEVAR A CABO ACTIVIDADES DE ESTERILIZACIONES DE PERROS Y GATOS, RECOLECCION DE MERMA DE FRUTAS Y VERDURAS EN SUPERMERCADOS Y REFORESTACIONES"

MAPEO_SOLICITANTES = {
    12: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "85GWU7", "actividad": TXT_DESARROLLO_URBANO},
    13: {"solicita": "PEREZ MAZIN CARLOS EDUARDO", "vehiculo": "MOTO SUSUKI", "placa": "86GWU7", "actividad": TXT_DESARROLLO_URBANO},
    14: {"solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "vehiculo": "MOTO SUSUKI", "placa": "86GWU8", "actividad": TXT_DESARROLLO_URBANO},
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
    return {"desbloqueo_horario": True}

def guardar_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f)
    except Exception:
        pass

def leer_historico():
    if os.path.exists(HISTORICO_FILE):
        try:
            with open(HISTORICO_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def guardar_historico(hist):
    try:
        with open(HISTORICO_FILE, "w") as f:
            json.dump(hist, f)
    except Exception:
        pass

# --- CONSULTA Y ENVÍO A GOOGLE SHEETS ---
def obtener_datos_sheets(forzar=False):
    if "df_datos" in st.session_state and not forzar:
        return st.session_state.df_datos

    try:
        res = requests.get(WEBHOOK_URL, timeout=8, allow_redirects=True)
        if res.status_code == 200:
            datos_raw = res.json().get("data", [])
            filas = []
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
            if filas:
                df = pd.DataFrame(filas)
                st.session_state.df_datos = df
                return df
    except Exception:
        pass

    if "df_datos" in st.session_state:
        return st.session_state.df_datos

    filas_base = []
    for r, info in MAPEO_SOLICITANTES.items():
        filas_base.append({
            "row": r,
            "Solicitante": info["solicita"],
            "Vehículo": info["vehiculo"],
            "Placa": info["placa"],
            "Actividad": info["actividad"],
            "Operador_Sol": "FRANCISCO ALONZO" if info["solicita"] == "LIAN" else "",
            "Importe_Sol": 150.0 if info["solicita"] == "LIAN" else 0.0,
            "Operador_Real": "",
            "Importe_Real": 0.0
        })
    df_def = pd.DataFrame(filas_base)
    st.session_state.df_datos = df_def
    return df_def

def enviar_datos_sheets(registros, tipo="solicitado", f_elab=None, f_prog=None):
    payload = {"tipo": tipo, "registros": []}
    if f_elab:
        payload["fecha_elaboro"] = f_elab.strftime("%d/%m/%Y")
    if f_prog:
        payload["fecha_prog"] = f_prog.strftime("%d/%m/%Y")
        
    for _, fila in registros.iterrows():
        enc = fila["Operador_Sol"] if tipo == "solicitado" else fila.get("Operador_Real", fila["Operador_Sol"])
        imp = fila["Importe_Sol"] if tipo == "solicitado" else fila["Importe_Real"]
        payload["registros"].append({
            "row": int(fila["row"]),
            "encargado": str(enc).strip() if pd.notna(enc) else "",
            "importe": float(imp) if pd.notna(imp) else 0.0
        })
        
    st.session_state.df_datos = registros.copy()
    try:
        res = requests.post(WEBHOOK_URL, json=payload, timeout=10, allow_redirects=True)
        return res.status_code == 200
    except Exception:
        return False

# ==========================================
# GENERADORES DE ARCHIVOS OFICIALES
# ==========================================
def generar_excel_completo(df_datos, f_elab, f_prog):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    fuente_titulo = Font(name="Calibri", size=12, bold=True, color="1F497D")
    fuente_sub = Font(name="Calibri", size=10, bold=True)
    fuente_header_sol = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fuente_header_real = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fuente_bold = Font(name="Calibri", size=9, bold=True)
    fuente_normal = Font(name="Calibri", size=9)
    
    fill_sol = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_real = PatternFill(start_color="1E4D2B", end_color="1E4D2B", fill_type="solid")
    fill_total = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
    
    border_fino = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )

    ws = wb.active
    ws.title = "CONTROL COMBUSTIBLE"
    ws.views.sheetView[0].showGridLines = True
    
    ws["B2"] = "H. AYUNTAMIENTO DE CAMPECHE"
    ws["B2"].font = fuente_titulo
    ws["B3"] = "DIRECCIÓN DE DESARROLLO URBANO Y MEDIO AMBIENTE"
    ws["B3"].font = fuente_sub
    
    ws["F5"] = f"Elaboró: {f_elab.strftime('%d/%m/%Y')}"
    ws["F6"] = f"Programación: {f_prog.strftime('%d/%m/%Y')}"
    ws["F5"].font = fuente_sub
    ws["F6"].font = fuente_sub

    headers = [
        "", "ENCARGADO (SOLICITADO)", "VEHÍCULO", "PLACA", "RÉGIMEN", "IMPORTE SOL. ($)", "TIPO", "ACTIVIDAD",
        "", "VEHÍCULO", "PLACA", "IMPORTE REAL ($)", "DIFERENCIA ($)", "% EJERCIDO"
    ]
    ws.append([])
    ws.append(headers)
    
    for c_idx in range(2, 9):
        cell = ws.cell(row=8, column=c_idx)
        cell.font = fuente_header_sol
        cell.fill = fill_sol
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for c_idx in range(10, 15):
        cell = ws.cell(row=8, column=c_idx)
        cell.font = fuente_header_real
        cell.fill = fill_real
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    start_row = 9
    for i, row in df_datos.iterrows():
        curr_row = start_row + i
        f_dif = f"=F{curr_row}-L{curr_row}"
        f_pct = f"=IF(F{curr_row}>0, L{curr_row}/F{curr_row}, 0)"
        
        ws.append([
            "",
            row["Operador_Sol"],
            row["Vehículo"],
            row["Placa"],
            "OFICIAL",
            float(row["Importe_Sol"]),
            "MAGNA",
            row["Actividad"],
            "",
            row["Vehículo"],
            row["Placa"],
            float(row["Importe_Real"]),
            f_dif,
            f_pct
        ])
        
        for c in range(2, 15):
            if c == 9:
                continue
            celda = ws.cell(row=curr_row, column=c)
            celda.font = fuente_normal
            celda.border = border_fino
            if c in [6, 12, 13]:
                celda.number_format = '$#,##0.00'
                celda.alignment = Alignment(horizontal="right")
            elif c == 14:
                celda.number_format = '0.0%'
                celda.alignment = Alignment(horizontal="center")
            elif c in [4, 5, 7, 11]:
                celda.alignment = Alignment(horizontal="center")
                
    end_row = start_row + len(df_datos) - 1
    total_row = end_row + 1
    
    ws.cell(row=total_row, column=5, value="TOTAL SOLICITADO:").font = fuente_bold
    ws.cell(row=total_row, column=5).alignment = Alignment(horizontal="right")
    c_tot_sol = ws.cell(row=total_row, column=6, value=f"=SUM(F{start_row}:F{end_row})")
    c_tot_sol.font = fuente_bold
    c_tot_sol.number_format = '$#,##0.00'
    c_tot_sol.fill = fill_total
    
    ws.cell(row=total_row, column=11, value="TOTAL REAL:").font = fuente_bold
    ws.cell(row=total_row, column=11).alignment = Alignment(horizontal="right")
    c_tot_real = ws.cell(row=total_row, column=12, value=f"=SUM(L{start_row}:L{end_row})")
    c_tot_real.font = fuente_bold
    c_tot_real.number_format = '$#,##0.00'
    c_tot_real.fill = fill_total
    
    c_tot_dif = ws.cell(row=total_row, column=13, value=f"=F{total_row}-L{total_row}")
    c_tot_dif.font = fuente_bold
    c_tot_dif.number_format = '$#,##0.00'
    c_tot_dif.fill = fill_total
    
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        if col_letter in ['A', 'I']:
            ws.column_dimensions[col_letter].width = 3
        else:
            ws.column_dimensions[col_letter].width = 22
            
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
# 2. ENCABEZADO
# ==========================================
usuario_real = st.session_state.usuario_logueado
es_admin_real = (usuario_real == "LIAN")
usuario_efectivo = st.session_state.vista_simulada if (es_admin_real and st.session_state.vista_simulada) else usuario_real
es_admin = (usuario_efectivo == "LIAN")

ahora_local = datetime.now(ZONA_HORARIA)
hora_actual = ahora_local.time()

c1, c2, c3 = st.columns([2.5, 1, 0.8])
with c1:
    st.title("⛽ Control de Combustible")
    if es_admin_real and st.session_state.vista_simulada:
        st.warning(f"🧪 **MODO DE PRUEBA ACTIVO**: Simulando vista como **{usuario_efectivo}**")
    else:
        st.markdown("👑 **ADMINISTRADOR GENERAL**" if es_admin else f"👤 Solicitante: **{usuario_efectivo}**")
    st.caption(f"🕒 {ahora_local.strftime('%I:%M %p')} | 🟢 Horario Abierto para Captura")

with c2:
    st.write("")
    if st.button("🔄 Sincronizar Sheets", use_container_width=True):
        st.session_state.df_datos = obtener_datos_sheets(forzar=True)
        st.toast("Datos actualizados desde Google Sheets.", icon="✅")
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
# 3. VISTA SOLICITANTE (MÓVIL)
# ==========================================
if not es_admin:
    presupuesto_propio = PRESUPUESTO_POR_SOLICITANTE.get(usuario_efectivo, 0.00)
    df_solicitante = df_actual[df_actual["Solicitante"] == usuario_efectivo].copy()
    
    st.caption("📱 Llena los datos de los vehículos que cargarán esta semana y presiona **Guardar Solicitud**:")
    
    with st.form("form_solicitante_movil"):
        nuevos_valores = []
        
        for idx, row in df_solicitante.iterrows():
            with st.container(border=True):
                st.markdown(f"🛵 **{row['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row['Placa']}`**")
                st.caption(f"📋 **Actividad:** {row['Actividad']}")
                
                c_op, c_imp = st.columns([1.5, 1])
                
                with c_op:
                    val_encargado = st.text_input(
                        "Nombre del Conductor / Operador",
                        value=row["Operador_Sol"],
                        key=f"op_{row['row']}",
                        placeholder="Ej. Juan Pérez"
                    )
                with c_imp:
                    val_importe = st.number_input(
                        "Monto ($)",
                        value=float(row["Importe_Sol"]),
                        step=50.0,
                        min_value=0.0,
                        key=f"imp_{row['row']}",
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
        m1.metric("Presupuesto", f"${presupuesto_propio:,.2f}")
        m2.metric("Total a Cargar", f"${total_capturado:,.2f}")
        m3.metric("Saldo Disponible", f"${saldo_restante:,.2f}", delta_color="normal" if saldo_restante >= 0 else "off")
        
        if total_capturado > presupuesto_propio:
            st.error(f"⚠️ Excedes tu presupuesto por **${abs(saldo_restante):,.2f} MXN**.")
            
        btn_guardar = st.form_submit_button("💾 Guardar Solicitud", type="primary", use_container_width=True)
        
        if btn_guardar:
            if total_capturado > presupuesto_propio:
                st.error("No puedes guardar si excedes el presupuesto autorizado.")
            else:
                with st.spinner("Guardando en Google Sheets..."):
                    for _, r_m in df_edit_movil.iterrows():
                        mask = df_actual["row"] == r_m["row"]
                        df_actual.loc[mask, "Operador_Sol"] = r_m["Operador_Sol"]
                        df_actual.loc[mask, "Importe_Sol"] = r_m["Importe_Sol"]
                    enviar_datos_sheets(df_actual, tipo="solicitado")
                    st.success("✅ ¡Solicitud guardada con éxito en Google Sheets!")
                    st.rerun()

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Panel de Consolidación y Descarga Oficial")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración", value=date.today())
    with col_f2:
        f_prog = st.date_input("Programación para el día", value=date.today())

    tab_saldos, tab_solicitud_final, tab_mi_carga, tab_auditoria, tab_historico, tab_mantenimiento = st.tabs([
        "📊 Monitoreo de Saldos por Área",
        "📄 Solicitud Final (Solo Vehículos que Cargan)",
        "🛵 Mi Carga (LIAN)",
        "✅ Auditoría y Carga Real",
        "📁 Histórico de Cargas",
        "🛠️ Modo Pruebas"
    ])

    df_solo_cargas_sol = df_actual[df_actual["Importe_Sol"] > 0].copy()

    # 1. MONITOREO DE SALDOS
    with tab_saldos:
        st.markdown("##### 💵 Balance de Presupuestos en Tiempo Real")
        total_global_sol = df_actual["Importe_Sol"].sum()
        saldo_global_sol = PRESUPUESTO_GLOBAL - total_global_sol
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Presupuesto Global", f"${PRESUPUESTO_GLOBAL:,.2f}")
        c2.metric("Total Ya Solicitado", f"${total_global_sol:,.2f}", delta=f"{total_global_sol - PRESUPUESTO_GLOBAL:,.2f}", delta_color="inverse")
        c3.metric("Saldo Disponible Global", f"${saldo_global_sol:,.2f}", delta_color="normal" if saldo_global_sol >= 0 else "off")
        c4.metric("Bolsa Comodín", "$200.00")
        
        st.write("")
        filas_saldos_area = []
        for sol, p_base in PRESUPUESTO_POR_SOLICITANTE.items():
            sub_df = df_actual[df_actual["Solicitante"] == sol]
            sol_monto = sub_df["Importe_Sol"].sum()
            disp_monto = p_base - sol_monto
            pct_ejercido = (sol_monto / p_base * 100) if p_base > 0 else 0
            
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
                "Presupuesto Base": p_base,
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
                "Monto Solicitado": st.column_config.NumberColumn(format="$%.2f"),
                "Saldo Disponible": st.column_config.NumberColumn(format="$%.2f"),
            }
        )

    # 2. SOLICITUD FINAL
    with tab_solicitud_final:
        st.markdown("##### 🚗 Lista Oficial de Unidades a Cargar")
        
        if df_solo_cargas_sol.empty:
            st.warning("⚠️ No hay vehículos con monto asignado para generar el reporte.")
        else:
            st.dataframe(
                df_solo_cargas_sol[["Operador_Sol", "Vehículo", "Placa", "Importe_Sol", "Actividad"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Operador_Sol": st.column_config.TextColumn("Operador / Encargado"),
                    "Vehículo": st.column_config.TextColumn("Vehículo"),
                    "Placa": st.column_config.TextColumn("Placa"),
                    "Importe_Sol": st.column_config.NumberColumn("Importe ($)", format="$%.2f"),
                    "Actividad": st.column_config.TextColumn("Actividad", width="large")
                }
            )
            
            st.markdown("---")
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                excel_bytes = generar_excel_completo(df_actual, f_elab, f_prog)
                st.download_button(
                    label="📥 Descargar Reporte en Excel (.xlsx)",
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
        
        if not df_lian.empty:
            row_lian = df_lian.iloc[0]
            with st.container(border=True):
                st.markdown(f"🛵 **{row_lian['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row_lian['Placa']}`** &nbsp;|&nbsp; Presupuesto Base: **$150.00**")
                st.caption(f"📋 **Actividad:** {row_lian['Actividad']}")
                
                c_op_lian, c_imp_lian = st.columns([1.5, 1])
                with c_op_lian:
                    val_op_lian = st.text_input(
                        "Nombre del Conductor / Operador", 
                        value=row_lian["Operador_Sol"] if row_lian["Operador_Sol"] else "FRANCISCO ALONZO",
                        key="admin_op_lian_tab"
                    )
                with c_imp_lian:
                    val_imp_lian = st.number_input(
                        "Importe Solicitado ($)", 
                        value=float(row_lian["Importe_Sol"]) if row_lian["Importe_Sol"] > 0 else 150.00,
                        step=50.0,
                        min_value=0.0,
                        key="admin_imp_lian_tab",
                        format="%.2f"
                    )
                
                if st.button("💾 Guardar Mi Carga", type="primary", use_container_width=True):
                    mask_lian = df_actual["Solicitante"] == "LIAN"
                    df_actual.loc[mask_lian, "Operador_Sol"] = val_op_lian
                    df_actual.loc[mask_lian, "Importe_Sol"] = val_imp_lian
                    enviar_datos_sheets(df_actual, tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                    st.success("✅ Tu carga fue registrada correctamente.")
                    st.rerun()

    # 4. AUDITORÍA Y CARGA REAL
    with tab_auditoria:
        st.markdown("##### 🔍 Conciliación y Registro de Cargas Reales Comprobadas")
        st.caption("Captura el monto realmente cargado para calcular diferencias y remanentes:")
        
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
        btn_c1, btn_c2 = st.columns(2)
        with btn_c1:
            if st.button("💾 Guardar Cargas Reales en Sheets", type="primary", use_container_width=True):
                enviar_datos_sheets(df_real_edit, tipo="real", f_elab=f_elab, f_prog=f_prog)
                st.success("✅ Cargas reales sincronizadas correctamente.")
                st.rerun()
        with btn_c2:
            if st.button("📁 Archivar en Histórico Oficial", type="secondary", use_container_width=True):
                cargas_finales = df_real_edit[df_real_edit["Importe_Real"] > 0].copy()
                if cargas_finales.empty:
                    st.warning("No hay registros mayores a $0 para archivar.")
                else:
                    historial = leer_historico()
                    folio = f"CARGA-{f_prog.strftime('%Y%m%d')}-{len(historial)+1}"
                    
                    registro_cierre_local = {
                        "folio": folio,
                        "fecha_elaboro": f_elab.strftime("%d/%m/%Y"),
                        "fecha_programacion": f_prog.strftime("%d/%m/%Y"),
                        "fecha_registro_sistema": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p"),
                        "total_solicitado": float(total_global_sol),
                        "total_ejercido": float(total_global_real),
                        "ahorro_remanente": float(saldo_global_disponible),
                        "total_vehiculos": int(len(cargas_finales)),
                        "detalle": cargas_finales[["Solicitante", "Vehículo", "Placa", "Operador_Sol", "Importe_Real"]].to_dict(orient="records")
                    }
                    historial.insert(0, registro_cierre_local)
                    guardar_historico(historial)
                    st.success(f"✅ ¡Folio **{folio}** archivado exitosamente!")
                    st.rerun()

    # 5. HISTÓRICO DE CARGAS
    with tab_historico:
        st.markdown("##### 📁 Registro Histórico de Cargas Semanales Finalizadas")
        historial_registros = leer_historico()
        
        if not historial_registros:
            st.info("Aún no hay registros de cargas archivadas en el histórico.")
        else:
            filas_resumen_hist = []
            for h in historial_registros:
                filas_resumen_hist.append({
                    "Folio": h["folio"],
                    "Fecha Programada": h.get("fecha_programacion", h.get("fecha_prog", "")),
                    "Fecha Elaboró": h["fecha_elaboro"],
                    "Vehículos": f"{h['total_vehiculos']} uds",
                    "Total Ejercido ($)": h["total_ejercido"],
                    "Ahorro / Remanente ($)": h.get("ahorro_remanente", h.get("ahorro", 0.0)),
                    "Fecha de Archivo": h.get("fecha_registro_sistema", h.get("fecha_registro", ""))
                })
            
            df_hist_resumen = pd.DataFrame(filas_resumen_hist)
            st.dataframe(
                df_hist_resumen,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Total Ejercido ($)": st.column_config.NumberColumn(format="$%.2f"),
                    "Ahorro / Remanente ($)": st.column_config.NumberColumn(format="$%.2f"),
                }
            )
            
            st.divider()
            st.markdown("##### 🔍 Detalle Individual por Folio / Semana")
            folios_disponibles = [h["folio"] for h in historial_registros]
            folio_sel = st.selectbox("Selecciona un folio:", folios_disponibles)
            registro_sel = next((h for h in historial_registros if h["folio"] == folio_sel), None)
            
            if registro_sel and "detalle" in registro_sel:
                col_h1, col_h2, col_h3 = st.columns(3)
                col_h1.metric("Total Ejercido", f"${registro_sel['total_ejercido']:,.2f}")
                col_h2.metric("Ahorro / Remanente", f"${registro_sel.get('ahorro_remanente', 0.0):,.2f}")
                col_h3.metric("Vehículos que Cargaron", f"{registro_sel['total_vehiculos']} unidades")
                
                df_detalle_folio = pd.DataFrame(registro_sel["detalle"])
                st.dataframe(
                    df_detalle_folio,
                    use_container_width=True,
                    hide_index=True,
                    column_config={"Importe_Real": st.column_config.NumberColumn("Importe Real ($)", format="$%.2f")}
                )

    # 6. MODO PRUEBAS Y MANTENIMIENTO
    with tab_mantenimiento:
        st.markdown("##### 🛠️ Control de Horarios, Simulación y Limpieza")
        
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
            st.subheader("⏰ Estado de Captura de Horario")
            st.success("🟢 Modo de captura 24/7 ACTIVO para todas las áreas (Límite de 3:00 PM suspendido para pruebas).")

        with st.container(border=True):
            st.subheader("🧪 Probar Vista Móvil de Solicitante")
            usuarios_para_test = [u for u in USUARIOS_PASSWORD.keys() if u != "LIAN"]
            solicitante_a_testear = st.selectbox("Selecciona al solicitante a simular:", usuarios_para_test)
            
            if st.button("👁️ Entrar a Modo Simulación", type="secondary", use_container_width=True):
                st.session_state.vista_simulada = solicitante_a_testear
                st.rerun()
