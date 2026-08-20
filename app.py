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

PRESUPUESTO_GLOBAL = 3800.00
HORA_LIMITE = time(15, 0)  # 3:00 PM
ZONA_HORARIA = pytz.timezone("America/Merida")
CONFIG_FILE = "config_sistema.json"
AVISOS_FILE = "avisos_sistema.json"
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
                return json.load(f)[cite: 1]
        except Exception:
            pass
    return {"desbloqueo_horario": False}

def guardar_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f)

def leer_avisos():
    if os.path.exists(AVISOS_FILE):
        try:
            with open(AVISOS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def guardar_avisos(avisos):
    with open(AVISOS_FILE, "w") as f:
        json.dump(avisos, f)

def leer_historico():
    if os.path.exists(HISTORICO_FILE):
        try:
            with open(HISTORICO_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def guardar_historico(hist):
    with open(HISTORICO_FILE, "w") as f:
        json.dump(hist, f)

config_sistema = leer_config()

def obtener_datos_sheets():
    try:
        res = requests.get(WEBHOOK_URL, timeout=15)
        if res.status_code == 200:
            datos_raw = res.json().get("data", [])
            filas_limpias = []
            for item in datos_raw:
                r = int(item["row"])
                if r in MAPEO_SOLICITANTES:
                    filas_limpias.append({
                        "row": r,
                        "Solicitante": MAPEO_SOLICITANTES[r]["solicita"],
                        "Vehículo": MAPEO_SOLICITANTES[r]["vehiculo"],
                        "Placa": MAPEO_SOLICITANTES[r]["placa"],
                        "Actividad": MAPEO_SOLICITANTES[r]["actividad"],
                        "Operador / Encargado": str(item.get("encargado", "")).strip(),
                        "Importe ($)": float(item.get("importe", 0.0)) if item.get("importe") else 0.0
                    })
            return pd.DataFrame(filas_limpias)
    except Exception as err:
        st.error(f"Error de conexión con Google Sheets: {err}")
    return pd.DataFrame()

def guardar_en_sheets(registros, f_elab=None, f_prog=None):
    payload = {"registros": []}
    if f_elab:
        payload["fecha_elaboro"] = f_elab.strftime("%d/%m/%Y")
    if f_prog:
        payload["fecha_prog"] = f_prog.strftime("%d/%m/%Y")
        
    for _, fila in registros.iterrows():
        payload["registros"].append({
            "row": int(fila["row"]),
            "encargado": str(fila["Operador / Encargado"]).strip() if pd.notna(fila["Operador / Encargado"]) else "",
            "importe": float(fila["Importe ($)"]) if pd.notna(fila["Importe ($)"]) else 0.0
        })
        
    try:
        res = requests.post(WEBHOOK_URL, json=payload, timeout=15)
        return res.status_code == 200
    except Exception as e:
        st.error(f"Error al enviar datos: {e}")
        return False

# ==========================================
# GENERADOR EXCEL CON FÓRMULAS NATIVAS
# ==========================================
def generar_excel_magico(df_solicitado, df_real, f_elab, f_prog):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    # Estilos
    fuente_titulo = Font(name="Calibri", size=13, bold=True, color="1F497D")
    fuente_sub = Font(name="Calibri", size=10, bold=True)
    fuente_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    fuente_bold = Font(name="Calibri", size=10, bold=True)
    fuente_normal = Font(name="Calibri", size=10)
    
    fill_header_sol = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_header_real = PatternFill(start_color="1E4D2B", end_color="1E4D2B", fill_type="solid")
    fill_total = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    # ---------------------------------------------
    # PESTAÑA 1: CARGA SOLICITADA
    # ---------------------------------------------
    ws1 = wb.active
    ws1.title = "1. CARGA SOLICITADA"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1["A1"] = "H. AYUNTAMIENTO DE CAMPECHE"
    ws1["A1"].font = fuente_titulo
    ws1["A2"] = "DIRECCIÓN DE DESARROLLO URBANO Y MEDIO AMBIENTE - SOLICITUD PREVENTIVA"
    ws1["A2"].font = fuente_sub
    
    ws1["E3"] = f"Elaboró: {f_elab.strftime('%d/%m/%Y')}"
    ws1["E4"] = f"Programación: {f_prog.strftime('%d/%m/%Y')}"
    ws1["E3"].font = fuente_sub
    ws1["E4"].font = fuente_sub
    
    headers_sol = ["NO.", "SOLICITA", "NOMBRE DEL ENCARGADO", "VEHÍCULO", "PLACA", "RÉGIMEN", "IMPORTE SOLICITADO ($)", "TIPO", "ACTIVIDAD"]
    ws1.append([])
    ws1.append(headers_sol)
    
    header_row_idx = 6
    for col_idx, h in enumerate(headers_sol, 1):
        cell = ws1.cell(row=header_row_idx, column=col_idx)
        cell.font = fuente_header
        cell.fill = fill_header_sol
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    df_sol_activas = df_solicitado[df_solicitado["Importe ($)"] > 0].reset_index(drop=True)
    start_row = 7
    for i, row in df_sol_activas.iterrows():
        current_r = start_row + i
        ws1.append([
            i + 1,
            row["Solicitante"],
            row["Operador / Encargado"],
            row["Vehículo"],
            row["Placa"],
            "OFICIAL",
            float(row["Importe ($)"]),
            "MAGNA",
            row["Actividad"]
        ])
        for c in range(1, 10):
            cell = ws1.cell(row=current_r, column=c)
            cell.font = fuente_normal
            cell.border = thin_border
            if c == 7:
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif c in [1, 5, 6, 8]:
                cell.alignment = Alignment(horizontal="center")
                
    end_row = start_row + len(df_sol_activas) - 1
    if len(df_sol_activas) > 0:
        total_row_sol = end_row + 1
        ws1.cell(row=total_row_sol, column=6, value="TOTAL AUTORIZADO:").font = fuente_bold
        total_cell = ws1.cell(row=total_row_sol, column=7, value=f"=SUM(G{start_row}:G{end_row})")
        total_cell.font = fuente_bold
        total_cell.number_format = '$#,##0.00'
        total_cell.fill = fill_total
        ws1.cell(row=total_row_sol, column=6).alignment = Alignment(horizontal="right")

    # ---------------------------------------------
    # PESTAÑA 2: CARGA REAL Y COMPARATIVA CON FÓRMULAS
    # ---------------------------------------------
    ws2 = wb.create_sheet(title="2. CONCILIACION Y CARGA REAL")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2["A1"] = "H. AYUNTAMIENTO DE CAMPECHE"
    ws2["A1"].font = fuente_titulo
    ws2["A2"] = "AUDITORÍA Y COMPARATIVO DE EJECUCIÓN (SOLICITADO VS REAL)"
    ws2["A2"].font = fuente_sub
    
    headers_real = [
        "NO.", "ÁREA / SOLICITANTE", "VEHÍCULO", "PLACA", "OPERADOR", 
        "SOLICITADO ($)", "CARGA REAL ($)", "DIFERENCIA / AHORRO ($)", "% EJERCIDO", "ESTATUS"
    ]
    ws2.append([])
    ws2.append([])
    ws2.append(headers_real)
    
    r_header_idx = 5
    for col_idx, h in enumerate(headers_real, 1):
        cell = ws2.cell(row=r_header_idx, column=col_idx)
        cell.font = fuente_header
        cell.fill = fill_header_real
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    start_r2 = 6
    for i, row in df_real.iterrows():
        curr = start_r2 + i
        solicitado_val = float(df_solicitado.loc[df_solicitado['row'] == row['row'], 'Importe ($)'].values[0])
        real_val = float(row["Importe ($)"])
        
        # Inserción de fórmulas automatizadas en Excel
        formula_dif = f"=F{curr}-G{curr}"
        formula_pct = f"=IF(F{curr}>0, G{curr}/F{curr}, 0)"
        formula_est = f'=IF(H{curr}=0, "CARGA COMPLETA", IF(H{curr}>0, "CON AHORRO/REMANENTE", "EXCEDIDO"))'
        
        ws2.append([
            i + 1,
            row["Solicitante"],
            row["Vehículo"],
            row["Placa"],
            row["Operador / Encargado"],
            solicitado_val,
            real_val,
            formula_dif,
            formula_pct,
            formula_est
        ])
        
        for c in range(1, 11):
            cell = ws2.cell(row=curr, column=c)
            cell.font = fuente_normal
            cell.border = thin_border
            if c in [6, 7, 8]:
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            elif c == 9:
                cell.number_format = '0.0%'
                cell.alignment = Alignment(horizontal="center")
            elif c in [1, 4, 10]:
                cell.alignment = Alignment(horizontal="center")
                
    end_r2 = start_r2 + len(df_real) - 1
    total_r2 = end_r2 + 1
    
    ws2.cell(row=total_r2, column=5, value="TOTALES:").font = fuente_bold
    ws2.cell(row=total_r2, column=5).alignment = Alignment(horizontal="right")
    
    cell_tot_sol = ws2.cell(row=total_r2, column=6, value=f"=SUM(F{start_r2}:F{end_r2})")
    cell_tot_sol.font = fuente_bold
    cell_tot_sol.number_format = '$#,##0.00'
    cell_tot_sol.fill = fill_total
    
    cell_tot_real = ws2.cell(row=total_r2, column=7, value=f"=SUM(G{start_r2}:G{end_r2})")
    cell_tot_real.font = fuente_bold
    cell_tot_real.number_format = '$#,##0.00'
    cell_tot_real.fill = fill_total
    
    cell_tot_dif = ws2.cell(row=total_r2, column=8, value=f"=F{total_r2}-G{total_r2}")
    cell_tot_dif.font = fuente_bold
    cell_tot_dif.number_format = '$#,##0.00'
    cell_tot_dif.fill = fill_total
    
    cell_tot_pct = ws2.cell(row=total_r2, column=9, value=f"=IF(F{total_r2}>0, G{total_r2}/F{total_r2}, 0)")
    cell_tot_pct.font = fuente_bold
    cell_tot_pct.number_format = '0.0%'
    cell_tot_pct.fill = fill_total
    
    # Autoajuste de anchos de columna
    for sheet in [ws1, ws2]:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
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
            data_row.cell(str(row["Operador / Encargado"]).strip())
            data_row.cell(str(row["Vehículo"]).strip())
            data_row.cell(str(row["Placa"]).strip())
            data_row.cell('OFICIAL')
            data_row.cell(f"${float(row['Importe ($)']):,.2f}")
            data_row.cell('MAGNA')
            data_row.cell(str(row["Actividad"]).strip())
            total += float(row["Importe ($)"])
            
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
# 2. ENCABEZADO Y REVISIÓN DE HORARIO
# ==========================================
usuario_real = st.session_state.usuario_logueado
es_admin_real = (usuario_real == "LIAN")

usuario_efectivo = st.session_state.vista_simulada if (es_admin_real and st.session_state.vista_simulada) else usuario_real
es_admin = (usuario_efectivo == "LIAN")

ahora_local = datetime.now(ZONA_HORARIA)
hora_actual = ahora_local.time()

desbloqueo_activo = config_sistema.get("desbloqueo_horario", False)
sistema_bloqueado = (hora_actual >= HORA_LIMITE) and not es_admin_real and not desbloqueo_activo

c1, c2 = st.columns([3, 1])
with c1:
    st.title("⛽ Control de Combustible")
    if es_admin_real and st.session_state.vista_simulada:
        st.warning(f"🧪 **MODO DE PRUEBA ACTIVO**: Simulando vista como **{usuario_efectivo}**")
    else:
        st.markdown("👑 **ADMINISTRADOR GENERAL**" if es_admin else f"👤 Solicitante: **{usuario_efectivo}**")
        
    estado_horario = "🟢 Horario Abierto" if not sistema_bloqueado else "🔴 Horario Cerrado"
    if desbloqueo_activo:
        estado_horario += " (⚡ Desbloqueo temporal activo)"
    st.caption(f"🕒 {ahora_local.strftime('%I:%M %p')} | {estado_horario}")

with c2:
    st.write("")
    if es_admin_real and st.session_state.vista_simulada:
        if st.button("⬅️ Regresar a Admin", use_container_width=True):
            st.session_state.vista_simulada = None
            st.rerun()
    else:
        if st.button("🚪 Salir", use_container_width=True):
            st.session_state.usuario_logueado = None
            st.session_state.vista_simulada = None
            st.rerun()

df_actual = obtener_datos_sheets()

if df_actual.empty:
    st.warning("Cargando datos...")
    st.stop()

# ==========================================
# 3. VISTA SOLICITANTE
# ==========================================
if not es_admin:
    lista_avisos = leer_avisos()
    mis_avisos = [a for a in lista_avisos if a["destinatario"] in ["TODOS", usuario_efectivo]]
    
    if mis_avisos:
        for av in mis_avisos:
            tipo = av.get("tipo", "Informativo")
            texto_aviso = f"**{av['fecha']} - {av['destinatario']}:** {av['mensaje']}"
            if tipo == "Urgente":
                st.error(f"🚨 {texto_aviso}")
            elif tipo == "Importante":
                st.warning(f"⚠️ {texto_aviso}")
            else:
                st.info(f"📢 {texto_aviso}")
                
    presupuesto_propio = PRESUPUESTO_POR_SOLICITANTE.get(usuario_efectivo, 0.00)
    df_solicitante = df_actual[df_actual["Solicitante"] == usuario_efectivo].copy()
    
    if sistema_bloqueado:
        st.error("🔒 **SISTEMA CERRADO POR HORARIO (3:00 PM)**. La captura está deshabilitada.")
    else:
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
                        value=row["Operador / Encargado"],
                        key=f"op_{row['row']}",
                        disabled=sistema_bloqueado,
                        placeholder="Ej. Juan Pérez"
                    )
                with c_imp:
                    val_importe = st.number_input(
                        "Monto ($)",
                        value=float(row["Importe ($)"]),
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
                    "Operador / Encargado": val_encargado,
                    "Importe ($)": val_importe
                })
        
        df_edit_movil = pd.DataFrame(nuevos_valores)
        total_capturado = df_edit_movil["Importe ($)"].sum()
        saldo_restante = presupuesto_propio - total_capturado
        
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("Presupuesto", f"${presupuesto_propio:,.2f}")
        m2.metric("Total a Cargar", f"${total_capturado:,.2f}")
        m3.metric("Saldo Disponible", f"${saldo_restante:,.2f}", delta_color="normal" if saldo_restante >= 0 else "off")
        
        if total_capturado > presupuesto_propio:
            st.error(f"⚠️ Excedes tu presupuesto por **${abs(saldo_restante):,.2f} MXN**.")
            
        btn_guardar = st.form_submit_button("💾 Guardar Solicitud", type="primary", use_container_width=True, disabled=sistema_bloqueado)
        
        if btn_guardar:
            if total_capturado > presupuesto_propio:
                st.error("No puedes guardar si excedes el presupuesto autorizado.")
            else:
                with st.spinner("Guardando en Google Sheets..."):
                    if guardar_en_sheets(df_edit_movil):
                        st.success("✅ ¡Solicitud guardada con éxito!")
                        st.rerun()
                    else:
                        st.error("❌ Error al guardar.")

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Panel de Consolidación, Auditoría y Automatización")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración", value=date.today())
    with col_f2:
        f_prog = st.date_input("Programación para el día", value=date.today())

    tab_seccion1, tab_seccion2, tab_historico, tab_avisos, tab_mantenimiento = st.tabs([
        "1️⃣ SECCIÓN: Carga Solicitada (Preventiva)", 
        "2️⃣ SECCIÓN: Carga Real y Conciliación",
        "📁 Histórico de Cargas",
        "📢 Centro de Avisos",
        "🛠️ Modo Pruebas"
    ])

    df_solo_cargas = df_actual[df_actual["Importe ($)"] > 0].copy()

    # ==========================================
    # SECCIÓN 1: CARGA SOLICITADA
    # ==========================================
    with tab_seccion1:
        st.markdown("#### 📋 1. Solicitudes Semanales Capturadas por las Áreas")
        
        # Métricas de Saldo
        total_global_sol = df_actual["Importe ($)"].sum()
        saldo_global_sol = PRESUPUESTO_GLOBAL - total_global_sol
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Presupuesto Global", f"${PRESUPUESTO_GLOBAL:,.2f}")
        c2.metric("Total Solicitado", f"${total_global_sol:,.2f}", delta=f"{total_global_sol - PRESUPUESTO_GLOBAL:,.2f}", delta_color="inverse")
        c3.metric("Saldo Disponible", f"${saldo_global_sol:,.2f}", delta_color="normal" if saldo_global_sol >= 0 else "off")
        c4.metric("Bolsa Comodín", "$200.00")
        
        st.markdown("##### 🚗 Lista Oficial de Unidades que Van a Cargar")
        if df_solo_cargas.empty:
            st.warning("⚠️ No hay unidades con monto asignado actualmente.")
        else:
            st.dataframe(
                df_solo_cargas[["Operador / Encargado", "Vehículo", "Placa", "Importe ($)", "Solicitante", "Actividad"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Importe ($)": st.column_config.NumberColumn(format="$%.2f"),
                    "Actividad": st.column_config.TextColumn("Actividad Oficial", width="large")
                }
            )
            
            st.divider()
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                pdf_bytes = generar_pdf_oficial(df_solo_cargas, f_elab, f_prog)
                st.download_button(
                    label="📄 Descargar Oficio Oficial en PDF",
                    data=pdf_bytes,
                    file_name=f"OFICIO_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col_d2:
                # Generación de Excel con ambas pestañas automatizadas
                excel_bytes = generar_excel_magico(df_actual, df_actual, f_elab, f_prog)
                st.download_button(
                    label="📥 Descargar Libro de Excel Automatizado (.xlsx)",
                    data=excel_bytes,
                    file_name=f"CONTROL_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    # ==========================================
    # SECCIÓN 2: CARGA REAL Y CONCILIACIÓN
    # ==========================================
    with tab_seccion2:
        st.markdown("#### 🔍 2. Verificación de Cargas Reales y Conciliación")
        st.caption("Ingresa lo que realmente cargó cada vehículo al presentar sus comprobantes para calcular diferencias y remanentes:")
        
        # Tabla editable para ingresar montos reales
        df_real_edit = st.data_editor(
            df_actual.copy(),
            use_container_width=True,
            disabled=["row", "Solicitante", "Vehículo", "Placa", "Actividad"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Área"),
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placas"),
                "Operador / Encargado": st.column_config.TextColumn("Operador"),
                "Importe ($)": st.column_config.NumberColumn("Monto Realmente Cargado ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "row": None, "Actividad": None
            },
            hide_index=True,
            key="editor_seccion_real"
        )
        
        total_real_ejercido = df_real_edit["Importe ($)"].sum()
        remanente_total = PRESUPUESTO_GLOBAL - total_real_ejercido
        ahorro_vs_solicitado = total_global_sol - total_real_ejercido
        
        st.divider()
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Total Solicitado", f"${total_global_sol:,.2f}")
        r2.metric("Total Realmente Ejercido", f"${total_real_ejercido:,.2f}")
        r3.metric("Ahorro vs Solicitado", f"${ahorro_vs_solicitado:,.2f}", delta_color="normal")
        r4.metric("Saldo Global Recuperado", f"${remanente_total:,.2f}", delta_color="normal")
        
        st.write("")
        c_btn1, c_btn2, c_btn3 = st.columns(3)
        
        with c_btn1:
            if st.button("💾 1. Guardar Cargas Reales en Sheets", type="primary", use_container_width=True):
                with st.spinner("Actualizando Google Sheets..."):
                    if guardar_en_sheets(df_real_edit, f_elab, f_prog):
                        st.success("✅ Datos actualizados en Google Sheets.")
                        st.rerun()
                    else:
                        st.error("❌ Error al actualizar.")
        with c_btn2:
            excel_magico_bytes = generar_excel_magico(df_actual, df_real_edit, f_elab, f_prog)
            st.download_button(
                label="📊 2. Descargar Excel con Fórmulas (.xlsx)",
                data=excel_magico_bytes,
                file_name=f"CONCILIACION_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_btn3:
            if st.button("📁 3. Archivar en Histórico Oficial", type="secondary", use_container_width=True):
                cargas_finales = df_real_edit[df_real_edit["Importe ($)"] > 0].copy()
                if cargas_finales.empty:
                    st.warning("No hay registros mayores a $0 para archivar.")
                else:
                    historial = leer_historico()
                    folio = f"CARGA-{f_prog.strftime('%Y%m%d')}-{len(historial)+1}"
                    
                    registro_cierre = {
                        "folio": folio,
                        "fecha_elaboro": f_elab.strftime("%d/%m/%Y"),
                        "fecha_programacion": f_prog.strftime("%d/%m/%Y"),
                        "fecha_registro_sistema": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p"),
                        "total_solicitado": float(total_global_sol),
                        "total_ejercido": float(total_real_ejercido),
                        "ahorro_remanente": float(remanente_total),
                        "total_vehiculos": int(len(cargas_finales)),
                        "detalle": cargas_finales[["Solicitante", "Vehículo", "Placa", "Operador / Encargado", "Importe ($)"]].to_dict(orient="records")
                    }
                    historial.insert(0, registro_cierre)
                    guardar_historico(historial)
                    st.success(f"✅ ¡Folio **{folio}** archivado exitosamente!")
                    st.rerun()

    # ==========================================
    # HISTÓRICO Y REGISTROS
    # ==========================================
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
                    "Fecha Programada": h["fecha_programacion"],
                    "Fecha Elaboró": h["fecha_elaboro"],
                    "Vehículos": f"{h['total_vehiculos']} uds",
                    "Total Ejercido ($)": h["total_ejercido"],
                    "Ahorro / Remanente ($)": h["ahorro_remanente"],
                    "Fecha de Archivo": h["fecha_registro_sistema"]
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
            
            if registro_sel:
                col_h1, col_h2, col_h3 = st.columns(3)
                col_h1.metric("Total Ejercido", f"${registro_sel['total_ejercido']:,.2f}")
                col_h2.metric("Ahorro / Remanente", f"${registro_sel['ahorro_remanente']:,.2f}")
                col_h3.metric("Vehículos que Cargaron", f"{registro_sel['total_vehiculos']} unidades")
                
                df_detalle_folio = pd.DataFrame(registro_sel["detalle"])
                st.dataframe(
                    df_detalle_folio,
                    use_container_width=True,
                    hide_index=True,
                    column_config={"Importe ($)": st.column_config.NumberColumn(format="$%.2f")}
                )

    # ==========================================
    # CENTRO DE AVISOS
    # ==========================================
    with tab_avisos:
        st.markdown("##### 📢 Publicar Comunicados y Mensajes Personalizados")
        
        with st.form("form_nuevo_aviso"):
            c_dest, c_tipo = st.columns([2, 1])
            with c_dest:
                destinatarios_opciones = ["TODOS"] + [u for u in USUARIOS_PASSWORD.keys() if u != "LIAN"]
                destinatario_sel = st.selectbox("Dirigido a:", destinatarios_opciones)
            with c_tipo:
                tipo_aviso = st.selectbox("Nivel de Prioridad:", ["Informativo", "Importante", "Urgente"])
                
            texto_aviso_nuevo = st.text_area("Contenido del Mensaje:", placeholder="Ej. Recuerden verificar el kilometraje antes de solicitar...")
            
            if st.form_submit_button("📤 Publicar Aviso", type="primary", use_container_width=True):
                if texto_aviso_nuevo.strip():
                    avisos_act = leer_avisos()
                    nuevo_obj = {
                        "id": datetime.now().strftime("%Y%m%d%H%M%S"),
                        "destinatario": destinatario_sel,
                        "tipo": tipo_aviso,
                        "mensaje": texto_aviso_nuevo.strip(),
                        "fecha": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p")
                    }
                    avisos_act.insert(0, nuevo_obj)
                    guardar_avisos(avisos_act)
                    st.success("✅ Aviso publicado exitosamente.")
                    st.rerun()
                else:
                    st.warning("Escribe un mensaje antes de publicar.")
                    
        st.divider()
        st.markdown("##### 📋 Avisos Activos en el Sistema")
        avisos_existentes = leer_avisos()
        
        if not avisos_existentes:
            st.info("No hay avisos ni mensajes activos actualmente.")
        else:
            for idx, av in enumerate(avisos_existentes):
                with st.container(border=True):
                    col_info, col_del = st.columns([5, 1])
                    with col_info:
                        st.markdown(f"**Para:** `{av['destinatario']}` &nbsp;|&nbsp; Prioridad: **{av['tipo']}** &nbsp;|&nbsp; *{av['fecha']}*")
                        st.write(av["mensaje"])
                    with col_del:
                        if st.button("🗑️ Borrar", key=f"del_aviso_{av['id']}"):
                            avisos_existentes.pop(idx)
                            guardar_avisos(avisos_existentes)
                            st.rerun()

    # ==========================================
    # MODO PRUEBAS Y MANTENIMIENTO
    # ==========================================
    with tab_mantenimiento:
        st.markdown("##### 🛠️ Control de Horarios, Simulación y Limpieza")
        
        with st.container(border=True):
            st.subheader("🧹 Reiniciar / Limpiar Cargas a $0.00")
            st.write("Esta acción regresa todos los importes a **$0.00** en Google Sheets para iniciar una nueva semana de captura.")
            
            if st.button("🗑️ Limpiar Todo a $0.00", type="secondary", use_container_width=True):
                df_limpio = df_actual.copy()
                df_limpio["Operador / Encargado"] = ""
                df_limpio["Importe ($)"] = 0.0
                
                with st.spinner("Limpiando registros en Google Sheets..."):
                    if guardar_en_sheets(df_limpio, f_elab, f_prog):
                        st.success("✅ Sistema restablecido a $0.00.")
                        st.rerun()
                    else:
                        st.error("❌ Error al reiniciar datos.")

        with st.container(border=True):
            st.subheader("⏰ Desbloqueo Extemporáneo de Horario (Bypass 3:00 PM)")
            estado_actual_toggle = config_sistema.get("desbloqueo_horario", False)
            toggle_horario = st.toggle("Habilitar captura 24/7 (Desactivar límite de 3:00 PM)", value=estado_actual_toggle)
            
            if toggle_horario != estado_actual_toggle:
                config_sistema["desbloqueo_horario"] = toggle_horario
                guardar_config(config_sistema)
                st.toast("Configuración de horario actualizada.", icon="⏰")
                st.rerun()

        with st.container(border=True):
            st.subheader("🧪 Probar Vista Móvil de Solicitante")
            usuarios_para_test = [u for u in USUARIOS_PASSWORD.keys() if u != "LIAN"]
            solicitante_a_testear = st.selectbox("Selecciona al solicitante a simular:", usuarios_para_test)
            
            if st.button("👁️ Entrar a Modo Simulación", type="secondary", use_container_width=True):
                st.session_state.vista_simulada = solicitante_a_testear
                st.rerun()
