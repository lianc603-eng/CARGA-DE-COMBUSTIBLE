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
# CONFIGURACIONES GENERALES
# ==========================================================
PRESUPUESTO_GLOBAL = 4300.00  # $4,100.00 Base + $350.00 Comodín
BOLSA_COMODIN_TOTAL = 350.00  # Ajustado a $200.00 totales
ZONA_HORARIA = pytz.timezone("America/Merida")

CONFIG_FILE = "config_sistema.json"
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzOjgha2Zjyog01t6LmA_R--EB4Ecqv2ifO_i2YJbLRLbXGShbu5uzFVi85FUTGplM8/exec"

PRESUPUESTO_BASE_POR_SOLICITANTE = {
    "COB CHAVEZ NARCISO DEL JESUS": 200.00,
    "PEREZ MAZIN CARLOS EDUARDO": 200.00,
    "DE LA CRUZ PEREZ WILLIAN ARLEY": 200.00,
    "NOEL CHAN": 850.00,
    "QUEVEDO": 1500.00,
    "RENAN/HELDER": 1000.00,
}

OPERADORES_POR_SOLICITANTE = {
    "COB CHAVEZ NARCISO DEL JESUS": ["JESUS COB"],
    "PEREZ MAZIN CARLOS EDUARDO": ["EDUARDO PEREZ"],
    "DE LA CRUZ PEREZ WILLIAN ARLEY": ["WILLIAN PEREZ", "FRANCISCO ALONZO"],
    "NOEL CHAN": ["AXEL SARAVIA", "NOEL CHAN", "ROMAN DZUL", "ROGER DUARTE", "LUIS CAHUICH", "FRANCISCO CONTRERAS"],
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
    15: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "87GWU8", "actividad": TXT_DESARROLLO_URBANO},
    16: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "88GWU7", "actividad": TXT_DESARROLLO_URBANO},
    17: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "88GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    18: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    19: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    20: {"solicita": "NOEL CHAN", "vehiculo": "MOTO DINAMO", "placa": "90GWU7", "actividad": TXT_MEDIO_AMBIENTE},
    21: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "90GWU8", "actividad": TXT_MEDIO_AMBIENTE},
    23: {"solicita": "RENAN/HELDER", "vehiculo": "AUTOMOVIL JETTA", "placa": "DFT565C", "actividad": TXT_DESARROLLO_URBANO},
    24: {"solicita": "QUEVEDO", "vehiculo": "CAMIONETA RAM 701", "placa": "CN2633B", "actividad": TXT_RAM_AMBIENTAL},
}

# --- USUARIO ÚNICO ADMINISTRADOR ---
PASSWORDS_DEFAULT = {
    "LIAN": "admin123"
}

MAPA_DIA_CARGA = {
    "Lunes": "MARTES",
    "Jueves": "VIERNES"
}

def obtener_nombre_archivo_oficial(prefijo, turno, f_prog, extension):
    dia_nombre = MAPA_DIA_CARGA.get(turno, turno.upper())
    fecha_str = f_prog.strftime("%d_%m_%Y")
    return f"{prefijo}_{dia_nombre}_{fecha_str}.{extension}"

# --- PERSISTENCIA LOCAL Y EN MEMORIA ---
def leer_config():
    if "config_memoria" in st.session_state:
        return st.session_state.config_memoria

    cfg_defecto = {
        "asignacion_comodin": {},
        "transferencias_libres": [],
        "historial_movimientos": [],
        "passwords": PASSWORDS_DEFAULT.copy()
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                c = json.load(f)
                cfg_defecto.update(c)
        except Exception:
            pass
    pwds_limpios = {"LIAN": cfg_defecto.get("passwords", {}).get("LIAN", "admin123")}
    cfg_defecto["passwords"] = pwds_limpios
    if "historial_movimientos" not in cfg_defecto:
        cfg_defecto["historial_movimientos"] = []
        
    st.session_state.config_memoria = cfg_defecto
    return cfg_defecto

def guardar_config(cfg):
    st.session_state.config_memoria = cfg
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f)
    except Exception:
        pass

def obtener_passwords():
    cfg = leer_config()
    return cfg.get("passwords", PASSWORDS_DEFAULT.copy())

def calcular_presupuesto_efectivo():
    cfg = leer_config()
    asig_comodin = cfg.get("asignacion_comodin", {})
    transferencias = cfg.get("transferencias_libres", [])
    
    presupuestos = PRESUPUESTO_BASE_POR_SOLICITANTE.copy()
    
    # 1. Asignar Comodín registrado
    for solicitante, extra in asig_comodin.items():
        if solicitante in presupuestos:
            presupuestos[solicitante] += float(extra)
            
    # 2. Aplicar transferencias registradas entre áreas
    for t in transferencias:
        de_quien = t.get("origen")
        a_quien = t.get("destino")
        monto = float(t.get("monto", 0.0))
        if de_quien in presupuestos and a_quien in presupuestos:
            presupuestos[de_quien] = max(0.0, presupuestos[de_quien] - monto)
            presupuestos[a_quien] += monto
            
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
                    "Operador": op_l, "Importe": imp_l, "Real": real_l
                })
                
                item_j = dict_j.get(r, {})
                op_j = limpiar_texto_operador(item_j.get("encargado", ""))
                imp_j = float(item_j.get("importe", 0.0)) if item_j.get("importe") else 0.0
                real_j = float(item_j.get("real", 0.0)) if item_j.get("real") else 0.0
                
                filas_jueves.append({
                    "row": r, "Solicitante": sol, "Vehículo": info["vehiculo"],
                    "Placa": info["placa"], "Actividad": info["actividad"],
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
                "Operador": "", "Importe": 0.0, "Real": 0.0
            })
            filas_jueves.append({
                "row": r, "Solicitante": info["solicita"], "Vehículo": info["vehiculo"],
                "Placa": info["placa"], "Actividad": info["actividad"],
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

    for r_num in sorted(MAPEO_SOLICITANTES.keys()):
        row_info = df_datos[df_datos["row"] == r_num]
        
        if not row_info.empty:
            item = row_info.iloc[0]
            op_nombre = str(item["Operador"]).strip()
            veh = str(item["Vehículo"]).strip()
            plc = str(item["Placa"]).strip()
            imp_val = float(item["Importe"])
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

    anchos_cols = {'A': 3, 'B': 24, 'C': 18, 'D': 12, 'E': 14, 'F': 10, 'G': 14, 'H': 14, 'I': 52}
    for col_letra, ancho in anchos_cols.items():
        ws.column_dimensions[col_letra].width = ancho

    wb.save(output)
    return output.getvalue()

def generar_pdf_oficial(df_cargas, dia_nombre, f_elab, f_prog):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    dia_encabezado = MAPA_DIA_CARGA.get(dia_nombre, dia_nombre.upper())
    
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 6, 'H. AYUNTAMIENTO DE CAMPECHE', 0, 1, 'C')
    pdf.set_font('Helvetica', 'B', 10)
    pdf.cell(0, 5, f'DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE - OFICIO DE CARGA {dia_encabezado}', 0, 1, 'C')
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
            data_row.cell(str(row["Operador"]).strip())
            data_row.cell(str(row["Vehículo"]).strip())
            data_row.cell(str(row["Placa"]).strip())
            data_row.cell('OFICIAL')
            data_row.cell(f"${float(row['Importe']):,.2f}")
            data_row.cell('MAGNA')
            data_row.cell(str(row["Actividad"]).strip())
            total += float(row["Importe"])
            
        pdf.set_font('Helvetica', 'B', 8)
        tot_row = table.row()
        tot_row.cell(f'TOTAL AUTORIZADO {dia_encabezado}:', colspan=4, align="RIGHT")
        tot_row.cell(f"${total:,.2f}", align="RIGHT")
        tot_row.cell('', colspan=2)
        
    return bytes(pdf.output())

def verificar_bloqueo_descarga(df_dia_edit, df_otro_dia, presupuestos_efectivos):
    bloqueado = False
    mensajes = []
    for sol_nom, pres_max in presupuestos_efectivos.items():
        gasto_1 = df_dia_edit[df_dia_edit["Solicitante"] == sol_nom]["Importe"].sum()
        gasto_2 = df_otro_dia[df_otro_dia["Solicitante"] == sol_nom]["Importe"].sum()
        tot_sem = gasto_1 + gasto_2
        if tot_sem > pres_max:
            bloqueado = True
            diff = tot_sem - pres_max
            placas = ", ".join(df_dia_edit[df_dia_edit["Solicitante"] == sol_nom]["Placa"].tolist())
            mensajes.append(
                f"• **{sol_nom}** (Placas: `{placas}`): acumulado semanal **${tot_sem:,.2f}** vs autorizado **${pres_max:,.2f}** (Excedido por **${diff:,.2f}**)."
            )
    tot_global = df_dia_edit["Importe"].sum() + df_otro_dia["Importe"].sum()
    if tot_global > PRESUPUESTO_GLOBAL:
        bloqueo_true = True
        bloqueado = True
        diff_g = tot_global - PRESUPUESTO_GLOBAL
        mensajes.append(
            f"• **Presupuesto Global Semanal**: Acumulado de **${tot_global:,.2f}** excede los **${PRESUPUESTO_GLOBAL:,.2f}** por **${diff_g:,.2f}**."
        )
    return bloqueado, mensajes

# ==========================================
# 1. INICIO DE SESIÓN (SOLO LIAN)
# ==========================================
credenciales_actuales = obtener_passwords()

if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None

if st.session_state.usuario_logueado is None:
    st.title("⛽ Control y Solicitud de Combustible")
    st.caption("Dirección de Desarrollo Urbano y Medio Ambiente — H. Ayuntamiento de Campeche")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("form_login"):
            st.subheader("🔐 Acceso Administrador (LIAN)")
            pwd = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Entrar", use_container_width=True):
                if pwd == credenciales_actuales.get("LIAN", "admin123"):
                    st.session_state.usuario_logueado = "LIAN"
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
    st.stop()

# ==========================================
# 2. ENCABEZADO Y CONFIGURACIÓN
# ==========================================
ahora_local = datetime.now(ZONA_HORARIA)
FECHA_HOY = date.today()
FECHA_PROGRAMACION_DEFECTO = FECHA_HOY + timedelta(days=1)

cfg_actual = leer_config()
presupuestos_actuales = calcular_presupuesto_efectivo()

c1, c2 = st.columns([3.2, 0.8])
with c1:
    st.title("⛽ Control de Combustible")
    st.markdown("👑 **ADMINISTRADOR GENERAL (LIAN)**")
    st.caption(f"🕒 {ahora_local.strftime('%d/%m/%Y %I:%M %p')}")

with c2:
    st.write("")
    if st.button("🔄 Sincronizar", use_container_width=True):
        obtener_datos_dos_hojas(forzar=True)
        st.toast("Datos sincronizados con Google Sheets.", icon="✅")
        st.rerun()
    if st.button("🚪 Salir", use_container_width=True):
        st.session_state.usuario_logueado = None
        st.rerun()

# --- CAMBIO DE CONTRASEÑA ---
with st.expander("🔑 Cambiar mi contraseña de acceso"):
    c_p1, c_p2, c_p3 = st.columns([2, 2, 1.2])
    with c_p1:
        nueva_pass = st.text_input("Nueva contraseña (mínimo 4 caracteres):", type="password", key="inp_nueva_pass")
    with c_p2:
        conf_pass = st.text_input("Confirma la nueva contraseña:", type="password", key="inp_conf_pass")
    with c_p3:
        st.write("")
        st.write("")
        if st.button("💾 Actualizar Clave", use_container_width=True):
            if len(nueva_pass.strip()) < 4:
                st.error("La contraseña debe tener al menos 4 caracteres.")
            elif nueva_pass != conf_pass:
                st.error("Las contraseñas no coinciden.")
            else:
                cfg_actual["passwords"] = {"LIAN": nueva_pass.strip()}
                guardar_config(cfg_actual)
                st.success("✅ ¡Contraseña actualizada exitosamente!")

df_lunes, df_jueves = obtener_datos_dos_hojas()

# ==========================================
# 3. PANEL PRINCIPAL ADMINISTRADOR (LIAN)
# ==========================================
st.subheader("⚙️ Panel de Consolidación, Edición y Descarga Oficial")

col_f1, col_f2 = st.columns(2)
with col_f1:
    f_elab = st.date_input("Fecha de Elaboración", value=FECHA_HOY)
with col_f2:
    f_prog = st.date_input("Programación para el día", value=FECHA_PROGRAMACION_DEFECTO)

tab_saldos, tab_lunes, tab_jueves, tab_mi_carga, tab_auditoria, tab_mantenimiento = st.tabs([
    "📊 Monitoreo y Asignación de Presupuestos",
    "📄 Formato y Carga LUNES",
    "📄 Formato y Carga JUEVES",
    "🛵 Mi Carga (LIAN / Alonzo)",
    "✅ Auditoría y Cierre Semanal",
    "🛠️ Cierre de Viernes"
])

# 1. MONITOREO Y TRANSFERENCIA DE PRESUPUESTOS
with tab_saldos:
    st.markdown("##### 💵 Balance Semanal de Presupuestos en Tiempo Real")
    
    tot_lunes_sol = df_lunes["Importe"].sum()
    tot_jueves_sol = df_jueves["Importe"].sum()
    tot_global_semana = tot_lunes_sol + tot_jueves_sol
    saldo_global_semana = PRESUPUESTO_GLOBAL - tot_global_semana
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Presupuesto Semanal Global", f"${PRESUPUESTO_GLOBAL:,.2f}")
    c2.metric("Total Solicitado Lunes", f"${tot_lunes_sol:,.2f}")
    c3.metric("Total Solicitado Jueves", f"${tot_jueves_sol:,.2f}")
    c4.metric(
        "Saldo Libre Semana", 
        f"${saldo_global_semana:,.2f}", 
        delta=f"${tot_global_semana:,.2f} total pedido",
        delta_color="normal" if saldo_global_semana >= 0 else "inverse"
    )
    
    st.write("")
    st.markdown("##### 📋 Resumen Financiero por Solicitante")
    filas_saldos_area = []
    saldos_restantes_por_solicitante = {}
    excedentes_sin_cubrir = 0.0

    for sol, p_efectivo in presupuestos_actuales.items():
        sub_l = df_lunes[df_lunes["Solicitante"] == sol]
        sub_j = df_jueves[df_jueves["Solicitante"] == sol]
        sol_l = sub_l["Importe"].sum()
        sol_j = sub_j["Importe"].sum()
        sol_tot = sol_l + sol_j
        disp_monto = p_efectivo - sol_tot
        saldos_restantes_por_solicitante[sol] = disp_monto
        
        if disp_monto < 0:
            excedentes_sin_cubrir += abs(disp_monto)
        
        if disp_monto == 0:
            estatus = "✅ 100% Ejercido"
        elif disp_monto < 0:
            estatus = "⚠️ Excedido (Requiere Transferencia/Comodín)"
        elif sol_tot > 0:
            estatus = "🟢 Con Saldo"
        else:
            estatus = "⚪ Sin Carga"
            
        filas_saldos_area.append({
            "Solicitante / Área": sol,
            "Presupuesto Base": PRESUPUESTO_BASE_POR_SOLICITANTE.get(sol, 0.0),
            "Presupuesto Semanal Autorizado": p_efectivo,
            "Carga Lunes": sol_l,
            "Carga Jueves": sol_j,
            "Total Solicitado": sol_tot,
            "Saldo Disponible Restante": disp_monto,
            "Estatus": estatus
        })
        
    df_saldos_area = pd.DataFrame(filas_saldos_area)
    st.dataframe(
        df_saldos_area,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Presupuesto Base": st.column_config.NumberColumn(format="$%.2f"),
            "Presupuesto Semanal Autorizado": st.column_config.NumberColumn(format="$%.2f"),
            "Carga Lunes": st.column_config.NumberColumn(format="$%.2f"),
            "Carga Jueves": st.column_config.NumberColumn(format="$%.2f"),
            "Total Solicitado": st.column_config.NumberColumn(format="$%.2f"),
            "Saldo Disponible Restante": st.column_config.NumberColumn(format="$%.2f"),
        }
    )

    st.divider()
    st.markdown("##### 🔀 Asignar Comodín y Transferir Presupuesto entre Solicitantes")
    
    # Cálculo estricto de Comodín Disponible ($1,000.00 Totales)
    asig_comodin_dict = cfg_actual.get("asignacion_comodin", {})
    total_comodin_asignado = sum(asig_comodin_dict.values())
    comodin_restante_bolsa = max(0.0, BOLSA_COMODIN_TOTAL - total_comodin_asignado)
    comodin_real_disponible = max(0.0, min(comodin_restante_bolsa, max(0.0, saldo_global_semana + excedentes_sin_cubrir)))

    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        with st.container(border=True):
            st.markdown(f"🎁 **Bolsa Comodín (${BOLSA_COMODIN_TOTAL:,.2f} Semanales)**")
            if comodin_real_disponible > 0:
                st.info(
                    f"💰 **Saldo Comodín Disponible para Asignar:** `${comodin_real_disponible:,.2f}`\n\n"
                    f"*(Ya asignado en la semana: `${total_comodin_asignado:,.2f}` de `${BOLSA_COMODIN_TOTAL:,.2f}`)*"
                )
            else:
                st.error(
                    f"🔒 **COMODÍN AGOTADO ($0.00 DISPONIBLE)**\n\n"
                    f"Ya se asignaron o consumieron los `${BOLSA_COMODIN_TOTAL:,.2f}` del comodín en la semana."
                )
            
            destinatario_comodin = st.selectbox(
                "Asignar Comodín a:", 
                list(PRESUPUESTO_BASE_POR_SOLICITANTE.keys()), 
                key="sel_comodin",
                disabled=(comodin_real_disponible <= 0)
            )
            saldo_dest_comodin = saldos_restantes_por_solicitante.get(destinatario_comodin, 0.0)
            comodin_ya_recibido = asig_comodin_dict.get(destinatario_comodin, 0.0)
            
            st.caption(
                f"📊 Saldo restante actual de **{destinatario_comodin}**: **`${saldo_dest_comodin:,.2f}`** "
                f"| Comodín recibido: **`${comodin_ya_recibido:,.2f}`**"
            )
            
            monto_comodin_add = st.number_input(
                "Monto de Comodín a Asignar ($)", 
                min_value=0.0, 
                max_value=float(comodin_real_disponible) if comodin_real_disponible > 0 else 0.0, 
                step=50.0, 
                value=0.0, 
                key="inp_comodin",
                disabled=(comodin_real_disponible <= 0)
            )
            
            if st.button("➕ Registrar y Asignar Comodín", use_container_width=True, disabled=(comodin_real_disponible <= 0)):
                if monto_comodin_add <= 0:
                    st.warning("Ingresa un monto mayor a $0.00.")
                elif monto_comodin_add > comodin_real_disponible:
                    st.error("No hay suficiente saldo en el comodín.")
                else:
                    asig_act = cfg_actual.get("asignacion_comodin", {})
                    asig_act[destinatario_comodin] = asig_act.get(destinatario_comodin, 0.0) + monto_comodin_add
                    cfg_actual["asignacion_comodin"] = asig_act
                    
                    hist = cfg_actual.get("historial_movimientos", [])
                    hist.append({
                        "fecha": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p"),
                        "tipo": "🎁 ASIGNACIÓN COMODÍN",
                        "origen": "BOLSA COMODÍN",
                        "destino": destinatario_comodin,
                        "monto": float(monto_comodin_add)
                    })
                    cfg_actual["historial_movimientos"] = hist
                    guardar_config(cfg_actual)
                    st.toast(f"Asignados ${monto_comodin_add:,.2f} de comodín a {destinatario_comodin}", icon="✅")
                    st.rerun()

    with col_t2:
        with st.container(border=True):
            st.markdown("🤝 **Transferencia Directa entre Áreas (Solo Saldo Libre)**")
            lista_solicitantes = list(PRESUPUESTO_BASE_POR_SOLICITANTE.keys())
            
            c_origen, c_destino = st.columns(2)
            with c_origen:
                sol_origen = st.selectbox("Ceder de (Origen):", lista_solicitantes, key="sel_origen_trans")
            with c_destino:
                destinatarios_posibles = [s for s in lista_solicitantes if s != sol_origen]
                sol_destino = st.selectbox("Transferir a (Destino):", destinatarios_posibles, key="sel_destino_trans")
            
            saldo_restante_origen = max(0.0, float(saldos_restantes_por_solicitante.get(sol_origen, 0.0)))
            saldo_restante_destino = float(saldos_restantes_por_solicitante.get(sol_destino, 0.0))
            
            if saldo_restante_origen > 0:
                st.info(
                    f"💵 **Saldo Disponible para Transferir en {sol_origen}:** `${saldo_restante_origen:,.2f}`\n\n"
                    f"📥 **Saldo Restante Actual en {sol_destino}:** `${saldo_restante_destino:,.2f}`"
                )
            else:
                st.error(
                    f"🔒 **{sol_origen} NO TIENE SALDO DISPONIBLE ($0.00)**\n\n"
                    f"Ya ejerció o cedió todo su presupuesto en la solicitud de la semana. No puede transferir fondos."
                )
            
            monto_transferir = st.number_input(
                "Monto a Transferir ($)", 
                min_value=0.0, 
                max_value=saldo_restante_origen if saldo_restante_origen > 0 else 0.0, 
                step=50.0, 
                value=0.0, 
                key="inp_monto_trans",
                disabled=(saldo_restante_origen <= 0)
            )
            
            if st.button("🔄 Registrar y Ejecutar Transferencia", use_container_width=True, disabled=(saldo_restante_origen <= 0)):
                if monto_transferir <= 0:
                    st.warning("Ingresa un monto mayor a $0.00.")
                elif monto_transferir > saldo_restante_origen:
                    st.error("El origen no cuenta con ese saldo disponible.")
                else:
                    trans_list = cfg_actual.get("transferencias_libres", [])
                    trans_list.append({
                        "origen": sol_origen,
                        "destino": sol_destino,
                        "monto": float(monto_transferir)
                    })
                    cfg_actual["transferencias_libres"] = trans_list
                    
                    hist = cfg_actual.get("historial_movimientos", [])
                    hist.append({
                        "fecha": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p"),
                        "tipo": "🤝 TRANSFERENCIA DE SALDO",
                        "origen": sol_origen,
                        "destino": sol_destino,
                        "monto": float(monto_transferir)
                    })
                    cfg_actual["historial_movimientos"] = hist
                    guardar_config(cfg_actual)
                    st.toast(f"Transferidos ${monto_transferir:,.2f} de {sol_origen} a {sol_destino}", icon="🤝")
                    st.rerun()

    # --- BITÁCORA DE MOVIMIENTOS REGISTRADOS EN LA SEMANA ---
    historial_movs = cfg_actual.get("historial_movimientos", [])
    if historial_movs or cfg_actual.get("asignacion_comodin") or cfg_actual.get("transferencias_libres"):
        st.write("")
        st.markdown("##### 📜 Bitácora de Asignaciones de Comodín y Transferencias Registradas")
        if historial_movs:
            df_hist_mov = pd.DataFrame(historial_movs)
            df_hist_mov.columns = ["Fecha y Hora", "Tipo de Movimiento", "Origen", "Destino", "Monto ($)"]
            st.dataframe(
                df_hist_mov,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Monto ($)": st.column_config.NumberColumn(format="$%.2f")
                }
            )
        if st.button("🔄 Restablecer Todas las Transferencias y Comodín a Valores Base", type="secondary", use_container_width=True):
            cfg_actual["asignacion_comodin"] = {}
            cfg_actual["transferencias_libres"] = []
            cfg_actual["historial_movimientos"] = []
            cfg_actual["cesion_lian"] = {}
            guardar_config(cfg_actual)
            st.toast("Presupuestos y bitácora restablecidos a valores originales.", icon="🔄")
            st.rerun()

# 2. CARGA LUNES
with tab_lunes:
    st.markdown("##### 🚗 Solicitud Oficial de Carga del LUNES (Carga Efectiva: MARTES)")
    st.caption("Selecciona el operador o usa el desplegable inferior para quitar una carga y liberar el presupuesto:")
    
    with st.form("form_admin_lunes_exclusivo"):
        nuevos_lunes = []
        
        c_h1, c_h2, c_h3, c_h4, c_h5 = st.columns([1.6, 1.3, 1.0, 2.0, 1.1])
        c_h1.markdown("**Área / Solicitante**")
        c_h2.markdown("**Vehículo**")
        c_h3.markdown("**Placa**")
        c_h4.markdown("**Operador Autorizado**")
        c_h5.markdown("**Importe ($)**")
        st.divider()

        for idx, row in df_lunes.iterrows():
            r_num = int(row["row"])
            sol = row["Solicitante"]
            ops_disponibles = [""] + OPERADORES_POR_SOLICITANTE.get(sol, [])
            
            val_act = limpiar_texto_operador(row["Operador"])
            if val_act and val_act not in ops_disponibles:
                ops_disponibles.append(val_act)
            ops_disponibles.append("➕ OTRO (ESCRIBIR NOMBRE)")
                
            idx_op = ops_disponibles.index(val_act) if val_act in ops_disponibles else 0
            
            c_a, c_v, c_p, c_o, c_i = st.columns([1.6, 1.3, 1.0, 2.0, 1.1])
            c_a.write(sol)
            c_v.write(row["Vehículo"])
            c_p.code(row["Placa"])
            
            sel_op = c_o.selectbox(
                f"Op Lunes {r_num}",
                options=ops_disponibles,
                index=idx_op,
                key=f"adm_op_l_{r_num}",
                label_visibility="collapsed"
            )
            
            op_final_l = sel_op
            if sel_op == "➕ OTRO (ESCRIBIR NOMBRE)":
                op_ext_l = st.text_input(
                    f"Nuevo operador {row['Vehículo']}:",
                    key=f"adm_ext_op_l_{r_num}",
                    placeholder="NOMBRE COMPLETO"
                )
                op_final_l = op_ext_l.strip().upper()
            
            inp_imp = c_i.number_input(
                f"Imp Lunes {r_num}",
                value=float(row["Importe"]),
                step=50.0,
                min_value=0.0,
                key=f"adm_imp_l_{r_num}",
                format="%.2f",
                label_visibility="collapsed"
            )
            
            nuevos_lunes.append({
                "row": r_num, "Solicitante": sol, "Vehículo": row["Vehículo"],
                "Placa": row["Placa"], "Actividad": row["Actividad"],
                "Operador": op_final_l, "Importe": inp_imp, "Real": row["Real"]
            })
            
        df_lunes_edit = pd.DataFrame(nuevos_lunes)
        
        btn_guardar_lunes = st.form_submit_button("💾 Guardar Cambios de LUNES en Google Sheets", type="primary", use_container_width=True)
        if btn_guardar_lunes:
            with st.spinner("Guardando en pestaña 'lunes'..."):
                exito = enviar_datos_hoja(df_lunes_edit, hoja="lunes", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                if exito:
                    st.success("✅ ¡Cargas del Lunes guardadas y sincronizadas!")
                    st.rerun()
                else:
                    st.error("Error al guardar en Google Sheets.")

    with st.expander("🗑️ Quitar Carga Específica del LUNES (Regresar Saldo al Presupuesto)"):
        c_del1, c_del2 = st.columns([3, 1])
        with c_del1:
            veh_borrar_l = st.selectbox(
                "Selecciona el vehículo cuya carga deseas eliminar:",
                options=[f"Fila {r} - {info['solicita']} ({info['vehiculo']} {info['placa']})" for r, info in MAPEO_SOLICITANTES.items()],
                key="sel_del_lunes"
            )
        with c_del2:
            st.write("")
            if st.button("❌ Quitar Carga", key="btn_del_lunes", use_container_width=True):
                r_target = int(veh_borrar_l.split(" - ")[0].replace("Fila ", ""))
                df_lunes_edit.loc[df_lunes_edit["row"] == r_target, "Operador"] = ""
                df_lunes_edit.loc[df_lunes_edit["row"] == r_target, "Importe"] = 0.0
                enviar_datos_hoja(df_lunes_edit, hoja="lunes", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                st.toast(f"Carga de la Fila {r_target} eliminada. Saldo reintegrado.", icon="✅")
                st.rerun()

    st.markdown("---")
    bloqueo_l, det_l = verificar_bloqueo_descarga(df_lunes_edit, df_jueves, presupuestos_actuales)
    if bloqueo_l:
        st.error("⛔ **DESCARGA DE ARCHIVOS BLOQUEADA POR EXCEDENTE DE PRESUPUESTO**")
        st.warning("Revisa las siguientes unidades:\n\n" + "\n".join(det_l) + "\n\n💡 *Asigna comodín o transfiere presupuesto en la Pestaña 1 para desbloquear.*")
    else:
        st.success("✅ Presupuestos en orden. Formatos listos para descarga.")

    df_solo_lunes = df_lunes_edit[df_lunes_edit["Importe"] > 0].copy()
    nombre_excel_l = obtener_nombre_archivo_oficial("SOLICITUD", "Lunes", f_prog, "xlsx")
    nombre_pdf_l = obtener_nombre_archivo_oficial("OFICIO", "Lunes", f_prog, "pdf")
    
    excel_lunes = generar_excel_oficial_formato(df_lunes_edit, "Lunes", f_elab, f_prog) if not bloqueo_l else b""
    pdf_lunes = generar_pdf_oficial(df_solo_lunes, "Lunes", f_elab, f_prog) if not bloqueo_l else b""
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label=f"📥 Descargar Formato Excel ({nombre_excel_l})",
            data=excel_lunes,
            file_name=nombre_excel_l,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=bloqueo_l
        )
    with col_dl2:
        st.download_button(
            label=f"📄 Descargar Oficio PDF ({nombre_pdf_l})",
            data=pdf_lunes,
            file_name=nombre_pdf_l,
            mime="application/pdf",
            use_container_width=True,
            disabled=bloqueo_l
        )

# 3. CARGA JUEVES
with tab_jueves:
    st.markdown("##### 🚗 Solicitud Oficial de Carga del JUEVES (Carga Efectiva: VIERNES)")
    st.info(f"💡 El Lunes se solicitaron **${tot_lunes_sol:,.2f}**. El saldo libre global restante para cerrar la semana es de **${max(0.0, PRESUPUESTO_GLOBAL - tot_lunes_sol):,.2f}**.")
    
    with st.form("form_admin_jueves_exclusivo"):
        nuevos_jueves = []
        
        c_jh1, c_jh2, c_jh3, c_jh4, c_jh5 = st.columns([1.6, 1.3, 1.0, 2.0, 1.1])
        c_jh1.markdown("**Área / Solicitante**")
        c_jh2.markdown("**Vehículo**")
        c_jh3.markdown("**Placa**")
        c_jh4.markdown("**Operador Autorizado**")
        c_jh5.markdown("**Importe ($)**")
        st.divider()

        for idx, row in df_jueves.iterrows():
            r_num = int(row["row"])
            sol = row["Solicitante"]
            ops_disponibles = [""] + OPERADORES_POR_SOLICITANTE.get(sol, [])
            
            val_act = limpiar_texto_operador(row["Operador"])
            if val_act and val_act not in ops_disponibles:
                ops_disponibles.append(val_act)
            ops_disponibles.append("➕ OTRO (ESCRIBIR NOMBRE)")
                
            idx_op = ops_disponibles.index(val_act) if val_act in ops_disponibles else 0
            
            c_a, c_v, c_p, c_o, c_i = st.columns([1.6, 1.3, 1.0, 2.0, 1.1])
            c_a.write(sol)
            c_v.write(row["Vehículo"])
            c_p.code(row["Placa"])
            
            sel_op = c_o.selectbox(
                f"Op Jueves {r_num}",
                options=ops_disponibles,
                index=idx_op,
                key=f"adm_op_j_{r_num}",
                label_visibility="collapsed"
            )
            
            op_final_j = sel_op
            if sel_op == "➕ OTRO (ESCRIBIR NOMBRE)":
                op_ext_j = st.text_input(
                    f"Nuevo operador {row['Vehículo']}:",
                    key=f"adm_ext_op_j_{r_num}",
                    placeholder="NOMBRE COMPLETO"
                )
                op_final_j = op_ext_j.strip().upper()
            
            inp_imp = c_i.number_input(
                f"Imp Jueves {r_num}",
                value=float(row["Importe"]),
                step=50.0,
                min_value=0.0,
                key=f"adm_imp_j_{r_num}",
                format="%.2f",
                label_visibility="collapsed"
            )
            
            nuevos_jueves.append({
                "row": r_num, "Solicitante": sol, "Vehículo": row["Vehículo"],
                "Placa": row["Placa"], "Actividad": row["Actividad"],
                "Operador": op_final_j, "Importe": inp_imp, "Real": row["Real"]
            })
            
        df_jueves_edit = pd.DataFrame(nuevos_jueves)
        
        btn_guardar_jueves = st.form_submit_button("💾 Guardar Cambios de JUEVES en Google Sheets", type="primary", use_container_width=True)
        if btn_guardar_jueves:
            with st.spinner("Guardando en pestaña 'jueves'..."):
                exito = enviar_datos_hoja(df_jueves_edit, hoja="jueves", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                if exito:
                    st.success("✅ ¡Cargas del Jueves guardadas y sincronizadas!")
                    st.rerun()
                else:
                    st.error("Error al guardar en Google Sheets.")

    with st.expander("🗑️ Quitar Carga Específica del JUEVES (Regresar Saldo al Presupuesto)"):
        c_delj1, c_delj2 = st.columns([3, 1])
        with c_delj1:
            veh_borrar_j = st.selectbox(
                "Selecciona el vehículo cuya carga deseas eliminar:",
                options=[f"Fila {r} - {info['solicita']} ({info['vehiculo']} {info['placa']})" for r, info in MAPEO_SOLICITANTES.items()],
                key="sel_del_jueves"
            )
        with c_delj2:
            st.write("")
            if st.button("❌ Quitar Carga", key="btn_del_jueves", use_container_width=True):
                r_target = int(veh_borrar_j.split(" - ")[0].replace("Fila ", ""))
                df_jueves_edit.loc[df_jueves_edit["row"] == r_target, "Operador"] = ""
                df_jueves_edit.loc[df_jueves_edit["row"] == r_target, "Importe"] = 0.0
                enviar_datos_hoja(df_jueves_edit, hoja="jueves", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                st.toast(f"Carga de la Fila {r_target} eliminada. Saldo reintegrado.", icon="✅")
                st.rerun()

    st.markdown("---")
    bloqueo_j, det_j = verificar_bloqueo_descarga(df_jueves_edit, df_lunes, presupuestos_actuales)
    if bloqueo_j:
        st.error("⛔ **DESCARGA DE ARCHIVOS BLOQUEADA POR EXCEDENTE DE PRESUPUESTO**")
        st.warning("Revisa las siguientes unidades:\n\n" + "\n".join(det_j) + "\n\n💡 *Asigna comodín o transfiere presupuesto en la Pestaña 1 para desbloquear.*")
    else:
        st.success("✅ Presupuestos en orden. Formatos listos para descarga.")

    df_solo_jueves = df_jueves_edit[df_jueves_edit["Importe"] > 0].copy()
    nombre_excel_j = obtener_nombre_archivo_oficial("SOLICITUD", "Jueves", f_prog, "xlsx")
    nombre_pdf_j = obtener_nombre_archivo_oficial("OFICIO", "Jueves", f_prog, "pdf")
    
    excel_jueves = generar_excel_oficial_formato(df_jueves_edit, "Jueves", f_elab, f_prog) if not bloqueo_j else b""
    pdf_jueves = generar_pdf_oficial(df_solo_jueves, "Jueves", f_elab, f_prog) if not bloqueo_j else b""
    
    col_dj1, col_dj2 = st.columns(2)
    with col_dj1:
        st.download_button(
            label=f"📥 Descargar Formato Excel ({nombre_excel_j})",
            data=excel_jueves,
            file_name=nombre_excel_j,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=bloqueo_j
        )
    with col_dj2:
        st.download_button(
            label=f"📄 Descargar Oficio PDF ({nombre_pdf_j})",
            data=pdf_jueves,
            file_name=nombre_pdf_j,
            mime="application/pdf",
            use_container_width=True,
            disabled=bloqueo_j
        )

# 4. MI CARGA (LIAN / FRANCISCO ALONZO)
with tab_mi_carga:
    st.markdown("##### 🛵 Mi Carga (LIAN / Francisco Alonzo)")
    sub_l_lian = df_lunes[df_lunes["Solicitante"] == "LIAN"]
    sub_j_lian = df_jueves[df_jueves["Solicitante"] == "LIAN"]
    presupuesto_lian_efectivo = presupuestos_actuales["LIAN"]
    
    if not sub_l_lian.empty:
        row_l_lian = sub_l_lian.iloc[0]
        row_j_lian = sub_j_lian.iloc[0]
        
        with st.container(border=True):
            st.markdown(f"🛵 **{row_l_lian['Vehículo']}** &nbsp;|&nbsp; Placa: **`{row_l_lian['Placa']}`** &nbsp;|&nbsp; Presupuesto Autorizado: **${presupuesto_lian_efectivo:,.2f}**")
            
            c_ml1, c_ml2 = st.columns(2)
            with c_ml1:
                st.markdown("🗓️ **Carga del Lunes**")
                op_l_guardado = limpiar_texto_operador(row_l_lian["Operador"])
                op_l_sel = st.selectbox("Operador Lunes", ["", "FRANCISCO ALONZO"], index=1 if op_l_guardado == "FRANCISCO ALONZO" else 0, key="ml_op_l")
                
                val_l_actual = float(row_l_lian["Importe"])
                max_l = float(presupuesto_lian_efectivo)
                val_l_seguro = max(0.0, min(val_l_actual, max_l))
                
                imp_l_sel = st.number_input("Monto Lunes ($)", value=val_l_seguro, step=50.0, min_value=0.0, max_value=max_l, key="ml_imp_l")
            
            with c_ml2:
                st.markdown("🗓️ **Carga del Jueves**")
                disponible_jueves_lian = max(0.0, presupuesto_lian_efectivo - imp_l_sel)
                op_j_guardado = limpiar_texto_operador(row_j_lian["Operador"])
                op_j_sel = st.selectbox("Operador Jueves", ["", "FRANCISCO ALONZO"], index=1 if op_j_guardado == "FRANCISCO ALONZO" else 0, key="ml_op_j")
                
                val_j_actual = float(row_j_lian["Importe"])
                max_j = float(disponible_jueves_lian)
                val_j_seguro = max(0.0, min(val_j_actual, max_j))
                
                imp_j_sel = st.number_input("Monto Jueves ($)", value=val_j_seguro, step=50.0, min_value=0.0, max_value=max_j, key="ml_imp_j")
            
            if st.button("💾 Guardar Mi Carga (Lunes y Jueves)", type="primary", use_container_width=True):
                df_mi_l = df_lunes[df_lunes["Solicitante"] == "LIAN"].copy()
                df_mi_l["Operador"] = op_l_sel
                df_mi_l["Importe"] = imp_l_sel
                enviar_datos_hoja(df_mi_l, hoja="lunes", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                
                df_mi_j = df_jueves[df_jueves["Solicitante"] == "LIAN"].copy()
                df_mi_j["Operador"] = op_j_sel
                df_mi_j["Importe"] = imp_j_sel
                enviar_datos_hoja(df_mi_j, hoja="jueves", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
                st.success("✅ Tu carga fue registrada correctamente.")
                st.rerun()

# 5. AUDITORÍA Y COMPROBACIÓN REAL (COLUMNA BASE TOTALMENTE ELIMINADA)
with tab_auditoria:
    st.markdown("##### 🔍 Auditoría y Comprobación Real (Pestañas 'lunes' y 'jueves')")
    st.caption("Captura lo que realmente cargaron cada día para conciliar y archivar en 'historico':")
    
    c_aud_l, c_aud_j = st.columns(2)
    cols_visibles_aud = ["Vehículo", "Placa", "Importe", "Real"]
    
    with c_aud_l:
        st.markdown("📋 **Comprobación Real LUNES**")
        df_aud_l_edit = st.data_editor(
            df_lunes[cols_visibles_aud].copy(),
            use_container_width=True,
            height=450,
            disabled=["Vehículo", "Placa", "Importe"],
            column_config={
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placa"),
                "Importe": st.column_config.NumberColumn("Sol. Lunes ($)", format="$%.2f"),
                "Real": st.column_config.NumberColumn("Real Lunes ($)", min_value=0.0, step=50.0, format="$%.2f"),
            },
            hide_index=True,
            key="aud_editor_lunes_clean_v12"
        )
        
        if st.button("💾 Guardar Real LUNES", type="secondary", use_container_width=True):
            df_l_guardar = df_lunes.copy()
            df_l_guardar["Real"] = df_aud_l_edit["Real"]
            exito = enviar_datos_hoja(df_l_guardar, hoja="lunes", tipo="real", f_elab=f_elab, f_prog=f_prog)
            if exito:
                st.success("✅ Real del Lunes guardado.")

    with c_aud_j:
        st.markdown("📋 **Comprobación Real JUEVES**")
        df_aud_j_edit = st.data_editor(
            df_jueves[cols_visibles_aud].copy(),
            use_container_width=True,
            height=450,
            disabled=["Vehículo", "Placa", "Importe"],
            column_config={
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placa"),
                "Importe": st.column_config.NumberColumn("Sol. Jueves ($)", format="$%.2f"),
                "Real": st.column_config.NumberColumn("Real Jueves ($)", min_value=0.0, step=50.0, format="$%.2f"),
            },
            hide_index=True,
            key="aud_editor_jueves_clean_v12"
        )
        
        if st.button("💾 Guardar Real JUEVES", type="secondary", use_container_width=True):
            df_j_guardar = df_jueves.copy()
            df_j_guardar["Real"] = df_aud_j_edit["Real"]
            exito = enviar_datos_hoja(df_j_guardar, hoja="jueves", tipo="real", f_elab=f_elab, f_prog=f_prog)
            if exito:
                st.success("✅ Real del Jueves guardado.")

    total_sol_semana = df_aud_l_edit["Importe"].sum() + df_aud_j_edit["Importe"].sum()
    total_real_semana = df_aud_l_edit["Real"].sum() + df_aud_j_edit["Real"].sum()
    ahorro_semana = total_sol_semana - total_real_semana
    saldo_libre_global = PRESUPUESTO_GLOBAL - total_real_semana
    
    st.divider()
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Total Solicitado Semana", f"${total_sol_semana:,.2f}")
    r2.metric("Total Real Comprobado", f"${total_real_semana:,.2f}")
    r3.metric("Ahorro / Remanente", f"${ahorro_semana:,.2f}", delta_color="normal")
    r4.metric("Saldo Libre Total", f"${saldo_libre_global:,.2f}", delta_color="normal")
    
    st.write("")
    if st.button("💾 Archivar Balance en Pestaña 'historico'", type="primary", use_container_width=True):
        with st.spinner("Archivando en pestaña 'historico'..."):
            detalles_txt_lista = []
            for r_num in sorted(MAPEO_SOLICITANTES.keys()):
                r_l = df_lunes[df_lunes["row"] == r_num].iloc[0]
                r_j = df_jueves[df_jueves["row"] == r_num].iloc[0]
                tot_u = r_l["Real"] + r_j["Real"]
                if tot_u > 0:
                    detalles_txt_lista.append(f"{r_l['Vehículo']} - ${tot_u:,.2f}")
            detalles_txt = "; ".join(detalles_txt_lista) if detalles_txt_lista else "Sin cargas reales"
            
            folio_generado = f"SEMANA-{f_prog.strftime('%Y%m%d')}"
            historico_payload = {
                "folio": folio_generado,
                "fecha_elaboro": f_elab.strftime("%d/%m/%Y"),
                "fecha_prog": f_prog.strftime("%d/%m/%Y"),
                "fecha_registro": datetime.now(ZONA_HORARIA).strftime("%d/%m/%Y %I:%M %p"),
                "total_solicitado": float(total_sol_semana),
                "total_ejercido": float(total_real_semana),
                "ahorro": float(ahorro_semana),
                "detalle_unidades": detalles_txt
            }
            
            exito = enviar_datos_hoja(df_jueves, hoja="jueves", tipo="real", f_elab=f_elab, f_prog=f_prog, historico_obj=historico_payload)
            if exito:
                st.success(f"✅ ¡Folio **{folio_generado}** archivado en 'historico'!")
            else:
                st.error("Error al archivar.")

# 6. CIERRE DE VIERNES
with tab_mantenimiento:
    st.markdown("##### 🛠️ Mantenimiento y Cierre de Viernes")
    
    with st.container(border=True):
        st.subheader("🧹 Reestablecer Todo a $0.00 (Cada Viernes)")
        st.write(
            "Usa este botón los **viernes** tras guardar la auditoría semanal. Limpiará ambas pestañas (`lunes` y `jueves`) a **$0.00**, así como el comodín y las transferencias registradas, para arrancar la siguiente semana en ceros."
        )
        
        if st.button("🗑️ Reestablecer Todo para la Próxima Semana", type="secondary", use_container_width=True):
            df_l_limpio = df_lunes.copy()
            df_l_limpio["Operador"] = ""
            df_l_limpio["Importe"] = 0.0
            df_l_limpio["Real"] = 0.0
            
            df_j_limpio = df_jueves.copy()
            df_j_limpio["Operador"] = ""
            df_j_limpio["Importe"] = 0.0
            df_j_limpio["Real"] = 0.0
            
            enviar_datos_hoja(df_l_limpio, hoja="lunes", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
            enviar_datos_hoja(df_l_limpio, hoja="lunes", tipo="real", f_elab=f_elab, f_prog=f_prog)
            enviar_datos_hoja(df_j_limpio, hoja="jueves", tipo="solicitado", f_elab=f_elab, f_prog=f_prog)
            enviar_datos_hoja(df_j_limpio, hoja="jueves", tipo="real", f_elab=f_elab, f_prog=f_prog)
            
            cfg_actual["asignacion_comodin"] = {}
            cfg_actual["transferencias_libres"] = []
            cfg_actual["historial_movimientos"] = []
            cfg_actual["cesion_lian"] = {}
            guardar_config(cfg_actual)
            st.success("✅ ¡Pestañas 'lunes' y 'jueves' reiniciadas a $0.00 y transferencias restablecidas!")
            st.rerun()
