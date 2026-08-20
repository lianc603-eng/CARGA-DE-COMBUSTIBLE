import streamlit as st
import pandas as pd
from datetime import datetime, date, time
import pytz
import requests
import json
import io
import os
import openpyxl
from fpdf import FPDF

st.set_page_config(page_title="Control de Combustible", layout="wide", page_icon="⛽")

PRESUPUESTO_GLOBAL = 3800.00
HORA_LIMITE = time(15, 0)  # 3:00 PM
ZONA_HORARIA = pytz.timezone("America/Merida")
CONFIG_FILE = "config_sistema.json"
AVISOS_FILE = "avisos_sistema.json"

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

# --- GESTIÓN DE CONFIGURACIÓN Y MENSAJES ---
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
# GENERADORES DE ARCHIVOS OFICIALES
# ==========================================
def generar_excel_filtrado(df_cargas, f_elab, f_prog):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Solicitud de Carga"
    
    ws["B2"] = "H. AYUNTAMIENTO DE CAMPECHE"
    ws["B4"] = "Unidad: DIRECCION DE DESARROLLO URBANO Y MEDIO AMBIENTE"
    ws["F6"] = f"Elaboró: {f_elab.strftime('%d/%m/%Y')}"
    ws["F7"] = f"Programación para el día: {f_prog.strftime('%d/%m/%Y')}"
    
    headers = ["NOMBRE DEL ENCARGADO", "VEHÍCULO", "PLACA", "OFICIAL/COMODATO", "IMPORTE", "COMBUSTIBLE", "ACTIVIDAD"]
    ws.append([])
    ws.append(headers)
    
    total = 0.0
    for _, row in df_cargas.iterrows():
        ws.append([
            row["Operador / Encargado"],
            row["Vehículo"],
            row["Placa"],
            "OFICIAL",
            float(row["Importe ($)"]),
            "MAGNA",
            row["Actividad"]
        ])
        total += float(row["Importe ($)"])
        
    ws.append(["", "", "", "TOTAL:", total, "", ""])
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
        estado_horario += " (⚡ Desbloqueo temporal de Admin activo)"
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
# 3. VISTA SOLICITANTE (TARJETAS MÓVILES + AVISOS)
# ==========================================
if not es_admin:
    # Despliegue de avisos dirigidos al solicitante o generales
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
    st.subheader("⚙️ Panel de Consolidación, Auditoría y Gestión")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración", value=date.today())
    with col_f2:
        f_prog = st.date_input("Programación para el día", value=date.today())

    tab_saldos, tab_cargas_activas, tab_confirmacion, tab_avisos, tab_edicion, tab_mantenimiento = st.tabs([
        "📊 Saldos por Área",
        "📄 Solicitud Final", 
        "✅ Auditoría Real",
        "📢 Centro de Avisos",
        "✏️ Editor General",
        "🛠️ Modo Pruebas"
    ])

    df_solo_cargas = df_actual[df_actual["Importe ($)"] > 0].copy()

    with tab_saldos:
        filas_reporte = []
        for sol, p_base in PRESUPUESTO_POR_SOLICITANTE.items():
            sub_df = df_actual[df_actual["Solicitante"] == sol]
            solicitado = sub_df["Importe ($)"].sum()
            disponible = p_base - solicitado
            pct_usado = (solicitado / p_base * 100) if p_base > 0 else 0
            
            filas_reporte.append({
                "Solicitante": sol,
                "Presupuesto Base": p_base,
                "Monto Solicitado": solicitado,
                "Saldo Disponible": disponible,
                "% Ejercido": f"{pct_usado:.1f}%",
                "Unidades Activas": f"{len(sub_df[sub_df['Importe ($)'] > 0])} de {len(sub_df)}",
                "Estatus": "✅ 100% Ejercido" if disponible == 0 else ("⚠️ Excedido" if disponible < 0 else "🟢 Con Saldo")
            })
            
        total_global = df_actual["Importe ($)"].sum()
        saldo_global = PRESUPUESTO_GLOBAL - total_global
        
        c_m1, c_m2, c_m3, c_m4 = st.columns(4)
        c_m1.metric("Presupuesto Global", f"${PRESUPUESTO_GLOBAL:,.2f}")
        c_m2.metric("Total Distribuido", f"${total_global:,.2f}", delta=f"{total_global - PRESUPUESTO_GLOBAL:,.2f}", delta_color="inverse")
        c_m3.metric("Saldo Disponible Global", f"${saldo_global:,.2f}", delta_color="normal" if saldo_global >= 0 else "off")
        c_m4.metric("Bolsa Comodín", "$200.00")
        
        st.dataframe(
            pd.DataFrame(filas_reporte),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Presupuesto Base": st.column_config.NumberColumn(format="$%.2f"),
                "Monto Solicitado": st.column_config.NumberColumn(format="$%.2f"),
                "Saldo Disponible": st.column_config.NumberColumn(format="$%.2f"),
            }
        )

    with tab_cargas_activas:
        st.markdown("##### 🚗 Lista Oficial de Unidades a Cargar")
        
        if df_solo_cargas.empty:
            st.warning("⚠️ No hay vehículos con monto asignado para generar el reporte.")
        else:
            st.dataframe(
                df_solo_cargas[["Operador / Encargado", "Vehículo", "Placa", "Importe ($)", "Actividad"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Importe ($)": st.column_config.NumberColumn(format="$%.2f"),
                    "Actividad": st.column_config.TextColumn("Actividad Oficial", width="large")
                }
            )
            
            st.markdown("---")
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                excel_bytes = generar_excel_filtrado(df_solo_cargas, f_elab, f_prog)
                st.download_button(
                    label="📥 Descargar Reporte en Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"SOLICITUD_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col_d2:
                pdf_bytes = generar_pdf_oficial(df_solo_cargas, f_elab, f_prog)
                st.download_button(
                    label="📄 Descargar Oficio Oficial en PDF",
                    data=pdf_bytes,
                    file_name=f"OFICIO_COMBUSTIBLE_{f_prog.strftime('%d%m%Y')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

    with tab_confirmacion:
        st.markdown("##### ✅ Validación y Conciliación de Cargas Realizadas")
        st.caption("Verifica si las unidades cargaron su monto completo. Ajusta si cargaron menos para recalcular saldos:")
        
        df_auditoria = df_actual.copy()
        df_audit_edit = st.data_editor(
            df_auditoria,
            use_container_width=True,
            disabled=["row", "Solicitante", "Vehículo", "Placa", "Actividad"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Área / Solicitante"),
                "Vehículo": st.column_config.TextColumn("Vehículo"),
                "Placa": st.column_config.TextColumn("Placas"),
                "Operador / Encargado": st.column_config.TextColumn("Operador"),
                "Importe ($)": st.column_config.NumberColumn("Monto Realmente Cargado ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "row": None, "Actividad": None
            },
            hide_index=True,
            key="editor_auditoria"
        )
        
        total_real_ejercido = df_audit_edit["Importe ($)"].sum()
        remanente_recuperado = PRESUPUESTO_GLOBAL - total_real_ejercido
        
        st.markdown("---")
        ar1, ar2, ar3 = st.columns(3)
        ar1.metric("Total Real Ejercido", f"${total_real_ejercido:,.2f}")
        ar2.metric("Saldo No Ejercido / Recuperado", f"${remanente_recuperado:,.2f}", delta_color="normal")
        ar3.metric("Unidades con Carga Real", f"{len(df_audit_edit[df_audit_edit['Importe ($)'] > 0])} unidades")
        
        if st.button("💾 Confirmar Cargas Reales y Guardar Registro Final", type="primary", use_container_width=True):
            with st.spinner("Actualizando Google Sheets..."):
                if guardar_en_sheets(df_audit_edit, f_elab, f_prog):
                    st.success("✅ Cargas confirmadas y saldos recalculados correctamente.")
                    st.rerun()
                else:
                    st.error("❌ Error al actualizar los registros.")

    # NUEVO: Pestaña de Centro de Avisos y Notificaciones
    with tab_avisos:
        st.markdown("##### 📢 Publicar Comunicados y Mensajes Personalizados")
        
        with st.form("form_nuevo_aviso"):
            c_dest, c_tipo = st.columns([2, 1])
            with c_dest:
                destinatarios_opciones = ["TODOS"] + [u for u in USUARIOS_PASSWORD.keys() if u != "LIAN"]
                destinatario_sel = st.selectbox("Dirigido a:", destinatarios_opciones)
            with c_tipo:
                tipo_aviso = st.selectbox("Nivel de Prioridad:", ["Informativo", "Importante", "Urgente"])
                
            texto_aviso_nuevo = st.text_area("Contenido del Mensaje o Instrucción:", placeholder="Ej. Recuerden verificar el kilometraje antes de solicitar la carga...")
            
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

    with tab_edicion:
        df_admin_edit = st.data_editor(
            df_actual,
            use_container_width=True,
            disabled=["row", "Vehículo", "Placa"],
            column_config={
                "Solicitante": st.column_config.TextColumn("Solicitante", disabled=False),
                "Operador / Encargado": st.column_config.TextColumn("Operador / Encargado"),
                "Importe ($)": st.column_config.NumberColumn("Importe ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "Actividad": st.column_config.TextColumn("Actividad", width="large"),
                "row": None
            },
            hide_index=True,
            key="editor_admin"
        )
        
        if st.button("💾 Sincronizar y Guardar Todo en Google Sheets", type="primary", use_container_width=True):
            with st.spinner("Actualizando Google Sheets..."):
                if guardar_en_sheets(df_admin_edit, f_elab, f_prog):
                    st.success("✅ Datos sincronizados correctamente.")
                    st.rerun()

    with tab_mantenimiento:
        st.markdown("##### 🛠️ Control de Horarios y Simulación de Pruebas")
        
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
