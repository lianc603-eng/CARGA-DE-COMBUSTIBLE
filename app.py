import streamlit as st
import pandas as pd
from datetime import datetime, date, time
import pytz
import requests
import json

st.set_page_config(page_title="Control de Combustible", layout="wide", page_icon="⛽")

PRESUPUESTO_GLOBAL = 3800.00
HORA_LIMITE = time(15, 0)  # 3:00 PM
ZONA_HORARIA = pytz.timezone("America/Merida")

WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzOjgha2Zjyog01t6LmA_R--EB4Ecqv2ifO_i2YJbLRLbXGShbu5uzFVi85FUTGplM8/exec"

PRESUPUESTO_POR_SOLICITANTE = {
    "COB CHAVEZ NARCISO DEL JESUS": 400.00,
    "PEREZ MAZIN CARLOS EDUARDO": 200.00,
    "DE LA CRUZ PEREZ WILLIAN ARLEY": 200.00,
    "NOEL CHAN": 850.00,
    "LIAN": 150.00,
    "RENAN/HELDER": 500.00,
    "QUEVEDO": 1500.00,
}

MAPEO_SOLICITANTES = {
    12: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "85GWU7"},
    13: {"solicita": "PEREZ MAZIN CARLOS EDUARDO", "vehiculo": "MOTO SUSUKI", "placa": "86GWU7"},
    14: {"solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "vehiculo": "MOTO SUSUKI", "placa": "86GWU8"},
    15: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "87GWU8"},
    16: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "88GWU7"},
    17: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "88GWU8"},
    18: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU7"},
    19: {"solicita": "NOEL CHAN", "vehiculo": "MOTO SUSUKI", "placa": "89GWU8"},
    20: {"solicita": "NOEL CHAN", "vehiculo": "MOTO DINAMO", "placa": "90GWU7"},
    21: {"solicita": "NOEL CHAN", "vehiculo": "MOTO HONDA", "placa": "90GWU8"},
    22: {"solicita": "LIAN", "vehiculo": "MOTO SUSUKI", "placa": "91GWU7"},
    23: {"solicita": "RENAN/HELDER", "vehiculo": "AUTOMOVIL JETTA", "placa": "DFT565C"},
    24: {"solicita": "QUEVEDO", "vehiculo": "CAMIONETA RAM 701", "placa": "CN2633B"},
    25: {"solicita": "QUEVEDO", "vehiculo": "CAMIONETA NISSAN", "placa": "CN2632B"},
    26: {"solicita": "QUEVEDO", "vehiculo": "MOTOSIERRA 310", "placa": "00611-23-MOT-49234"},
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

def obtener_datos_sheets():
    try:
        res = requests.get(WEBHOOK_URL, timeout=15)
        if res.status_code == 200:
            datos_raw = res.json().get("data", [])
            for item in datos_raw:
                r = int(item["row"])
                if r in MAPEO_SOLICITANTES:
                    item["Solicita"] = MAPEO_SOLICITANTES[r]["solicita"]
                    item["Vehículo"] = MAPEO_SOLICITANTES[r]["vehiculo"]
                    item["Placa"] = MAPEO_SOLICITANTES[r]["placa"]
                    item["Encargado"] = str(item.get("encargado", "")).strip()
                    try:
                        item["Importe"] = float(item.get("importe", 0.0))
                    except (ValueError, TypeError):
                        item["Importe"] = 0.0
            return pd.DataFrame(datos_raw)
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
            "encargado": str(fila["Encargado"]).strip() if pd.notna(fila["Encargado"]) else "",
            "importe": float(fila["Importe"]) if pd.notna(fila["Importe"]) else 0.0
        })
        
    try:
        res = requests.post(WEBHOOK_URL, json=payload, timeout=15)
        return res.status_code == 200
    except Exception as e:
        st.error(f"Error al enviar datos: {e}")
        return False

# ==========================================
# 1. INICIO DE SESIÓN
# ==========================================
if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None

if st.session_state.usuario_logueado is None:
    st.title("⛽ Sistema de Solicitud de Combustible")
    st.caption("Dirección de Desarrollo Urbano y Medio Ambiente")
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.form("form_login"):
            st.subheader("🔐 Iniciar Sesión")
            usr = st.selectbox("Selecciona tu Usuario / Solicitante", list(USUARIOS_PASSWORD.keys()))
            pwd = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Ingresar al Sistema", use_container_width=True):
                if pwd == USUARIOS_PASSWORD.get(usr):
                    st.session_state.usuario_logueado = usr
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta.")
    st.stop()

# ==========================================
# 2. ENCABEZADO Y REVISIÓN DE HORARIO
# ==========================================
usuario = st.session_state.usuario_logueado
es_admin = (usuario == "LIAN")

# Evaluación del límite de horario
ahora_local = datetime.now(ZONA_HORARIA)
hora_actual = ahora_local.time()
sistema_bloqueado = (hora_actual >= HORA_LIMITE) and not es_admin

c1, c2 = st.columns([4, 1])
with c1:
    st.title("⛽ Control Semanal de Combustible")
    st.markdown("👑 **ADMINISTRADOR GENERAL**" if es_admin else f"👤 Solicitante: **{usuario}**")
    st.caption(f"🕒 Hora local: **{ahora_local.strftime('%I:%M %p')}** | Límite de carga: **3:00 PM**")
with c2:
    st.write("")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.usuario_logueado = None
        st.rerun()

df_actual = obtener_datos_sheets()

if df_actual.empty:
    st.warning("Cargando datos desde Google Sheets...")
    st.stop()

# ==========================================
# 3. VISTA SOLICITANTE
# ==========================================
if not es_admin:
    presupuesto_propio = PRESUPUESTO_POR_SOLICITANTE.get(usuario, 0.00)
    df_solicitante = df_actual[df_actual["Solicita"] == usuario].copy()
    
    if sistema_bloqueado:
        st.error("🔒 **SISTEMA CERRADO POR HORARIO LIMITE (3:00 PM)**. La captura está deshabilitada. Contacta al Administrador para cualquier modificación.")
    
    tab_solicitud, tab_resumen = st.tabs(["📝 Captura y Distribución", "📊 Resumen y Saldo de Cargas"])
    
    with tab_solicitud:
        if not sistema_bloqueado:
            st.caption("Captura el nombre del encargado y el importe para las unidades bajo tu cargo:")
        
        columnas_bloqueadas = ["row", "Solicita", "Vehículo", "Placa", "encargado", "importe"]
        if sistema_bloqueado:
            columnas_bloqueadas.extend(["Encargado", "Importe"])
        
        df_edit = st.data_editor(
            df_solicitante,
            use_container_width=True,
            disabled=columnas_bloqueadas,
            column_config={
                "Encargado": st.column_config.TextColumn("Nombre del Encargado / Operador", required=True),
                "Importe": st.column_config.NumberColumn("Importe ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "row": None, "encargado": None, "importe": None, "Solicita": None
            },
            hide_index=True,
            key="editor_usuario"
        )
        
        total_solicitado = df_edit["Importe"].sum()
        saldo_disponible = presupuesto_propio - total_solicitado
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Presupuesto Asignado a tu Área", f"${presupuesto_propio:,.2f}")
        m2.metric("Total Distribuido en Unidades", f"${total_solicitado:,.2f}", delta=f"{total_solicitado - presupuesto_propio:,.2f}", delta_color="inverse")
        m3.metric("Saldo Disponible Restante", f"${saldo_disponible:,.2f}", delta_color="normal" if saldo_disponible >= 0 else "off")
        
        if total_solicitado > presupuesto_propio:
            st.error(f"❌ Estás excediendo tu presupuesto autorizado por **${abs(saldo_disponible):,.2f} MXN**.")
        elif total_solicitado == presupuesto_propio:
            st.success("✅ Has distribuido exactamente el 100% de tu presupuesto semanal.")
        else:
            st.info(f"ℹ️ Tienes **${saldo_disponible:,.2f} MXN** disponibles para asignar.")
            
        if not sistema_bloqueado:
            if st.button("💾 Guardar Solicitud en Google Sheets", type="primary", use_container_width=True):
                with st.spinner("Guardando en Google Sheets..."):
                    if guardar_en_sheets(df_edit):
                        st.success("✅ Solicitud guardada y sincronizada correctamente.")
                    else:
                        st.error("❌ Error al guardar en Google Sheets.")

    with tab_resumen:
        st.subheader("🔍 Desglose de Cargas Solicitadas")
        cargas_activas = df_edit[df_edit["Importe"] > 0][["Vehículo", "Placa", "Encargado", "Importe"]]
        
        if not cargas_activas.empty:
            st.dataframe(cargas_activas, use_container_width=True, hide_index=True)
        else:
            st.warning("Aún no se ha asignado presupuesto a ninguna unidad.")
            
        st.markdown("---")
        st.markdown(f"**Resumen de cuenta:** Presupuesto: **${presupuesto_propio:,.2f}** | Asignado: **${total_solicitado:,.2f}** | Saldo Restante: **${saldo_disponible:,.2f}**")

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Consolidación General y Control de Saldos por Área")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración", value=date.today())
    with col_f2:
        f_prog = st.date_input("Programación para el día", value=date.today())

    tab_gral, tab_saldos = st.tabs(["📋 Tabla Consolidada General", "📊 Monitoreo de Presupuesto por Solicitante"])

    with tab_gral:
        df_admin_edit = st.data_editor(
            df_actual,
            use_container_width=True,
            disabled=["row", "Vehículo", "Placa", "encargado", "importe"],
            column_config={
                "Solicita": st.column_config.TextColumn("Solicitante", disabled=True),
                "Encargado": st.column_config.TextColumn("Operador / Encargado"),
                "Importe": st.column_config.NumberColumn("Importe ($)", min_value=0.0, step=50.0, format="$%.2f"),
                "row": None, "encargado": None, "importe": None
            },
            hide_index=True,
            key="editor_admin"
        )
        
        total_global = df_admin_edit["Importe"].sum()
        saldo_global = PRESUPUESTO_GLOBAL - total_global
        
        gm1, gm2, gm3 = st.columns(3)
        gm1.metric("Presupuesto Global Autorizado", f"${PRESUPUESTO_GLOBAL:,.2f}")
        gm2.metric("Total Solicitado General", f"${total_global:,.2f}", delta=f"{total_global - PRESUPUESTO_GLOBAL:,.2f}", delta_color="inverse")
        gm3.metric("Saldo Global Restante", f"${saldo_global:,.2f}", delta_color="normal" if saldo_global >= 0 else "off")
        
        if st.button("💾 Sincronizar y Guardar Todo en Google Sheets", type="primary", use_container_width=True):
            with st.spinner("Actualizando Google Sheets..."):
                if guardar_en_sheets(df_admin_edit, f_elab, f_prog):
                    st.success("✅ Hoja de cálculo actualizada con todas las cargas y fechas oficiales.")
                else:
                    st.error("❌ Error al guardar.")

    with tab_saldos:
        st.subheader("Estado Presupuestal por Solicitante")
        
        resumen_solicitantes = []
        for sol, p_base in PRESUPUESTO_POR_SOLICITANTE.items():
            filas_sol = df_admin_edit[df_admin_edit["Solicita"] == sol]
            gastado = filas_sol["Importe"].sum()
            restante = p_base - gastado
            unidades_cargando = len(filas_sol[filas_sol["Importe"] > 0])
            
            resumen_solicitantes.append({
                "Solicitante": sol,
                "Presupuesto Asignado": f"${p_base:,.2f}",
                "Total Solicitado": f"${gastado:,.2f}",
                "Saldo Disponible": f"${restante:,.2f}",
                "Unidades con Carga": f"{unidades_cargando} unidad(es)",
                "Estado": "✅ Exacto" if restante == 0 else ("⚠️ Excedido" if restante < 0 else "🟢 Con Saldo")
            })
            
        df_resumen = pd.DataFrame(resumen_solicitantes)
        st.dataframe(df_resumen, use_container_width=True, hide_index=True)
