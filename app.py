import streamlit as st
import pandas as pd
from datetime import date
import requests
import json

st.set_page_config(page_title="Control de Combustible", layout="wide", page_icon="⛽")

PRESUPUESTO_OBJETIVO = 3800.00

# URL de conexión con Google Apps Script
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbzOjgha2Zjyog01t6LmA_R--EB4Ecqv2ifO_i2YJbLRLbXGShbu5uzFVi85FUTGplM8/exec"

# Mapeo oficial de vehículos y solicitantes por fila (filas 12 a 26 del formato)
MAPEO_SOLICITANTES = {
    12: {"solicita": "COB CHAVEZ NARCISO DEL JESUS", "vehiculo": "MOTO SUSUKI", "placa": "85GWU7"},
    13: {"solicita": "PEREZ MAZIN CARLOS EDUARDO", "vehiculo": "MOTO SUSUKI", "placa": "86GWU7"},
    14: {"solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "vehiculo": "MOTO SUSUKI", "placa": "86GWU8"},
    15: {"solicita": "JESUS COB", "vehiculo": "MOTO SUSUKI", "placa": "87GWU8"},
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

# Credenciales de acceso por usuario
USUARIOS_PASSWORD = {
    "LIAN": "admin123",
    "NOEL CHAN": "inspeccion2026",
    "QUEVEDO": "ambiental2026",
    "RENAN/HELDER": "urbano2026",
    "COB CHAVEZ NARCISO DEL JESUS": "notif123",
    "PEREZ MAZIN CARLOS EDUARDO": "notif123",
    "DE LA CRUZ PEREZ WILLIAN ARLEY": "notif123",
    "JESUS COB": "notif123"
}

# Funciones de comunicación con Google Sheets
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
        st.error(f"Error al conectar con Google Sheets: {err}")
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
        st.error(f"Error al guardar los datos: {e}")
        return False

# ==========================================
# 1. GESTIÓN DE SESIÓN Y LOGIN
# ==========================================
if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None

if st.session_state.usuario_logueado is None:
    st.title("⛽ Sistema de Solicitud de Combustible")
    st.caption("H. Ayuntamiento de Campeche — Dirección de Desarrollo Urbano y Medio Ambiente")
    
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
# 2. ENCABEZADO Y BARRA SUPERIOR
# ==========================================
usuario = st.session_state.usuario_logueado
es_admin = (usuario == "LIAN")

c1, c2 = st.columns([4, 1])
with c1:
    st.title("⛽ Solicitud Semanal de Combustible")
    st.markdown("👑 **ADMINISTRADOR GENERAL**" if es_admin else f"👤 Solicitante: **{usuario}**")
with c2:
    st.write("")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.usuario_logueado = None
        st.rerun()

# Cargar los datos en tiempo real desde Google Sheets
df_actual = obtener_datos_sheets()

if df_actual.empty:
    st.warning("Conectando con la base de datos en Google Sheets...")
    st.stop()

# ==========================================
# 3. VISTA SOLICITANTE
# ==========================================
if not es_admin:
    st.info("👋 Asigna el **Nombre del Encargado** que conducirá la unidad esta semana e ingresa el **Importe ($)** correspondiente:")
    
    df_solicitante = df_actual[df_actual["Solicita"] == usuario].copy()
    
    df_edit = st.data_editor(
        df_solicitante,
        use_container_width=True,
        disabled=["row", "Solicita", "Vehículo", "Placa", "encargado", "importe"],
        column_config={
            "Encargado": st.column_config.TextColumn("Nombre del Encargado / Operador", required=True),
            "Importe": st.column_config.NumberColumn("Importe ($)", min_value=0.0, step=50.0, format="$%.2f"),
            "row": None, "encargado": None, "importe": None, "Solicita": None
        },
        hide_index=True
    )
    
    st.metric("Total Solicitado por tu Área", f"${df_edit['Importe'].sum():,.2f} MXN")
    
    if st.button("💾 Guardar Solicitud en Google Sheets", type="primary", use_container_width=True):
        with st.spinner("Guardando en Google Sheets..."):
            if guardar_en_sheets(df_edit):
                st.success("✅ ¡Datos guardados correctamente en la hoja de cálculo oficial!")
            else:
                st.error("❌ No se pudo guardar la información.")

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Consolidación General y Control Presupuestal")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_elab = st.date_input("Fecha de Elaboración (Oficio)", value=date.today())
    with col_f2:
        f_prog = st.date_input("Programación para el día", value=date.today())
        
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
        hide_index=True
    )
    
    total = df_admin_edit["Importe"].sum()
    saldo = PRESUPUESTO_OBJETIVO - total
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Presupuesto Semanal", f"${PRESUPUESTO_OBJETIVO:,.2f}")
    m2.metric("Total Distribuido", f"${total:,.2f}", delta=f"{total - PRESUPUESTO_OBJETIVO:,.2f}", delta_color="inverse")
    m3.metric("Saldo Disponible", f"${saldo:,.2f}", delta_color="normal" if saldo >= 0 else "off")
    
    if total > PRESUPUESTO_OBJETIVO:
        st.error(f"❌ Exceso de presupuesto por **${abs(saldo):,.2f} MXN**.")
    elif total == PRESUPUESTO_OBJETIVO:
        st.success("✅ Presupuesto distribuido al 100% ($3,800.00 MXN).")
    else:
        st.warning(f"ℹ️ Saldo por asignar: **${saldo:,.2f} MXN**.")
        
    if st.button("💾 Sincronizar y Guardar Todo en Google Sheets", type="primary", use_container_width=True):
        with st.spinner("Actualizando Google Sheets..."):
            if guardar_en_sheets(df_admin_edit, f_elab, f_prog):
                st.success("✅ Hoja de cálculo actualizada con todas las cargas y fechas oficiales.")
            else:
                st.error("❌ Error al sincronizar con Google Sheets.")
