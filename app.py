import streamlit as st
import pandas as pd
import openpyxl
from datetime import date
import io
import json
import os

st.set_page_config(page_title="Control de Combustible", layout="wide", page_icon="⛽")

PRESUPUESTO_OBJETIVO = 3800.00
ARCHIVO_PLANTILLA = "FORMATO.xlsx"
DB_FILE = "registro_semanal.json"

# Catálogo maestro mapeado exactamente a las filas de FORMATO.xlsx (filas 12 a 26)
CATALOGO_OFICIAL = [
    {"row": 12, "Solicita": "COB CHAVEZ NARCISO DEL JESUS", "Encargado": "COB CHAVEZ NARCISO DEL JESUS", "Vehículo": "MOTO SUSUKI", "Placa": "85GWU7", "Importe": 200.00},
    {"row": 13, "Solicita": "PEREZ MAZIN CARLOS EDUARDO", "Encargado": "PEREZ MAZIN CARLOS EDUARDO", "Vehículo": "MOTO SUSUKI", "Placa": "86GWU7", "Importe": 200.00},
    {"row": 14, "Solicita": "DE LA CRUZ PEREZ WILLIAN ARLEY", "Encargado": "DE LA CRUZ PEREZ WILLIAN ARLEY", "Vehículo": "MOTO SUSUKI", "Placa": "86GWU8", "Importe": 200.00},
    {"row": 15, "Solicita": "JESUS COB", "Encargado": "JESUS COB", "Vehículo": "MOTO SUSUKI", "Placa": "87GWU8", "Importe": 0.00},
    {"row": 16, "Solicita": "NOEL CHAN", "Encargado": "", "Vehículo": "MOTO HONDA", "Placa": "88GWU7", "Importe": 0.00},
    {"row": 17, "Solicita": "NOEL CHAN", "Encargado": "", "Vehículo": "MOTO SUSUKI", "Placa": "88GWU8", "Importe": 0.00},
    {"row": 18, "Solicita": "NOEL CHAN", "Encargado": "", "Vehículo": "MOTO SUSUKI", "Placa": "89GWU7", "Importe": 0.00},
    {"row": 19, "Solicita": "NOEL CHAN", "Encargado": "", "Vehículo": "MOTO SUSUKI", "Placa": "89GWU8", "Importe": 0.00},
    {"row": 20, "Solicita": "NOEL CHAN", "Encargado": "", "Vehículo": "MOTO DINAMO", "Placa": "90GWU7", "Importe": 0.00},
    {"row": 21, "Solicita": "NOEL CHAN", "Encargado": "", "Vehículo": "MOTO HONDA", "Placa": "90GWU8", "Importe": 0.00},
    {"row": 22, "Solicita": "LIAN", "Encargado": "LIAN", "Vehículo": "MOTO SUSUKI", "Placa": "91GWU7", "Importe": 150.00},
    {"row": 23, "Solicita": "RENAN/HELDER", "Encargado": "RENAN CETINA", "Vehículo": "AUTOMOVIL JETTA", "Placa": "DFT565C", "Importe": 500.00},
    {"row": 24, "Solicita": "QUEVEDO", "Encargado": "", "Vehículo": "CAMIONETA RAM 701", "Placa": "CN2633B", "Importe": 1000.00},
    {"row": 25, "Solicita": "QUEVEDO", "Encargado": "", "Vehículo": "CAMIONETA NISSAN", "Placa": "CN2632B", "Importe": 0.00},
    {"row": 26, "Solicita": "QUEVEDO", "Encargado": "OMAR OROPEZA", "Vehículo": "MOTOSIERRA 310", "Placa": "00611-23-MOT-49234", "Importe": 0.00},
]

# Diccionario de credenciales (Usuario : Contraseña)
USUARIOS_PASSWORD = {
    "LIAN": "admin123",                      # Administrador General
    "NOEL CHAN": "inspeccion2026",
    "QUEVEDO": "ambiental2026",
    "RENAN/HELDER": "urbano2026",
    "COB CHAVEZ NARCISO DEL JESUS": "notif123",
    "PEREZ MAZIN CARLOS EDUARDO": "notif123",
    "DE LA CRUZ PEREZ WILLIAN ARLEY": "notif123",
    "JESUS COB": "notif123"
}

# Funciones de persistencia de datos compartidos
def cargar_datos():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return pd.DataFrame(json.load(f))
        except Exception:
            return pd.DataFrame(CATALOGO_OFICIAL)
    return pd.DataFrame(CATALOGO_OFICIAL)

def guardar_datos(df):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, indent=4, ensure_ascii=False)

# Inicializar sesión
if "usuario_logueado" not in st.session_state:
    st.session_state.usuario_logueado = None

# ==========================================
# 1. PANTALLA DE LOGIN
# ==========================================
if st.session_state.usuario_logueado is None:
    st.title("⛽ Sistema de Solicitud de Combustible")
    st.caption("Dirección de Desarrollo Urbano y Medio Ambiente")
    
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        with st.form("form_login"):
            st.subheader("🔐 Inicio de Sesión")
            usuario = st.selectbox("Selecciona tu Usuario / Solicitante", list(USUARIOS_PASSWORD.keys()))
            password = st.text_input("Contraseña", type="password")
            btn_login = st.form_submit_button("Ingresar al Sistema", use_container_width=True)
            
            if btn_login:
                if password == USUARIOS_PASSWORD.get(usuario):
                    st.session_state.usuario_logueado = usuario
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta. Intenta nuevamente.")
    st.stop()

# ==========================================
# 2. ENCABEZADO Y SESIÓN ACTIVA
# ==========================================
usuario_actual = st.session_state.usuario_logueado
es_admin = (usuario_actual == "LIAN")

col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.title("⛽ Solicitud Semanal de Combustible")
    rol_label = "👑 **ADMINISTRADOR GENERAL**" if es_admin else f"👤 Solicitante: **{usuario_actual}**"
    st.markdown(rol_label)
with col_head2:
    st.write("")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.usuario_logueado = None
        st.rerun()

df_global = cargar_datos()

# ==========================================
# 3. VISTA SOLICITANTE
# ==========================================
if not es_admin:
    st.info("👋 Asigna el **Nombre del Encargado** que conducirá la unidad esta semana e ingresa el **Importe ($)** correspondiente.")
    
    # Filtrar solo los vehículos pertenecientes al usuario conectado
    filas_usuario = df_global[df_global["Solicita"] == usuario_actual].copy()
    
    df_user_editado = st.data_editor(
        filas_usuario,
        use_container_width=True,
        disabled=["row", "Solicita", "Vehículo", "Placa"],
        column_config={
            "Encargado": st.column_config.TextColumn(
                "Nombre del Encargado / Operador",
                help="Nombre de quien conducirá y cargará la unidad",
                required=True
            ),
            "Importe": st.column_config.NumberColumn(
                "Importe ($)",
                min_value=0.0,
                step=50.0,
                format="$%.2f"
            ),
            "row": None,
            "Solicita": None
        },
        hide_index=True,
        key="editor_solicitante"
    )
    
    subtotal = df_user_editado["Importe"].sum()
    st.metric("Total Solicitado por tu Área", f"${subtotal:,.2f} MXN")
    
    if st.button("💾 Guardar Mi Solicitud", type="primary", use_container_width=True):
        # Actualizar los registros modificados en la base global
        for _, row_edit in df_user_editado.iterrows():
            idx = df_global[df_global["row"] == row_edit["row"]].index
            df_global.loc[idx, "Encargado"] = row_edit["Encargado"]
            df_global.loc[idx, "Importe"] = row_edit["Importe"]
            
        guardar_datos(df_global)
        st.success("✅ Tu solicitud semanal ha sido registrada y enviada a Administración.")

# ==========================================
# 4. VISTA ADMINISTRADOR (LIAN)
# ==========================================
else:
    st.subheader("⚙️ Panel de Consolidación y Aprobación Administrativa")
    
    # Selector de fechas oficiales
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        fecha_elaboro = st.date_input("Fecha de Elaboración (Oficio)", value=date.today())
    with col_f2:
        fecha_programacion = st.date_input("Programación para el día", value=date.today())
        
    st.markdown("##### 📋 Tabla Consolidada General")
    st.caption("Puedes corregir operadores o importes directamente si es necesario:")
    
    df_admin_editado = st.data_editor(
        df_global,
        use_container_width=True,
        disabled=["row", "Vehículo", "Placa"],
        column_config={
            "Solicita": st.column_config.TextColumn("Solicitante", disabled=True),
            "Encargado": st.column_config.TextColumn("Operador / Encargado"),
            "Importe": st.column_config.NumberColumn("Importe ($)", min_value=0.0, step=50.0, format="$%.2f"),
            "row": None
        },
        hide_index=True,
        key="editor_admin"
    )
    
    # Cálculos y semáforo de presupuesto semanal
    total_asignado = df_admin_editado["Importe"].sum()
    diferencia = PRESUPUESTO_OBJETIVO - total_asignado
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Presupuesto Semanal", f"${PRESUPUESTO_OBJETIVO:,.2f}")
    col_m2.metric("Total Solicitado Global", f"${total_asignado:,.2f}", delta=f"{total_asignado - PRESUPUESTO_OBJETIVO:,.2f}", delta_color="inverse")
    col_m3.metric("Saldo Disponible", f"${diferencia:,.2f}", delta_color="normal" if diferencia >= 0 else "off")
    
    # Validaciones visuales
    if total_asignado > PRESUPUESTO_OBJETIVO:
        st.error(f"❌ El monto total excede el presupuesto autorizado por **${abs(diferencia):,.2f} MXN**.")
    elif total_asignado == PRESUPUESTO_OBJETIVO:
        st.success("✅ Presupuesto distribuido exactamente al 100% ($3,800.00 MXN).")
    else:
        st.warning(f"ℹ️ Aún tienes un saldo disponible de **${diferencia:,.2f} MXN** sin asignar.")

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Guardar Cambios Administrativos", use_container_width=True):
            guardar_datos(df_admin_editado)
            st.success("✅ Cambios guardados correctamente.")

    # Generación y descarga del Excel oficial
    def generar_excel_oficial(df_datos, f_elab, f_prog):
        wb = openpyxl.load_workbook(ARCHIVO_PLANTILLA)
        ws = wb.active
        
        # Inyectar fechas en las celdas del encabezado
        ws["I8"] = f_elab.strftime("%d/%m/%Y")
        ws["I9"] = f_prog.strftime("%d/%m/%Y")
        
        # Inyectar encargados e importes
        for _, fila in df_datos.iterrows():
            r = int(fila["row"])
            encargado_val = str(fila["Encargado"]).strip() if pd.notna(fila["Encargado"]) and str(fila["Encargado"]).strip() != "" else None
            importe_val = float(fila["Importe"]) if pd.notna(fila["Importe"]) and float(fila["Importe"]) > 0 else None
            
            ws.cell(row=r, column=2, value=encargado_val)
            ws.cell(row=r, column=7, value=importe_val)
            
        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    with col_btn2:
        try:
            archivo_generado = generar_excel_oficial(df_admin_editado, fecha_elaboro, fecha_programacion)
            st.download_button(
                label="📥 Descargar Formato Oficial (.xlsx)",
                data=archivo_generado,
                file_name=f"SOLICITUD_COMBUSTIBLE_{fecha_programacion.strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except FileNotFoundError:
            st.error("⚠️ No se encontró el archivo `FORMATO.xlsx` en la carpeta.")
