import streamlit as st
import pandas as pd
import os

# 1. Configuración principal de la página web
st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

# Rutas de trabajo (Detecta automáticamente la carpeta donde está este script)
CARPETA = os.getcwd() 
ARCHIVO_SALIDA = 'Reporte_Consolidado_Final.xlsx'
RUTA_COMPLETA_EXCEL = os.path.join(CARPETA, ARCHIVO_SALIDA)

# ==========================================
# MENÚ LATERAL (CONTROL Y LOGO)
# ==========================================
# Busca el logo.jpg en la carpeta actual
logo_path = os.path.join(CARPETA, "logo.jpg")

if os.path.exists(logo_path):
    st.sidebar.image(logo_path, use_container_width=True)
else:
    st.sidebar.warning(f"⚠️ No se encontró 'logo.jpg' en: {CARPETA}")

st.sidebar.header("Panel de Control")

def consolidar_archivos():
    dataframes = []
    # Busca archivos en la carpeta actual
    for archivo in os.listdir(CARPETA):
        if archivo.endswith('.xlsx') and archivo != ARCHIVO_SALIDA:
            try:
                ruta_completa = os.path.join(CARPETA, archivo)
                df_temp = pd.read_excel(ruta_completa)
                df_temp['Origen'] = archivo
                dataframes.append(df_temp)
            except Exception as e:
                st.sidebar.error(f"Error leyendo {archivo}: {e}")

    if dataframes:
        try:
            reporte_final = pd.concat(dataframes, ignore_index=True)
            with pd.ExcelWriter(RUTA_COMPLETA_EXCEL, engine='xlsxwriter') as writer:
                reporte_final.to_excel(writer, sheet_name='DATA', index=False)
                pd.DataFrame().to_excel(writer, sheet_name='DASHBOARD')
            return True
        except PermissionError:
            st.sidebar.error("❌ ¡ERROR! Cierra 'Reporte_Consolidado_Final.xlsx' e intenta de nuevo.")
            return False
    return False

if st.sidebar.button("🔄 Consolidar Nuevos Reportes", use_container_width=True):
    with st.spinner("Procesando archivos..."):
        if consolidar_archivos():
            st.sidebar.success("¡Consolidación exitosa!")
            st.cache_data.clear()
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.info("💡 Recuerda dar clic en Consolidar Nuevos Reportes si has modificado tus Excels.")

# ==========================================
# CARGA DE DATOS Y RENDERIZADO DEL DASHBOARD
# ==========================================
@st.cache_data(ttl=600)
def cargar_datos():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQeRzx7jkJ7S1F-5SzuKG35U8llKKTZ3QxlMyR5rzlN96vANkHWHF4wMcH4eYFt673J9LUnBEoUdXNG/pub?output=csv"
    df = pd.read_csv(url)
    return df

df_master = cargar_datos()

if df_master.empty:
    st.warning("⚠️ No se encontraron datos consolidados.")
else:
    st.title("📊 Reporte Semanal - Transportes Calpi")
    st.markdown("---")

    if 'Origen' in df_master.columns:
        df_master['Categoria'] = df_master['Origen'].str.replace('.xlsx', '', regex=False).str.strip().str.upper()
    else:
        df_master['Categoria'] = 'GENERAL'

    categorias_disponibles = df_master['Categoria'].dropna().unique()

    totales_por_categoria = {}
    for cat in categorias_disponibles:
        df_cat = df_master[df_master['Categoria'] == cat]
        columnas_reales = [c for c in df_cat.columns if c not in ['Origen', 'Categoria']]
        col_dinero = next((col for col in columnas_reales if "$$" in str(col)), None)
        
        if col_dinero:
            total = pd.to_numeric(df_cat[col_dinero], errors='coerce').fillna(0).sum()
            totales_por_categoria[cat] = total
        else:
            totales_por_categoria[cat] = 0.0

    lista_activos = [cat.strip().upper() for cat in [
        "CONTRUCCIONES PREDIO CALPI OFICINAS", "CUENTAS POR COBRAR GT", "CUENTAS POR COBRAR HN", 
        "CUENTAS POR COBRAR NI", "CUENTAS POR COBRAR SV", "DIESEL EN EQUIPOS Y ALMACENAMIENTOS", 
        "DISPONIBILIDAD", "EQUIPOS DE TRANNSPORTE", "EQUIPOS DE TRANSPORTE EN TRAMITE", 
        "EQUIPOS DE TRANSPORTE EN TRANSITO", "MOBILIARIO Y EQUIPO DE OFICINA", 
        "PENDIENTES DE FACTURAR", "PRESTAMOS A TERCEROS", "PROYECTOS CALPI", "TERRENOS PREDIO CALPI", "OTROS TERRENOS Y PROPIEDADES"
    ]]
    
    lista_pasivos = [cat.strip().upper() for cat in [
        "CUENTAS POR PAGAR SV COMBUSTIBLE", "CUENTAS POR PAGAR SV", 
        "PRESTAMOS ROTATIVOS Y DECRECIENTES", "TRANSPORTES AGREGADOS", "GASTOS POR PAIS Y OBLIGACIONES", "GASTOS MENSUALES EL SALVADOR", "GASTOS ANUALES EL SALVADOR"
    ]]

    col_izquierda, col_derecha = st.columns([3, 7])

    with col_izquierda:
        st.header("📋 Resumen Consolidado")
        with st.container(border=True):
            st.subheader("🟢 Activos")
            datos_activos = [{"Cuenta Contable": cat, "Saldo": totales_por_categoria.get(cat, 0.0)} for cat in lista_activos if totales_por_categoria.get(cat, 0.0) != 0]
            if datos_activos:
                st.dataframe(pd.DataFrame(datos_activos), hide_index=True, use_container_width=True, column_config={"Saldo": st.column_config.NumberColumn("Saldo", format="$ %,.2f")})
            total_activos = sum(totales_por_categoria.get(cat, 0.0) for cat in lista_activos)
            st.metric("TOTAL ACTIVOS", f"$ {total_activos:,.2f}")
        
        with st.container(border=True):
            st.subheader("🔴 Pasivos")
            datos_pasivos = [{"Cuenta Contable": cat, "Saldo": totales_por_categoria.get(cat, 0.0)} for cat in lista_pasivos if totales_por_categoria.get(cat, 0.0) != 0]
            if datos_pasivos:
                st.dataframe(pd.DataFrame(datos_pasivos), hide_index=True, use_container_width=True, column_config={"Saldo": st.column_config.NumberColumn("Saldo", format="$ %,.2f")})
            total_pasivos = sum(totales_por_categoria.get(cat, 0.0) for cat in lista_pasivos)
            st.metric("TOTAL PASIVOS", f"$ {total_pasivos:,.2f}")

    with col_derecha:
        st.header("📑 Detalle por Rubro")
        for cat in categorias_disponibles:
            with st.expander(f"🔹 {cat}"):
                df_cajon = df_master[df_master['Categoria'] == cat].copy()
                cols_finales = [c for c in df_cajon.columns if c not in ['Origen', 'Categoria']]
                
                # Conversión de fechas con día primero
                for col in df_cajon.columns:
                    if 'fecha' in str(col).lower():
                        df_cajon[col] = pd.to_datetime(df_cajon[col], errors='coerce', dayfirst=True)
                
                st.dataframe(df_cajon[cols_finales], hide_index=True, use_container_width=True)
                st.metric(label=f"Total {cat}", value=f"${totales_por_categoria.get(cat, 0.0):,.2f}")