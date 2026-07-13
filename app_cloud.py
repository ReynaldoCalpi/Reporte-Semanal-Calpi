import streamlit as st
import pandas as pd
import os

# Configuración inicial
st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

CARPETA = os.getcwd() 
ARCHIVO_SALIDA = 'Reporte_Consolidado_Final.xlsx'

# Inicializar estado
if 'selected_cat' not in st.session_state:
    st.session_state.selected_cat = None

# --- FUNCIONES ---
def consolidar_archivos():
    dataframes = []
    for archivo in os.listdir(CARPETA):
        if archivo.endswith('.xlsx') and archivo != ARCHIVO_SALIDA:
            try:
                df_temp = pd.read_excel(os.path.join(CARPETA, archivo))
                df_temp['Origen'] = archivo
                dataframes.append(df_temp)
            except: continue
    if dataframes:
        pd.concat(dataframes, ignore_index=True).to_excel(os.path.join(CARPETA, ARCHIVO_SALIDA), index=False)
        return True
    return False

@st.cache_data(ttl=600)
def cargar_datos():
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQeRzx7jkJ7S1F-5SzuKG35U8llKKTZ3QxlMyR5rzlN96vANkHWHF4wMcH4eYFt673J9LUnBEoUdXNG/pub?output=csv"
    return pd.read_csv(url)

# Carga de datos
df_master = cargar_datos()
if 'Origen' in df_master.columns:
    df_master['Categoria'] = df_master['Origen'].str.replace('.xlsx', '', regex=False).str.strip().str.upper()
else:
    df_master['Categoria'] = 'GENERAL'

# Listas de categorías
lista_activos = ["CONTRUCCIONES PREDIO CALPI OFICINAS", "CUENTAS POR COBRAR GT", "CUENTAS POR COBRAR HN", 
                 "CUENTAS POR COBRAR NI", "CUENTAS POR COBRAR SV", "DIESEL EN EQUIPOS Y ALMACENAMIENTOS", 
                 "DISPONIBILIDAD", "EQUIPOS DE TRANSPORTE", "PROYECTOS CALPI", "TERRENOS PREDIO CALPI"]
lista_pasivos = ["CUENTAS POR PAGAR SV COMBUSTIBLE", "CUENTAS POR PAGAR SV", 
                 "PRESTAMOS ROTATIVOS Y DECRECIENTES", "TRANSPORTES AGREGADOS", "GASTOS MENSUALES EL SALVADOR"]

# --- INTERFAZ ---
if os.path.exists(os.path.join(CARPETA, "logo.jpg")):
    st.sidebar.image("logo.jpg", use_container_width=True)

if st.sidebar.button("🔄 Consolidar Nuevos Reportes", use_container_width=True):
    if consolidar_archivos():
        st.cache_data.clear()
        st.rerun()

st.title("📊 Dashboard Financiero - Calpi")

col_izq, col_der = st.columns([1, 2])

# Cálculo de totales globales
totales = {}
for cat in df_master['Categoria'].unique():
    df_cat = df_master[df_master['Categoria'] == cat]
    col_dinero = next((c for c in df_cat.columns if "$$" in str(c)), None)
    totales[cat] = pd.to_numeric(df_cat[col_dinero], errors='coerce').sum() if col_dinero else 0

total_activos = sum(totales.get(cat, 0) for cat in lista_activos)
total_pasivos = sum(totales.get(cat, 0) for cat in lista_pasivos)

with col_izq:
    st.subheader("📋 Resumen Consolidado")
    
    with st.container(border=True):
        st.subheader("🟢 Activos")
        for cat in lista_activos:
            c1, c2 = st.columns([3, 1])
            c1.write(cat)
            if c2.button("Ver", key=f"act_{cat}"): st.session_state.selected_cat = cat
        st.metric("Total Activos", f"${total_activos:,.2f}")

    with st.container(border=True):
        st.subheader("🔴 Pasivos")
        for cat in lista_pasivos:
            c1, c2 = st.columns([3, 1])
            c1.write(cat)
            if c2.button("Ver", key=f"pas_{cat}"): st.session_state.selected_cat = cat
        st.metric("Total Pasivos", f"${total_pasivos:,.2f}")

    with st.container(border=True):
        st.subheader("🔵 Patrimonio")
        st.metric("Activos - Pasivos", f"${(total_activos - total_pasivos):,.2f}")

with col_der:
    if st.session_state.selected_cat:
        cat = st.session_state.selected_cat
        st.subheader(f"Detalle: {cat}")
        
        df_detalle = df_master[df_master['Categoria'] == cat].copy()
        df_detalle = df_detalle.dropna(axis=1, how='all')
        
        # Lógica de columnas: Mover $$ al final
        cols = [c for c in df_detalle.columns if c not in ['Origen', 'Categoria']]
        col_dinero = next((c for c in cols if "$$" in str(c)), None)
        
        columnas_ordenadas = [c for c in cols if "$$" not in str(c)]
        if col_dinero: columnas_ordenadas.append(col_dinero)
        
        # Configuración de columnas
        config = {}
        for c in cols:
            if 'fecha' in str(c).lower():
                df_detalle[c] = pd.to_datetime(df_detalle[c], errors='coerce', dayfirst=True)
                config[c] = st.column_config.DateColumn(c, format="DD/MM/YYYY")
            elif "$$" in str(c):
                config[c] = st.column_config.NumberColumn(c, format="$ %,.2f")
            else:
                # Intentar convertir números normales si no son $$
                df_detalle[c] = pd.to_numeric(df_detalle[c], errors='ignore')

        st.dataframe(df_detalle[columnas_ordenadas], use_container_width=True, column_config=config, hide_index=True)
        if col_dinero:
            st.metric(f"Total {cat}", f"${totales.get(cat, 0):,.2f}")
    else:
        st.info("Selecciona una cuenta a la izquierda para ver el detalle.")