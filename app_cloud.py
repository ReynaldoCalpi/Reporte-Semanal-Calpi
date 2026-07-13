import streamlit as st
import pandas as pd
import os

# Configuración inicial
st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

# Rutas: Usamos el directorio actual para que funcione en la nube y local
CARPETA = os.getcwd() 
ARCHIVO_SALIDA = 'Reporte_Consolidado_Final.xlsx'

# --- INICIALIZAR ESTADO ---
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

# --- CARGA DE DATOS ---
df_master = cargar_datos()
if 'Origen' in df_master.columns:
    df_master['Categoria'] = df_master['Origen'].str.replace('.xlsx', '', regex=False).str.strip().str.upper()
else:
    df_master['Categoria'] = 'GENERAL'

# --- INTERFAZ ---
# Sidebar para logo y botón
if os.path.exists(os.path.join(CARPETA, "logo.jpg")):
    st.sidebar.image("logo.jpg", use_container_width=True)

if st.sidebar.button("🔄 Consolidar Nuevos Reportes", use_container_width=True):
    if consolidar_archivos():
        st.sidebar.success("¡Éxito!")
        st.cache_data.clear()
        st.rerun()

st.title("📊 Dashboard Financiero - Calpi")

# Layout de dos columnas
col_izq, col_der = st.columns([1, 3])

# Listas de categorías
lista_categorias = df_master['Categoria'].unique()

with col_izq:
    st.subheader("📋 Resumen Consolidado")
    st.write("Selecciona una cuenta:")
    
    for cat in sorted(lista_categorias):
        if st.button(f"🔹 {cat}", key=cat, use_container_width=True):
            st.session_state.selected_cat = cat

with col_der:
    if st.session_state.selected_cat:
        cat = st.session_state.selected_cat
        st.subheader(f"Detalle: {cat}")
        
        # Filtrar datos
        df_detalle = df_master[df_master['Categoria'] == cat].copy()
        
        # 1. LIMPIEZA: Eliminar columnas totalmente vacías
        df_detalle = df_detalle.dropna(axis=1, how='all')
        
        # 2. FECHAS: Convertir correctamente
        for col in df_detalle.columns:
            if 'fecha' in str(col).lower():
                df_detalle[col] = pd.to_datetime(df_detalle[col], errors='coerce', dayfirst=True)
        
        # 3. Formateo y visualización
        st.dataframe(df_detalle, use_container_width=True)
        
        # Mostrar total si existe columna de dinero
        col_dinero = next((c for c in df_detalle.columns if "$$" in str(c)), None)
        if col_dinero:
            total = pd.to_numeric(df_detalle[col_dinero], errors='coerce').sum()
            st.metric(f"Total {cat}", f"${total:,.2f}")
    else:
        st.info("👈 Haz clic en una categoría de la izquierda para ver el detalle aquí.")