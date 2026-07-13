import streamlit as st
import pandas as pd

# Configuración
st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

# URL directa al CSV de tu Google Sheet
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ozc9yAbVZ3vEhjJEOuQd2D14vhFd7JFSf6D8Jr2R-OQ/export?format=csv"

# Logo desde GitHub (Asegúrate de que la URL sea la 'Raw' del archivo)
LOGO_URL = "https://github.com/ReynaldoCalpi/Reporte-Semanal-Calpi/blob/main/logo.jpg"

# Sidebar
st.sidebar.image(LOGO_URL, use_container_width=True)
st.sidebar.header("Panel de Control")

# --- LÓGICA DE CARGA ---
@st.cache_data(ttl=600) # Se actualiza automáticamente cada 10 min
def cargar_datos():
    try:
        # Leemos el CSV directamente desde el enlace
        df = pd.read_csv(GSHEET_URL)
        # Limpieza básica
        df = df.dropna(how='all') 
        return df
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return pd.DataFrame()

# Botón para refrescar manualmente
if st.sidebar.button("🔄 Refrescar Datos"):
    st.cache_data.clear()
    st.rerun()

df_master = cargar_datos()

# --- AQUÍ VA TU LÓGICA DE VISUALIZACIÓN ---
# A partir de aquí, el resto de tu código que ya tenías para 
# Activos, Pasivos y Desglose, simplemente trabajará con el 'df_master'
# que acabamos de cargar.

if df_master.empty:
    st.warning("⚠️ No se encontraron datos o el enlace no es público.")
else:
    st.title("📊 Reporte Semanal - Transportes Calpi")
    
    # ... (Todo tu bloque de col_izquierda, col_derecha, etc.) ...