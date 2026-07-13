import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

CARPETA = os.getcwd() 
ARCHIVO_SALIDA = 'Reporte_Consolidado_Final.xlsx'

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

# --- CARGA Y PROCESAMIENTO ---
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

# Calcular totales para el resumen
totales = {}
for cat in df_master['Categoria'].unique():
    df_cat = df_master[df_master['Categoria'] == cat]
    col_dinero = next((c for c in df_cat.columns if "$$" in str(c)), None)
    totales[cat] = pd.to_numeric(df_cat[col_dinero], errors='coerce').fillna(0).sum() if col_dinero else 0

total_activos = sum(totales.get(cat, 0) for cat in lista_activos)
total_pasivos = sum(totales.get(cat, 0) for cat in lista_pasivos)

# --- PANEL IZQUIERDO (Resumen Estático) ---
with col_izq:
    st.subheader("📋 Resumen Consolidado")
    
    # Tabla Activos
    with st.container(border=True):
        st.subheader("🟢 Activos")
        df_resumen_act = pd.DataFrame({"Cuenta Contable": lista_activos, "Saldo": [totales.get(cat, 0) for cat in lista_activos]})
        st.dataframe(df_resumen_act, hide_index=True, use_container_width=True, 
                     column_config={"Saldo": st.column_config.NumberColumn(format="$ %,.2f")})
        st.markdown(f"**DISPONIBILIDAD Y ACTIVOS**")
        st.subheader(f"$ {total_activos:,.2f}")

    # Tabla Pasivos
    with st.container(border=True):
        st.subheader("🔴 Pasivos")
        df_resumen_pas = pd.DataFrame({"Cuenta Contable": lista_pasivos, "Saldo": [totales.get(cat, 0) for cat in lista_pasivos]})
        st.dataframe(df_resumen_pas, hide_index=True, use_container_width=True, 
                     column_config={"Saldo": st.column_config.NumberColumn(format="$ %,.2f")})
        st.markdown(f"**PASIVOS Y DEUDAS**")
        st.subheader(f"$ {total_pasivos:,.2f}")

# --- PANEL DERECHO (Detalle por Rubro) ---
with col_der:
    st.subheader("📑 Detalle por Rubro")
    for cat in sorted(df_master['Categoria'].unique()):
        st.markdown("---")
        st.subheader(f"🔹 {cat}")
        
        df_detalle = df_master[df_master['Categoria'] == cat].copy()
        df_detalle = df_detalle.dropna(axis=1, how='all')
        
        # Mover $$ al final
        cols_dinero = [c for c in df_detalle.columns if "$$" in str(c)]
        cols_resto = [c for c in df_detalle.columns if "$$" not in str(c) and c != 'Categoria' and c != 'Origen']
        df_detalle = df_detalle[cols_resto + cols_dinero]
        
        # Configuración de columnas
        config = {}
        for c in df_detalle.columns:
            if 'fecha' in str(c).lower():
                df_detalle[c] = pd.to_datetime(df_detalle[c], errors='coerce', dayfirst=True)
                config[c] = st.column_config.DateColumn(c, format="DD/MM/YYYY")
            elif "$$" in str(c):
                config[c] = st.column_config.NumberColumn(c, format="$ %,.2f")
            else:
                df_detalle[c] = pd.to_numeric(df_detalle[c], errors='coerce') 

        st.dataframe(df_detalle, use_container_width=True, column_config=config, hide_index=True)
        st.write(f"**Total acumulado en {cat}:**")
        st.markdown(f"### ${totales.get(cat, 0):,.2f}")