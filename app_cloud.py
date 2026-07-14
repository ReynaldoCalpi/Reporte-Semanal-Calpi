import streamlit as st
import pandas as pd
import os

# 1. Configuración principal de la página web
st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

# --- INICIALIZAR ESTADO DE NAVEGACIÓN ---
if 'cat_seleccionada' not in st.session_state:
    st.session_state.cat_seleccionada = None

# Rutas de trabajo fijas
CARPETA = r'https://docs.google.com/spreadsheets/d/1ozc9yAbVZ3vEhjJEOuQd2D14vhFd7JFSf6D8Jr2R-OQ/edit?gid=0#gid=0'
ARCHIVO_SALIDA = 'Reporte_Consolidado_Final.xlsx'
RUTA_COMPLETA_EXCEL = os.path.join(CARPETA, ARCHIVO_SALIDA)

# ==========================================
# MENÚ LATERAL (CONTROL Y LOGO)
# ==========================================
st.sidebar.image("https://raw.githubusercontent.com/ReynaldoCalpi/Reporte-Semanal-Calpi/main/logo.jpg", use_container_width=True)
st.sidebar.header("Panel de Sugerencias")

st.sidebar.markdown("---")
st.sidebar.info("💡Por Favor utilizar este espacio para anotar obervaciones, sugerencias y mejoras y poder evacuarlas en proximas entregas.")
if "notas_calpi" not in st.session_state:
    st.session_state.notas_calpi = []

# Campo para escribir
nota_input = st.sidebar.text_area("Nueva observación:", height=100)

if st.sidebar.button("Guardar Nota"):
    if nota_input:
        st.session_state.notas_calpi.append(nota_input)
        st.sidebar.success("Nota guardada.")
    else:
        st.sidebar.warning("Escribe algo primero.")

# Mostrar notas acumuladas
if st.session_state.notas_calpi:
    st.sidebar.write("**Notas pendientes:**")
    for idx, n in enumerate(st.session_state.notas_calpi):
        st.sidebar.info(f"{idx+1}. {n}")
# ==========================================
# CARGA DE DATOS Y RENDERIZADO DEL DASHBOARD
# ==========================================
from streamlit_gsheets import GSheetsConnection

@st.cache_data(ttl=600)
def cargar_datos():
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = pd.read_csv("https://docs.google.com/spreadsheets/d/e/2PACX-1vQeRzx7jkJ7S1F-5SzuKG35U8llKKTZ3QxlMyR5rzlN96vANkHWHF4wMcH4eYFt673J9LUnBEoUdXNG/pub?output=csv")
    return df

df_master = cargar_datos()
if df_master.empty:
    st.warning("⚠️Anotar obervaciones, sugerencias y mejoras.")
else:
    st.title("📊 Reporte Semanal - Transportes Calpi")
    st.markdown("---")

    if 'Origen' in df_master.columns:
        df_master['Categoria'] = df_master['Origen'].str.replace('.xlsx', '', regex=False).str.strip().str.upper()
    else:
        df_master['Categoria'] = 'GENERAL'

    categorias_disponibles = df_master['Categoria'].dropna().unique()

    # Pre-calcular totales limpios por cajón
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

    # Cuentas Contables Clave
    lista_activos = [cat.strip().upper() for cat in [
        "DISPONIBILIDAD",
        "PRESTAMOS A TERCEROS",
        "TERRENOS PREDIO CALPI",
        "CUENTAS POR COBRAR GT",
        "CONTRUCCIONES PREDIO CALPI OFICINAS",
        "CUENTAS POR COBRAR HN",
        "CUENTAS POR COBRAR NI",
        "CUENTAS POR COBRAR SV",
        "DIESEL EN EQUIPOS Y ALMACENAMIENTOS",
        "EQUIPOS DE TRANNSPORTE",
        "EQUIPOS DE TRANSPORTE EN TRAMITE",
        "EQUIPOS DE TRANSPORTE EN TRANSITO",
        "MOBILIARIO Y EQUIPO DE OFICINA",
        "OTROS TERRENOS Y PROPIEDADES",
        "PENDIENTES DE FACTURAR",
        "PROYECTOS CALPI"
    ]]
    
    lista_pasivos = [cat.strip().upper() for cat in [
        "PRESTAMOS ROTATIVOS Y DECRECIENTES",
        "CUENTAS POR PAGAR SV COMBUSTIBLE",
        "CUENTAS POR PAGAR SV",
        "GASTOS MENSUALES EL SALVADOR",
        "GASTOS ANUALES EL SALVADOR",
        "GASTOS POR PAIS Y OBLIGACIONES",
        "TRANSPORTES AGREGADOS"
    ]]

    # Diseño de pantalla dividida
    col_izquierda, col_derecha = st.columns([3, 7])

    # ==========================================
    # EXTREMO IZQUIERDO: SECCIÓN FINANCIERA
    # ==========================================
    # ==========================================
    # EXTREMO IZQUIERDO: MENÚ DE NAVEGACIÓN
    # ==========================================
    with col_izquierda:
        st.header("📋 Resumen Consolidado")
        
        # --- ACTIVOS ---
        with st.container(border=True):
            st.subheader("🟢 Activos")
            for cat in lista_activos:
                saldo = totales_por_categoria.get(cat, 0.0)
                if saldo != 0:
                    if st.button(f"{cat}\n${saldo:,.2f}", use_container_width=True, key=f"btn_{cat}"):
                        st.session_state.cat_seleccionada = cat
            
            total_activos = sum(totales_por_categoria.get(cat, 0.0) for cat in lista_activos)
            st.metric("TOTAL ACTIVOS", f"$ {total_activos:,.2f}")
        
        # --- PASIVOS ---
        with st.container(border=True):
            st.subheader("🔴 Pasivos")
            for cat in lista_pasivos:
                saldo = totales_por_categoria.get(cat, 0.0)
                if saldo != 0:
                    if st.button(f"{cat}\n${saldo:,.2f}", use_container_width=True, key=f"btn_{cat}"):
                        st.session_state.cat_seleccionada = cat
                
            total_pasivos = sum(totales_por_categoria.get(cat, 0.0) for cat in lista_pasivos)
            st.metric("TOTAL PASIVOS", f"$ {total_pasivos:,.2f}")        
      
        with st.container(border=True):
            st.subheader("🔵 Patrimonio Consolidado")
            patrimonio = total_activos - total_pasivos
            st.metric("DIFERENCIA ENTRE ACTIVOS Y PASIVOS", f"$ {patrimonio:,.2f}")

    # ==========================================
    # LADO DERECHO: DETALLE DINÁMICO (VERSIÓN SEGURA)
    # ==========================================
    with col_derecha:
        st.header("📑 Detalle por Rubro")
        
        if st.session_state.cat_seleccionada is None:
            st.info("👈 Selecciona una cuenta en la izquierda para ver el detalle.")
        else:
            cat = st.session_state.cat_seleccionada
            st.subheader(f"🔹 Detalle: {cat}")
            
            # 1. Filtramos datos
            df_cajon = df_master[df_master['Categoria'] == cat].copy()
            
            # 2. Definir formatos (siempre)
            cols_finales = [c for c in df_cajon.columns if c not in ['Origen', 'Categoria']]
            formatos_columnas = {}
            for c in cols_finales:
                c_low = str(c).lower()
                # Detectar columnas de dinero para formatear
                if "$$" in c_low or any(k in c_low for k in ['contratado', 'disponible', 'monto', 'saldo', 'valor']):
                    df_cajon[c] = pd.to_numeric(df_cajon[c], errors='coerce').fillna(0)
                    formatos_columnas[c] = st.column_config.NumberColumn(format="$ %,.2f")
            
            # 3. Lógica de columnas (Reglas)
            # Aquí es donde llenamos las columnas que quieres ver
            cat_str = str(cat).upper().strip()
            columnas_resumen_vista = []
            
            # [PEGA TUS REGLAS IF/ELIF AQUÍ COMO LAS TENÍAS]
            # Ejemplo rápido de seguridad:
            if not columnas_resumen_vista:
                columnas_resumen_vista = cols_finales # Si las reglas fallan, muestra TODO

            # 4. Renderizado (Garantizado)
            # Convertimos a string los valores para evitar errores visuales con Nones
            df_visual = df_cajon[columnas_resumen_vista].astype(str).replace('nan', '')
            
            # Intentamos dibujar
            st.dataframe(
                df_cajon[columnas_resumen_vista], 
                hide_index=True, 
                use_container_width=True,
                column_config=formatos_columnas
            )
            
            st.metric(label=f"Total {cat}", value=f"$ {totales_por_categoria.get(cat, 0.0):,.2f}")
            
            if st.button("❌ Cerrar vista actual"):
                st.session_state.cat_seleccionada = None
                st.rerun()
                
                # --- LIMPIEZA PROFUNDA DE "NONE" Y NULOS ---
                df_cajon_limpio = df_cajon.copy()
                for col in df_cajon_limpio.columns:
                    if df_cajon_limpio[col].dtype == object:
                        df_cajon_limpio[col] = df_cajon_limpio[col].replace(['None', 'none', 'NaN', 'nan', '', ' '], pd.NA)
                
                # Borra columnas que quedaron 100% nulas/vacías
                df_cajon_limpio = df_cajon_limpio.dropna(axis=1, how='all')
                cols_existentes = list(df_cajon_limpio.columns)
                
                # Detectar la columna del dinero ($$)
                col_dinero = next((col for col in cols_existentes if "$$" in str(col)), None)
                
                if col_dinero:
                    df_cajon_limpio = df_cajon_limpio.dropna(subset=[col_dinero])

                # Buscar columnas comunes de forma flexible para mapeos
                col_fecha = next((c for c in cols_existentes if "fecha" in str(c).lower()), None)
                col_cliente = next((c for c in cols_existentes if "cliente" in str(c).lower()), None)
                col_documento = next((c for c in cols_existentes if "documento" in str(c).lower()), None)
                col_tipo = next((c for c in cols_existentes if "tipo" in str(c).lower()), None)
                col_marca = next((c for c in cols_existentes if "marca" in str(c).lower()), None)
                col_vin = next((c for c in cols_existentes if "vin" in str(c).lower()), None)
                col_poliza = next((c for c in cols_existentes if "póliza" in str(c).lower() or "poliza" in str(c).lower()), None)
                col_banco = next((c for c in cols_existentes if "banco" in str(c).lower()), None)
                col_nombre = next((c for c in cols_existentes if "nombre" in str(c).lower()), None)

                cat_str = str(cat).upper().strip()
                columnas_resumen_vista = []

                # ==================================================================
                # ORDENACIÓN EN LA VISTA PRINCIPAL SEGÚN REGLAS DE NEGOCIO
                # ==================================================================
                if "CONTRUCCIONES" in cat_str or "CONSTRUCCIONES" in cat_str:
                    col_area = next((c for c in cols_existentes if "area" in str(c).lower() or "área" in str(c).lower()), None)
                    if col_area: columnas_resumen_vista.append(col_area)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "CUENTAS POR COBRAR" in cat_str:
                    if col_fecha: columnas_resumen_vista.append(col_fecha)
                    if col_cliente: columnas_resumen_vista.append(col_cliente)
                    if col_documento: columnas_resumen_vista.append(col_documento)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "CUENTAS POR PAGAR" in cat_str:
                    col_proveedor_exacto = next((c for c in cols_existentes if str(c).lower().strip() == "proveedor"), None)
                    if col_fecha: columnas_resumen_vista.append(col_fecha)
                    if col_proveedor_exacto: columnas_resumen_vista.append(col_proveedor_exacto)
                    if col_documento: columnas_resumen_vista.append(col_documento)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "DIESEL" in cat_str:
                    if col_tipo: columnas_resumen_vista.append(col_tipo)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "DISPONIBILIDAD" in cat_str:
                    col_cuenta = next((c for c in cols_existentes if "cuenta" in str(c).lower() or "número" in str(c).lower() or "numero" in str(c).lower()), None)
                    if col_banco: columnas_resumen_vista.append(col_banco)
                    if col_cuenta: columnas_resumen_vista.append(col_cuenta)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "TRANSITO" in cat_str:
                    if col_fecha: columnas_resumen_vista.append(col_fecha)
                    if col_tipo: columnas_resumen_vista.append(col_tipo)
                    if col_marca: columnas_resumen_vista.append(col_marca)
                    if col_vin: columnas_resumen_vista.append(col_vin)
                    if col_poliza: columnas_resumen_vista.append(col_poliza)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "TRANSPORTE" in cat_str or "TRANNSPORTE" in cat_str or "TRAMITE" in cat_str:
                    col_placa = next((c for c in cols_existentes if "placa" in str(c).lower()), None)
                    if col_tipo: columnas_resumen_vista.append(col_tipo)
                    if col_marca: columnas_resumen_vista.append(col_marca)
                    if col_placa: columnas_resumen_vista.append(col_placa)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "MOBILIARIO" in cat_str or "OFICINA" in cat_str:
                    col_codigo = next((c for c in cols_existentes if "codigo" in str(c).lower() or "código" in str(c).lower()), None)
                    col_desc = next((c for c in cols_existentes if "descripcion" in str(c).lower() or "descripción" in str(c).lower() or "desc" in str(c).lower()), None)
                    col_depto = next((c for c in cols_existentes if "departamento" in str(c).lower() or "depto" in str(c).lower()), None)
                    if col_codigo: columnas_resumen_vista.append(col_codigo)
                    if col_desc: columnas_resumen_vista.append(col_desc)
                    if col_depto: columnas_resumen_vista.append(col_depto)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "PENDIENTES" in cat_str:
                    if col_fecha: columnas_resumen_vista.append(col_fecha)
                    if col_cliente: columnas_resumen_vista.append(col_cliente)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "TERCEROS" in cat_str:
                    col_garantia = next((c for c in cols_existentes if "garantia" in str(c).lower() or "garantía" in str(c).lower()), None)
                    if col_nombre: columnas_resumen_vista.append(col_nombre)
                    if col_garantia: columnas_resumen_vista.append(col_garantia)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "PROYECTOS" in cat_str:
                    if col_nombre: columnas_resumen_vista.append(col_nombre)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "ROTATIVOS" in cat_str or "DECRECIENTES" in cat_str:
                    col_contratado = next((c for c in cols_existentes if "contratado" in str(c).lower()), None)
                    col_disponible = next((c for c in cols_existentes if "disponible" in str(c).lower()), None)
                    col_nominal = next((c for c in cols_existentes if "nominal" in str(c).lower()), None)
                    col_efectiva = next((c for c in cols_existentes if "efectiva" in str(c).lower()), None)
                    col_plazo = next((c for c in cols_existentes if "plazo" in str(c).lower()), None)
                    if col_banco: columnas_resumen_vista.append(col_banco)
                    if col_contratado: columnas_resumen_vista.append(col_contratado)
                    if col_disponible: columnas_resumen_vista.append(col_disponible)
                    if col_nominal: columnas_resumen_vista.append(col_nominal)
                    if col_efectiva: columnas_resumen_vista.append(col_efectiva)
                    if col_plazo: columnas_resumen_vista.append(col_plazo)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                elif "TERRENOS" in cat_str:
                    col_matricula = next((c for c in cols_existentes if "matricula" in str(c).lower() or "matrícula" in str(c).lower() or "cnr" in str(c).lower()), None)
                    col_terreno = next((c for c in cols_existentes if "terreno" in str(c).lower()), None)
                    col_compra = next((c for c in cols_existentes if "compra" in str(c).lower()), None)
                    if col_matricula: columnas_resumen_vista.append(col_matricula)
                    if col_terreno: columnas_resumen_vista.append(col_terreno)
                    if col_compra: columnas_resumen_vista.append(col_compra)
                    if col_dinero: columnas_resumen_vista.append(col_dinero)

                # NUEVA REGLA: Gastos y Agregados muestran todas las columnas que tengan información real
                elif "AGREGADOS" in cat_str or "GASTOS" in cat_str:
                    columnas_resumen_vista = [c for c in cols_existentes if c not in ['Origen', 'Categoria']]
                    # Nos aseguramos de que el dinero siempre quede como última columna
                    if col_dinero in columnas_resumen_vista:
                        columnas_resumen_vista.remove(col_dinero)
                        columnas_resumen_vista.append(col_dinero)

                else:
                    if len(cols_finales) > 0:
                        col_texto_inicial = cols_finales[0]
                        columnas_resumen_vista = [col_texto_inicial, col_dinero] if col_texto_inicial != col_dinero else [col_dinero]

                # --- ESCUDO DE SEGURIDAD ---
                columnas_seguras_vista = [c for c in columnas_resumen_vista if c in df_cajon_limpio.columns]

                if columnas_seguras_vista:
                    st.dataframe(
                        df_cajon_limpio[columnas_seguras_vista], 
                        hide_index=True, 
                        use_container_width=True,
                        column_config=formatos_columnas
                    )
                
                monto_total_cajon = totales_por_categoria.get(cat, 0.0)
                st.metric(label=f"Total acumulado en {cat}", value=f"$ {monto_total_cajon:,.2f}")

             

