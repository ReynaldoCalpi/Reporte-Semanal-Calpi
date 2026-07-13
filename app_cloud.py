import streamlit as st
import pandas as pd
import os

# 1. Configuración principal de la página web
st.set_page_config(page_title="Dashboard Financiero Calpi", layout="wide")

# Rutas de trabajo fijas
CARPETA = r'https://docs.google.com/spreadsheets/d/1ozc9yAbVZ3vEhjJEOuQd2D14vhFd7JFSf6D8Jr2R-OQ/edit?gid=0#gid=0'
ARCHIVO_SALIDA = 'Reporte_Consolidado_Final.xlsx'
RUTA_COMPLETA_EXCEL = os.path.join(CARPETA, ARCHIVO_SALIDA)

# ==========================================
# MENÚ LATERAL (CONTROL Y LOGO)
# ==========================================
rutas_logo = [
    os.path.join(CARPETA, "logo.png"),
    os.path.join(CARPETA, "logo.jpg"),
    os.path.join(CARPETA, "logo.jpeg"),
    os.path.join(CARPETA, "logo.PNG"),
    os.path.join(CARPETA, "logo.JPG")
]

logo_encontrado = None
for ruta in rutas_logo:
    if os.path.exists(ruta):
        logo_encontrado = ruta
        break

if logo_encontrado:
    st.sidebar.image(logo_encontrado, use_container_width=True)
else:
    st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3303/3303038.png", width=70)
    st.sidebar.warning(f"⚠️ No se encontró el archivo logo.(png/jpg) en la ruta:\n{CARPETA}")

st.sidebar.header("Panel de Control")

def consolidar_archivos():
    dataframes = []
    if not os.path.exists(CARPETA):
        st.sidebar.error("❌ La ruta de la carpeta no existe. Revisa la dirección.")
        return False
        
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
            st.sidebar.error("❌ ¡ERROR DE PERMISOS! Por favor, CIERRA el archivo 'Reporte_Consolidado_Final.xlsx' en Excel e intenta de nuevo.")
            return False
    return False

if st.sidebar.button("🔄 Consolidar Nuevos Reportes", use_container_width=True):
    with st.spinner("Procesando archivos..."):
        if consolidar_archivos():
            st.sidebar.success("¡Consolidación exitosa!")
            st.cache_data.clear()
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.info("💡 Recuerde dac clic en Consolidar Nuevos Reportes arriba si ha hecho algun cambio en ellos.")

# ==========================================
# CARGA DE DATOS Y RENDERIZADO DEL DASHBOARD
# ==========================================
from streamlit_gsheets import GSheetsConnection

@st.cache_data(ttl=600)
def cargar_datos():
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = conn.read(spreadsheet="https://docs.google.com/spreadsheets/d/1ozc9yAbVZ3vEhjJEOuQd2D14vhFd7JFSf6D8Jr2R-OQ/edit", worksheet="DATA")
    return df

df_master = cargar_datos()
if df_master.empty:
    st.warning("⚠️ No se encontraron datos consolidados. Presiona el botón de 'Consolidar Nuevos Reportes' a la izquierda.")
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

    # Diseño de pantalla dividida
    col_izquierda, col_derecha = st.columns([3, 7])

    # ==========================================
    # EXTREMO IZQUIERDO: SECCIÓN FINANCIERA
    # ==========================================
    with col_izquierda:
        st.header("📋 Resumen Consolidado")
        
        with st.container(border=True):
            st.subheader("🟢 Activos")
            datos_activos = [{"Cuenta Contable": cat, "Saldo": totales_por_categoria.get(cat, 0.0)} for cat in lista_activos if totales_por_categoria.get(cat, 0.0) != 0]
            
            if datos_activos:
                df_activos = pd.DataFrame(datos_activos)
                st.dataframe(
                    df_activos, 
                    hide_index=True, 
                    use_container_width=True,
                    column_config={
                        "Cuenta Contable": st.column_config.TextColumn("Cuenta Contable"),
                        "Saldo": st.column_config.NumberColumn("Saldo", format="$ %,.2f")
                    }
                )
            
            total_activos = sum(totales_por_categoria.get(cat, 0.0) for cat in lista_activos)
            st.metric("DISPONIBILIDAD BANCARIA Y ACTIVOS REALIZABLES", f"$ {total_activos:,.2f}")
        
        with st.container(border=True):
            st.subheader("🔴 Pasivos")
            datos_pasivos = [{"Cuenta Contable": cat, "Saldo": totales_por_categoria.get(cat, 0.0)} for cat in lista_pasivos if totales_por_categoria.get(cat, 0.0) != 0]
            
            if datos_pasivos:
                df_pasivos = pd.DataFrame(datos_pasivos)
                st.dataframe(
                    df_pasivos, 
                    hide_index=True, 
                    use_container_width=True,
                    column_config={
                        "Cuenta Contable": st.column_config.TextColumn("Cuenta Contable"),
                        "Saldo": st.column_config.NumberColumn("Saldo", format="$ %,.2f")
                    }
                )
                
            total_pasivos = sum(totales_por_categoria.get(cat, 0.0) for cat in lista_pasivos)
            st.metric("PASIVOS Y DEUDAS A CORTO Y LARGO PLAZO", f"$ {total_pasivos:,.2f}")
        
        with st.container(border=True):
            st.subheader("🔵 Patrimonio Consolidado")
            patrimonio = total_activos - total_pasivos
            st.metric("DIFERENCIA ENTRE ACTIVOS Y PASIVOS", f"$ {patrimonio:,.2f}")

    # ==========================================
    # LADO DERECHO: DETALLES DE CAJONES LIMPIOS
    # ==========================================
    with col_derecha:
        st.header("📑 Detalle por Rubro")
        
        for cat in categorias_disponibles:
            with st.container(border=True):
                st.subheader(f"🔹 {cat}")
                
                df_cajon = df_master[df_master['Categoria'] == cat].copy()
                cols_finales = [c for c in df_cajon.columns if c not in ['Origen', 'Categoria']]
                
                # Formateador de monedas ($$) inteligente
                formatos_columnas = {}
                for c in cols_finales:
                    c_low = str(c).lower()
                    if "$$" in c_low or any(k in c_low for k in ['contratado', 'disponible', 'monto', 'saldo', 'valor']):
                        df_cajon[c] = pd.to_numeric(df_cajon[c], errors='coerce').fillna(0)
                        formatos_columnas[c] = st.column_config.NumberColumn(format="$ %,.2f")
                
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

                # --- RENDERIZAR DETALLES EXTRA COMPLETOS ---
                with st.expander(f"🔍 Ver detalles completos de {cat}"):
                    cols_completas_visualizar = [c for c in cols_existentes if c not in ['Origen', 'Categoria']]
                    if col_dinero in cols_completas_visualizar:
                        cols_completas_visualizar.remove(col_dinero)
                        cols_completas_visualizar.append(col_dinero)

                    st.dataframe(
                        df_cajon_limpio[cols_completas_visualizar], 
                        hide_index=True, 
                        use_container_width=True,
                        column_config=formatos_columnas
                    )
