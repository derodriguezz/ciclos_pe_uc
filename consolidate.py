import streamlit as st
import pandas as pd
import os
from pathlib import Path
import email
from email.policy import default
import xlwings as xw

# Función para guardar adjuntos de correos
def guardar_adjuntos_eml(uploaded_files, carpeta_salida):
    Path(carpeta_salida).mkdir(parents=True, exist_ok=True)
    for uploaded_file in uploaded_files:
        mensaje = email.message_from_bytes(uploaded_file.read(), policy=default)
        for parte in mensaje.walk():
            if parte.get_content_disposition() == "attachment":
                filename = parte.get_filename()
                if filename and (filename.endswith(".xlsx") or filename.endswith(".xls")):
                    ruta_salida = os.path.join(carpeta_salida, filename)
                    with open(ruta_salida, "wb") as archivo_salida:
                        archivo_salida.write(parte.get_payload(decode=True))

# Función para consolidar archivos
def consolidar_archivos(directorio):
    dataframes_validos = []
    for archivo in os.listdir(directorio):
        if archivo.endswith(".xlsx") or archivo.endswith(".xls"):
            ruta_archivo = os.path.join(directorio, archivo)
            xls = pd.ExcelFile(ruta_archivo)
            for hoja in xls.sheet_names:
                df = pd.read_excel(xls, hoja)
                if "CICLO INSCRIPCIÓN" in df.columns:
                    df.rename(columns={"CICLO INSCRIPCIÓN": "CICLO"}, inplace=True)
                if df.shape[1] == 25:
                    dataframes_validos.append(df)
    df_unificado = pd.concat(dataframes_validos, ignore_index=True)
    df_unificado['TIPO'] = df_unificado['NOMBRE DE EMPRESA  PROPULSOR  CONEMPLEO'].apply(
        lambda x: 'Empresa' if pd.notna(x) and x.strip() != '' else 'Cesante'
    )
    # Reemplazos en la columna "TALLER Y/O CURSO"
    reemplazos = {
        "DESARROLLO DE INNOVACIONES INTERNAS  INTRAEMPRENDIMIENTO": "DESARROLLO DE INNOVACIONES INTERNAS – INTRAEMPRENDIMIENTO",
        "ECOMMERCE Y NEGOCIOS DIGITALES": "E-COMMERCE Y NEGOCIOS DIGITALES",
        "EXCEL BASICO INTERMEDIO": "EXCEL BASICO - INTERMEDIO",
        "EXCEL BASICO INTERMEDIO ": "EXCEL BASICO - INTERMEDIO",
        "EXCEL INTERMEDIO AVANZADO": "EXCEL INTERMEDIO - AVANZADO",
        "CONSTRUCCION DE INDICADORES ":"CONSTRUCCION DE INDICADORES",
        "LENGUA DE SEÑAS PARA EL SERVICIO ": "LENGUA DE SEÑAS PARA EL SERVICIO",
        "LENGUAJE DE VENTAS ": "LENGUAJE DE VENTAS",
        "POWER BI PARA GESTION ADMINISTRATIVA": "POWER BI PARA LA GESTION ADMINISTRATIVA",
        "POWER BI PARA LA GESTION ADMINISTRATIVA ": "POWER BI PARA LA GESTION ADMINISTRATIVA",
        "REDACCION Y ORTOGRAFIA TECNICAS Y HABILIADADES PARA LA COMUNICACION ESCRITA": "REDACCION Y ORTOGRAFIA TECNICAS Y HABILIDADES PARA COMUNICACION ESCRITA",
        "REDACCION Y ORTOGRAFIA, TECNICAS Y HABILIDADES PARA COMUNICACION ESCRITA": "REDACCION Y ORTOGRAFIA TECNICAS Y HABILIDADES PARA COMUNICACION ESCRITA",
        "VOCACION DE SERVICIO AL CLIENTE ": "VOCACION DE SERVICIO AL CLIENTE",
        "CONVERSATIONAL ENGLISH FOR BPO'S":"CONVERSATIONAL ENGLISH FOR BPOS",
        " ENGLISH SKILLS":"ENGLISH SKILLS",
        "LOGISTICA Y CADENA DE ABASTECIMIENTO ":"LOGISTICA Y CADENA DE ABASTECIMIENTO",
        "LOGÍSTICA Y CADENA DE ABASTECIMIENTO":"LOGISTICA Y CADENA DE ABASTECIMIENTO",
        "REDACCION Y ORTOGRAFIA TECNICAS Y HABILIDADES PARA LA COMUNICACION ESCRITA":"REDACCION Y ORTOGRAFIA TECNICAS Y HABILIDADES PARA COMUNICACION ESCRITA",
        "GERENCIA DE PROYECTOS CON METODOLOGIA PMI ": "GERENCIA DE PROYECTOS CON METODOLOGIA PMI"
    }
    df_unificado['TALLER Y/O CURSO '] = df_unificado['TALLER Y/O CURSO '].replace(reemplazos)
    columnas_requeridas = ['SEDE PROVEEDOR', 'TIPO DE IDENTIFICACIÓN', 'NUMERO DE IDENTIFICACIÓN', 'GENERO', 
                       'PRIMER NOMBRE', 'SEGUNDO NOMBRE', 'PRIMER APELLIDO', 'SEGUNDO APELLIDO ', 
                       'CELULAR', 'TELEFONO', 'CORREO ELECTRÓNICO']
    df_unificado = df_unificado.dropna(subset=columnas_requeridas, how='all')
    return df_unificado

# Función para insertar tablas dinámicas
def insertar_tablas_dinamicas(ruta_archivo):
    # Abrir el archivo Excel con xlwings
    app = xw.App(visible=False)  # Cambiar a True si deseas observar el proceso
    wb = xw.Book(ruta_archivo)

    try:
        # Verificar si la hoja "TD" existe, si no, crearla
        if "TD" in [sheet.name for sheet in wb.sheets]:
            ws_td = wb.sheets["TD"]
            ws_td.clear()  # Limpiar contenido existente
        else:
            ws_td = wb.sheets.add("TD", after=wb.sheets[-1])  # Crear hoja para tablas dinámicas

        # Obtener la hoja de datos
        ws_datos = wb.sheets[0]  # Cambia el índice si necesitas otra hoja

        # Seleccionar rango de datos
        rango_datos = ws_datos.range("A1").expand("table")  # Expandir rango con datos
        source_data = f"'{ws_datos.name}'!{rango_datos.address}"  # Convertir a rango absoluto

        # Crear PivotCache para la primera tabla dinámica
        pivot_cache1 = wb.api.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=source_data
        )

        # ====================
        # Primera tabla dinámica
        # ====================
        tabla_dinamica1_rango = ws_td.range("A3").address
        tabla_dinamica1 = pivot_cache1.CreatePivotTable(
            TableDestination=ws_td.range("A3").api,
            TableName="TablaDinamica1"
        )

        # Configurar la primera tabla dinámica
        tabla_dinamica1.PivotFields("TALLER Y/O CURSO ").Orientation = 1  # xlRowField
        tabla_dinamica1.PivotFields("TALLER Y/O CURSO ").Position = 1
        tabla_dinamica1.PivotFields("TIPO").Orientation = 2  # xlColumnField
        tabla_dinamica1.PivotFields("TIPO").Position = 1

        # Agregar el campo de datos con formato sin decimales
        data_field = tabla_dinamica1.AddDataField(
            tabla_dinamica1.PivotFields("NUMERO DE IDENTIFICACIÓN"),
            "Recuento de IDENTIFICACIÓN",
            -4112  # xlCount
        )
        data_field.NumberFormat = "#,##0"  # Sin decimales

        # ====================
        # Segunda tabla dinámica
        # ====================
        # Crear PivotCache para la segunda tabla dinámica
        pivot_cache2 = wb.api.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=source_data
        )

        # Ubicación de la segunda tabla dinámica (debajo de la primera)
        tabla_dinamica2_rango = ws_td.range(f"A{ws_td.range('A3').end('down').row + 6}").address
        tabla_dinamica2 = pivot_cache2.CreatePivotTable(
            TableDestination=ws_td.range(tabla_dinamica2_rango).api,
            TableName="TablaDinamica2"
        )

        # Configurar la segunda tabla dinámica
        tabla_dinamica2.PivotFields("TALLER Y/O CURSO ").Orientation = 1  # xlRowField
        tabla_dinamica2.PivotFields("TALLER Y/O CURSO ").Position = 1

        # Configurar FECHA DE APERTURA como un campo de valores con el mínimo
        tabla_dinamica2.AddDataField(
            tabla_dinamica2.PivotFields("FECHA DE APERTURA"),
            "Mín. de FECHA DE APERTURA",
            -4136  # xlMin
        ).NumberFormat = "dd/mm/yyyy"

        # Configurar FECHA DE CIERRE como un campo de valores con el máximo
        tabla_dinamica2.AddDataField(
            tabla_dinamica2.PivotFields("FECHA DE CIERRE "),
            "Máx. de FECHA DE CIERRE",
            -4139  # xlMax
        ).NumberFormat = "dd/mm/yyyy"

        # Desactivar subtotales
        tabla_dinamica2.RowAxisLayout(2)  # xlTabularRow
        tabla_dinamica2.PivotFields("TALLER Y/O CURSO ").Subtotals = [False] * 12

        # Guardar los cambios
        wb.save()
        print("Tablas dinámicas creadas exitosamente.")
    
    except Exception as e:
        print(f"Error al crear tablas dinámicas: {e}")
    
    finally:
        # Cerrar el archivo y la aplicación
        wb.close()
        app.quit()


# Interfaz Streamlit
st.title("ETL de Consolidación de Ciclos")

# Encabezado e instrucciones
st.header("""
          Fundación Universitaria Compensar
          Proyectos Especiales, Coordinación Operativa""")

# Sección desplegable con información
with st.expander("Acerca de esta herramienta"):
    st.markdown("""
    ### ¿Qué hace esta herramienta?
    - Extrae y procesa automáticamente archivos adjuntos de correos en formato `.eml`.
    - Consolida múltiples archivos de Excel en un único archivo organizado.
    - Genera tablas dinámicas en el archivo consolidado.
    - Comprime los archivos individuales en un archivo ZIP para su descarga.

    ### Instrucciones:
    1. **Sube tus archivos:** Arrastra y suelta tus correos en formato `.eml` con los Excel adjuntos; o un archivo comprimido con todos los archivos de Excel a consolidar.
    2. **Número de ciclo:** Ingresa el número que identificará el archivo consolidado.
    3. **Procesa:** Haz clic en **"Procesar archivos"** para ejecutar el flujo.
    4. **Descarga:** 
        - El archivo consolidado con tablas dinámicas.
        - Un archivo ZIP con los archivos individuales.

    ### Al utilizar esta herramienta recuerda:
    - Asegúrate de que tus correos `.eml` contengan archivos de Excel como adjuntos.
    - El archivo comprimido debe tener únicamente los archivos de Excel remitidos por la Agencia de Empleo.
    - Los resultados estarán disponibles hasta que cierres la aplicación.
    """)

# Subir archivos
st.header("Subir correos .eml o archivos comprimidos")
uploaded_files = st.file_uploader("Sube tus correos .eml o un archivo comprimido:", accept_multiple_files=True)

# Carpeta temporal para procesar datos
temp_dir = Path("temp")
temp_dir.mkdir(exist_ok=True)

# Entrada de usuario para el número de ciclo
numero_ciclo = st.text_input("Número de ciclo para guardar el archivo consolidado:", value="---")

# Carpeta temporal para procesar datos
temp_dir = Path("temp")
temp_dir.mkdir(exist_ok=True)

# Función para comprimir archivos individuales en un ZIP
def crear_zip(carpeta, archivo_zip):
    """
    Comprime todos los archivos .xlsx y .xls en la carpeta especificada
    dentro de un archivo ZIP.
    
    Args:
        carpeta (str or Path): Ruta de la carpeta donde están los archivos.
        archivo_zip (str or Path): Ruta donde se guardará el archivo ZIP.
    """
    import zipfile
    with zipfile.ZipFile(archivo_zip, 'w') as zf:
        for archivo in os.listdir(carpeta):
            if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
                zf.write(os.path.join(carpeta, archivo), archivo)



# Procesar datos
if st.button("Procesar archivos"):
    if uploaded_files:
        try:
            # Crear carpeta para guardar los archivos extraídos
            carpeta_bases = temp_dir / "Bases"
            carpeta_bases.mkdir(parents=True, exist_ok=True)
            guardar_adjuntos_eml(uploaded_files, carpeta_bases)
            st.success("Archivos adjuntos extraídos correctamente.")

            # Consolidar archivos en un único DataFrame
            df_consolidado = consolidar_archivos(carpeta_bases)
            archivo_salida = temp_dir / f"Consolidado_Ciclo_{numero_ciclo}.xlsx"
            df_consolidado.to_excel(archivo_salida, index=False)
            st.success("Archivo consolidado generado correctamente.")

            # Insertar tablas dinámicas en el archivo consolidado
            insertar_tablas_dinamicas(archivo_salida)
            st.success("Tablas dinámicas insertadas correctamente.")

            # Crear un archivo ZIP con los Excel individuales
            archivo_zip = temp_dir / f"Archivos_Individuales_Ciclo_{numero_ciclo}.zip"
            crear_zip(carpeta_bases, archivo_zip)
            st.success("Archivo comprimido generado correctamente.")

            # Guardar rutas en `st.session_state`
            st.session_state['archivo_salida'] = archivo_salida
            st.session_state['archivo_zip'] = archivo_zip

        except Exception as e:
            st.error(f"Ocurrió un error al procesar los archivos: {e}")
    else:
        st.warning("Por favor, sube al menos un archivo para procesar.")

# Opciones de descarga
if 'archivo_salida' in st.session_state and 'archivo_zip' in st.session_state:
    with open(st.session_state['archivo_salida'], "rb") as f:
        st.download_button(
            label="Descargar archivo consolidado",
            data=f,
            file_name=f"Consolidado_Ciclo_{numero_ciclo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with open(st.session_state['archivo_zip'], "rb") as f:
        st.download_button(
            label="Descargar archivos comprimidos",
            data=f,
            file_name=f"Archivos_Individuales_Ciclo_{numero_ciclo}.zip",
            mime="application/zip",
        )
