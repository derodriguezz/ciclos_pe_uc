import streamlit as st
import pandas as pd
import os
from pathlib import Path
import email
from email.policy import default
import pyperclip
import shutil


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
    - Comprime los archivos individuales en un archivo ZIP para su descarga.

    ### Instrucciones:
    1. **Sube tus archivos:** Arrastra y suelta tus correos en formato `.eml` con los Excel adjuntos; o un archivo comprimido con todos los archivos de Excel a consolidar.
    2. **Número de ciclo:** Ingresa el número que identificará el archivo consolidado.
    3. **Procesa:** Haz clic en **"Procesar archivos"** para ejecutar el flujo.
    4. **Descarga:** 
        - El archivo consolidado.
        - Un archivo ZIP con los archivos individuales.
    5. **Copia la macro:** Utiliza la macro de Excel que se encuentra en el botón desplegable.
    6. **Pega y ejecuta la macro:** Utiliza la macro en el archivo de excel para generar las tablas dinámicas rápidamente. 

    ### Al utilizar esta herramienta recuerda:
    - Asegúrate de que tus correos `.eml` contengan archivos de Excel como adjuntos.
    - El archivo comprimido debe tener únicamente los archivos de Excel remitidos por la Agencia de Empleo.
    - Los resultados estarán disponibles hasta que cierres la aplicación.
    """)

    # Mostrar instrucciones y botón para copiar la macro
with st.expander("Instrucciones para usar la macro en Excel"):
    st.markdown("""
    ### Pasos para usar la macro:
    1. Haz clic en el botón **"Copiar Macro"** para copiar el código al portapapeles.
    2. Abre tu archivo consolidado en Excel.
    3. Presiona `Alt + F11` para abrir el editor de Visual Basic.
    4. En el menú, selecciona **Insertar > Módulo**.
    5. Pega el código en el módulo que se abrió.
    6. Desde el editor de Visual Basic, presiona `F5` o selecciona **Ejecutar > Ejecutar Sub/UserForm** para ejecutar el módulo.
    7. Guarda el archivo como un libro estándar de Excel (**.xlsx**), sin habilitar macros.

    **Nota:** 
    - No es necesario habilitar macros en Excel para este flujo.
    - Al guardar el archivo, elije 'Si' para guardarlo como libro sin macros.
                
    """)

    # Código de la macro
    macro = """
Sub CrearTablasDinamicas()
    Dim wsDatos As Worksheet
    Dim wsTD As Worksheet
    Dim pc As PivotCache
    Dim pt1 As PivotTable
    Dim pt2 As PivotTable
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rangoFuente As Range

    ' Intentar asignar la hoja de datos
    On Error Resume Next
    Set wsDatos = ThisWorkbook.Sheets("Sheet1") ' Cambiar "Sheet1" si tu hoja tiene otro nombre
    On Error GoTo 0
    If wsDatos Is Nothing Then
        MsgBox "No se encontró la hoja llamada 'Sheet1'. Verifica el nombre y vuelve a intentar.", vbCritical
        Exit Sub
    End If

    ' Configurar la hoja de resultados "TD"
    On Error Resume Next
    Set wsTD = ThisWorkbook.Sheets("TD")
    If wsTD Is Nothing Then
        Set wsTD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTD.Name = "TD" ' Crear hoja "TD" si no existe
    Else
        wsTD.Cells.Clear ' Limpiar contenido si ya existe
    End If
    On Error GoTo 0

    ' Verificar rango de datos
    ultimaFila = wsDatos.Cells(wsDatos.Rows.Count, 1).End(xlUp).Row
    ultimaColumna = wsDatos.Cells(1, wsDatos.Columns.Count).End(xlToLeft).Column
    If ultimaFila < 2 Or ultimaColumna < 2 Then
        MsgBox "La hoja de datos está vacía o no tiene un formato válido.", vbCritical
        Exit Sub
    End If
    Set rangoFuente = wsDatos.Range(wsDatos.Cells(1, 1), wsDatos.Cells(ultimaFila, ultimaColumna))

    ' Crear PivotCache
    On Error Resume Next
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rangoFuente.Address(True, True, xlR1C1, True))
    If pc Is Nothing Then
        MsgBox "No se pudo crear el PivotCache. Verifica que los datos sean válidos.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' ====================
    ' Primera tabla dinámica
    ' ====================
    On Error Resume Next
    Set pt1 = wsTD.PivotTables.Add(PivotCache:=pc, TableDestination:=wsTD.Cells(3, 1), TableName:="TablaDinamica1")
    On Error GoTo 0
    If pt1 Is Nothing Then
        MsgBox "Error al crear la primera tabla dinámica. Verifica los datos.", vbCritical
        Exit Sub
    End If

    With pt1
        .PivotFields("TALLER Y/O CURSO ").Orientation = xlRowField
        .PivotFields("TALLER Y/O CURSO ").Position = 1
        .PivotFields("TIPO").Orientation = xlColumnField
        .PivotFields("TIPO").Position = 1
        .AddDataField .PivotFields("NUMERO DE IDENTIFICACIÓN"), "Recuento de IDENTIFICACIÓN", xlCount
        .DataBodyRange.NumberFormat = "#,##0"
    End With

    ' ====================
    ' Segunda tabla dinámica
    ' ====================
    On Error Resume Next
    Set pt2 = wsTD.PivotTables.Add(PivotCache:=pc, TableDestination:=wsTD.Cells(pt1.TableRange2.Row + pt1.TableRange2.Rows.Count + 6, 1), TableName:="TablaDinamica2")
    On Error GoTo 0
    If pt2 Is Nothing Then
        MsgBox "Error al crear la segunda tabla dinámica. Verifica los datos.", vbCritical
        Exit Sub
    End If

    With pt2
        .PivotFields("TALLER Y/O CURSO ").Orientation = xlRowField
        .PivotFields("TALLER Y/O CURSO ").Position = 1
        .AddDataField .PivotFields("FECHA DE APERTURA"), "Mín. de FECHA DE APERTURA", xlMin
        .AddDataField .PivotFields("FECHA DE CIERRE "), "Máx. de FECHA DE CIERRE", xlMax
        .PivotFields("TALLER Y/O CURSO ").Subtotals(1) = False
        .DataBodyRange.NumberFormat = "dd/mm/yyyy"
    End With

    MsgBox "Tablas dinámicas creadas exitosamente.", vbInformation
End Sub



    """

# Expander independiente para el código
with st.expander("Ver código de la macro"):
    st.code(macro, language='vb')



# Subir archivos
st.header("Subir correos .eml o archivos comprimidos")
uploaded_files = st.file_uploader("Sube tus correos .eml o un archivo comprimido:", accept_multiple_files=True)

# Carpeta temporal para procesar datos
temp_dir = Path("temp")
temp_dir.mkdir(exist_ok=True)

# Entrada de usuario para el número de ciclo
numero_ciclo = st.text_input("Número de ciclo para guardar el archivo consolidado:", value="---")

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
        # Limpiar carpeta temporal
    if temp_dir.exists():
        shutil.rmtree(temp_dir)
    temp_dir.mkdir(exist_ok=True)
    
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
