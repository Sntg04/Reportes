"""
Módulo para el procesamiento del 4_Reporte_Calidad
Sistema para generación de reportes de calidad con métricas y análisis operativo
"""

import pandas as pd
import io
import os
import glob
from datetime import datetime
from flask import request, jsonify, send_file
from utils.file_utils import allowed_file

# Constantes de configuración
FILTRO_ASIGNACION_MIN = 5
VALOR_INDICADOR_CUMPLE = 0.15
M0_TOQUES_MIN = 120
OTRAS_CARTERAS_TOQUES_MIN = 160
PORCENTAJE_ASIGNACION_M0 = 0.9

def aplicar_formato_tabla(worksheet, dataframe, table_name):
    """
    Aplica formato de tabla Excel a una hoja de cálculo
    
    Args:
        worksheet: Hoja de Excel donde aplicar el formato
        dataframe: DataFrame con los datos
        table_name: Nombre de la tabla
    """
    try:
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.utils import get_column_letter
        
        # Definir el rango de la tabla
        if len(dataframe) == 0:
            # Si no hay datos, crear tabla solo con encabezados
            end_column = get_column_letter(len(dataframe.columns))
            table_range = f"A1:{end_column}1"
        else:
            end_column = get_column_letter(len(dataframe.columns))
            end_row = len(dataframe) + 1  # +1 para incluir encabezados
            table_range = f"A1:{end_column}{end_row}"
        
        # Crear tabla con estilo
        table = Table(displayName=table_name, ref=table_range)
        
        # Aplicar estilo de tabla
        style = TableStyleInfo(
            name="TableStyleMedium9",  # Estilo azul claro
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        
        # Agregar tabla a la hoja
        worksheet.add_table(table)
        
        print(f"OK: Formato de tabla aplicado: {table_name} ({table_range})")
        
    except Exception as e:
        print(f"WARN:  No se pudo aplicar formato de tabla a {table_name}: {str(e)}")

def procesar_reporte_calidad():
    """
    Funcion principal para procesar el reporte de calidad
    """
    try:
        print("INFO: Iniciando procesamiento del Reporte de Calidad...")
        print(f"INFO: Datos del formulario: {list(request.form.keys())}")
        print(f"INFO: Archivos recibidos: {list(request.files.keys())}")
        
        # Verificar si hay archivo automatico del paso 3
        reporte3_auto_file = request.form.get('reporte3_auto_file')
        archivo_reporte3_content = None
        
        if reporte3_auto_file:
            print(f"INFO: Usando archivo automatico del Paso 3: {reporte3_auto_file}")
            # Usar archivo temporal del paso 3
            temp_filepath = os.path.join('temp_files', reporte3_auto_file)
            if os.path.exists(temp_filepath):
                with open(temp_filepath, 'rb') as f:
                    file_content = f.read()
                archivo_reporte3_content = io.BytesIO(file_content)
                print(f"OK: Archivo temporal cargado: {len(file_content)} bytes")
            else:
                print(f"ERROR: Archivo temporal no encontrado: {temp_filepath}")
                return jsonify({
                    'success': False,
                    'message': 'Archivo temporal del Paso 3 no encontrado'
                }), 400
        else:
            print("INFO: Intentando usar archivo manual...")
            # Usar archivo manual
            if 'excelFileReporte3' not in request.files:
                print("ERROR: No se encontro 'excelFileReporte3' en request.files")
                return jsonify({
                    'success': False,
                    'message': 'Se requiere el archivo de Reporte 3'
                }), 400
            
            archivo_reporte3 = request.files['excelFileReporte3']
            if archivo_reporte3.filename == '':
                print("ERROR: Archivo de Reporte 3 sin nombre valido")
                return jsonify({
                    'success': False,
                    'message': 'Archivo de Reporte 3 sin nombre valido'
                }), 400
            
            archivo_reporte3_content = archivo_reporte3
            print(f"OK: Archivo manual recibido: {archivo_reporte3.filename}")
        
        # Generar el reporte de calidad
        resultado = generar_reporte_calidad(archivo_reporte3_content)
        
        return jsonify({
            'success': True,
            'message': 'Reporte de calidad generado exitosamente',
            'filename': resultado['filename'],
            'estadisticas': resultado.get('estadisticas', {}),
            'temp_file': resultado['filename']
        })
        
    except Exception as e:
        print(f"ERROR: Error en procesar_reporte_calidad: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({
            'success': False,
            'message': f'Error interno del servidor: {str(e)}'
        }), 500

def descargar_reporte4():
    """
    Descarga el archivo temporal del Reporte 4
    """
    try:
        temp_filename = request.json.get('temp_file')
        if not temp_filename:
            return jsonify({'success': False, 'message': 'Nombre de archivo temporal no proporcionado'}), 400
        
        temp_filepath = os.path.join('temp_files', temp_filename)
        
        if not os.path.exists(temp_filepath):
            return jsonify({'success': False, 'message': 'Archivo temporal no encontrado'}), 404
        
        return send_file(
            temp_filepath,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=temp_filename
        )
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al descargar archivo: {str(e)}'}), 500

def generar_prueba_reporte4():
    """
    Genera un archivo de prueba del Reporte 4 sin necesidad de archivos de entrada
    """
    try:
        # Generar el reporte de calidad sin archivos de entrada
        resultado = generar_reporte_calidad(None)
        
        return jsonify({
            'success': True,
            'message': 'Archivo de prueba generado exitosamente',
            'filename': resultado['filename'],
            'estadisticas': resultado.get('estadisticas', {}),
            'temp_file': resultado['filename']
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'message': f'Error generando archivo de prueba: {str(e)}'
        }), 500

def generar_reporte_calidad(archivo_reporte3):
    """
    Genera el archivo Excel del reporte de calidad con datos del Reporte 3 integrados
    """
    try:
        # Leer datos del Reporte 3 si existe
        df_reporte3 = None
        if archivo_reporte3:
            try:
                # Leer todas las hojas del archivo
                all_sheets = pd.read_excel(archivo_reporte3, sheet_name=None)
                print(f"Hojas encontradas en Reporte 3: {list(all_sheets.keys())}")
                
                # Combinar todas las hojas en un solo DataFrame
                dataframes_list = []
                total_registros = 0
                
                for sheet_name, df_sheet in all_sheets.items():
                    if not df_sheet.empty:
                        print(f"  - Procesando hoja '{sheet_name}': {len(df_sheet)} registros")
                        print(f"    Columnas: {list(df_sheet.columns)}")
                        
                        # Copiar sin modificaciones para preservar todos los datos
                        df_copy = df_sheet.copy()
                        dataframes_list.append(df_copy)
                        total_registros += len(df_copy)
                        print(f"  - OK: Hoja '{sheet_name}' agregada: {len(df_copy)} registros")
                    else:
                        print(f"WARN: Hoja '{sheet_name}' está vacía, omitida")
                
                print(f"INFO: Hojas procesadas: {len(dataframes_list)}, Registros totales: {total_registros}")
                
                if dataframes_list:
                    # Combinar todos los DataFrames SIN eliminar duplicados
                    df_reporte3 = pd.concat(dataframes_list, ignore_index=True, sort=False)
                    
                    # Verificar que no se perdieron registros
                    if len(df_reporte3) != total_registros:
                        print(f"ERROR: Se perdieron {total_registros - len(df_reporte3)} registros durante pd.concat")
                    else:
                        print(f"OK: Todos los registros preservados correctamente")
                    
                    # Procesar fechas SOLO si existe la columna y DESPUES de verificar integridad
                    if 'Fecha' in df_reporte3.columns:
                        # Procesamiento de fechas
                        print(f"INFO: Procesando fechas - Registros: {len(df_reporte3)}")
                        fechas_originales = df_reporte3['Fecha'].copy()
                        
                        # Convertir la columna Fecha a datetime usando múltiples formatos
                        df_reporte3['Fecha'] = pd.to_datetime(df_reporte3['Fecha'], 
                                                            format='%d/%m/%Y', 
                                                            errors='coerce')
                        
                        # Si falló, intentar con otros formatos comunes
                        mask_nulos = df_reporte3['Fecha'].isna()
                        if mask_nulos.any():
                            print(f"INFO: Intentando formato alternativo para {mask_nulos.sum()} fechas")
                            # Restaurar valores originales para las fechas que fallaron
                            df_reporte3.loc[mask_nulos, 'Fecha'] = fechas_originales.loc[mask_nulos]
                            fechas_temp = pd.to_datetime(df_reporte3.loc[mask_nulos, 'Fecha'], 
                                                       format='%m/%d/%Y', 
                                                       errors='coerce')
                            df_reporte3.loc[mask_nulos, 'Fecha'] = fechas_temp
                            
                        # Si aun fallo, intentar formato YYYY-MM-DD
                        mask_nulos = df_reporte3['Fecha'].isna()
                        if mask_nulos.any():
                            print(f"  - Intentando formato YYYY-MM-DD para {mask_nulos.sum()} fechas...")
                            # Restaurar valores originales para las fechas que aun fallan
                            df_reporte3.loc[mask_nulos, 'Fecha'] = fechas_originales.loc[mask_nulos]
                            fechas_temp = pd.to_datetime(df_reporte3.loc[mask_nulos, 'Fecha'], 
                                                       format='%Y-%m-%d', 
                                                       errors='coerce')
                            df_reporte3.loc[mask_nulos, 'Fecha'] = fechas_temp
                        
                        # Ordenar por fecha SOLO si se pudieron convertir las fechas
                        fechas_validas = df_reporte3['Fecha'].notna().sum()
                        print(f"  - Fechas validas para ordenamiento: {fechas_validas} de {len(df_reporte3)}")
                        
                        if fechas_validas > 0:
                            df_reporte3 = df_reporte3.sort_values('Fecha', ascending=True, na_position='last')
                            print(f"  - Datos ordenados por fecha (registros sin fecha al final)")
                        
                        print(f"  - Registros despues de procesar fechas: {len(df_reporte3)}")
                        
                        # Convertir fecha SIEMPRE a formato DD/MM/YYYY
                        df_reporte3['Fecha'] = df_reporte3['Fecha'].apply(
                            lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ''
                        )
                        
                        print(f"OK: Fechas procesadas y formato DD/MM/YYYY aplicado")
                    else:
                        print(f"WARN: No se encontró columna 'Fecha' para ordenar")
                    
                    # FILTRAR REGISTROS CON ASIGNACION MENOR O IGUAL A 5
                    registros_antes_filtro = len(df_reporte3)
                    
                    # Buscar la columna de asignacion (con o sin acento)
                    columna_asignacion = None
                    for col in df_reporte3.columns:
                        if col.lower().replace('ó', 'o').replace('ñ', 'n') in ['asignacion']:
                            columna_asignacion = col
                            break
                    
                    if columna_asignacion:
                        print(f"INFO: Columna de asignacion encontrada: '{columna_asignacion}'")
                        # Convertir a numerico para poder filtrar
                        df_reporte3[columna_asignacion] = pd.to_numeric(df_reporte3[columna_asignacion], errors='coerce')
                        
                        # Mostrar estadísticas antes del filtro
                        valores_menores_5 = df_reporte3[df_reporte3[columna_asignacion] <= FILTRO_ASIGNACION_MIN]
                        print(f"INFO: Registros con {columna_asignacion} <= {FILTRO_ASIGNACION_MIN}: {len(valores_menores_5)}")
                        if len(valores_menores_5) > 0:
                            print(f"INFO: Valores a eliminar: {valores_menores_5[columna_asignacion].value_counts().sort_index().to_dict()}")
                        
                        # Filtrar registros con Asignación > FILTRO_ASIGNACION_MIN
                        df_reporte3 = df_reporte3[df_reporte3[columna_asignacion] > FILTRO_ASIGNACION_MIN]
                        registros_eliminados = registros_antes_filtro - len(df_reporte3)
                        print(f"INFO: Filtro aplicado - Registros: {registros_antes_filtro} → {len(df_reporte3)} (eliminados: {registros_eliminados})")
                    else:
                        print(f"WARN: No se encontró columna de Asignación para filtrar")
                    
                    print(f"OK: Datos procesados: {len(df_reporte3)} registros de {len(dataframes_list)} hojas")
                else:
                    print("WARN: No se encontraron datos en ninguna hoja del Reporte 3")
                    df_reporte3 = None
                    
            except Exception as e:
                print(f"ERROR: Error leyendo Reporte 3: {str(e)}")
                df_reporte3 = None
        
        # Crear archivo Excel en memoria
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Crear hojas en orden inverso para que Planta quede primera
            # Crear hoja "Consolidado"
            crear_hoja_consolidado(writer)
            
            # Crear hoja "Gerente"
            crear_hoja_gerente(writer)
            
            # Crear hoja "Team"
            crear_hoja_team(writer)
            
            # Crear hoja "Operativo" con datos del Reporte 3
            crear_hoja_operativo(writer, df_reporte3)
            
            # Crear hoja "Calidad"
            crear_hoja_calidad(writer)
            
            # Crear hoja "Ausentismo"
            crear_hoja_ausentismo(writer)
            
            # Crear hoja "Asistencia Lideres"
            crear_hoja_asistencia_lideres(writer)
            
            # Crear hoja "Planta" con las tres tablas
            crear_hoja_planta(writer)
        
        output.seek(0)
        
        # Generar nombre de archivo
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"4_Reporte_Calidad_{timestamp}.xlsx"
        
        # Guardar archivo temporalmente
        os.makedirs('temp_files', exist_ok=True)
        temp_filepath = os.path.join('temp_files', filename)
        with open(temp_filepath, 'wb') as f:
            f.write(output.getvalue())
        
        return {
            'filename': filename,
            'estadisticas': {
                'hojas_creadas': ['Consolidado', 'Gerente', 'Team', 'Operativo', 'Calidad', 'Ausentismo', 'Asistencia Lideres', 'Planta'],
                'tablas_planta': 3,
                'total_hojas': 8
            }
        }
        
    except Exception as e:
        raise Exception(f"Error generando reporte: {str(e)}")

def crear_hoja_planta(writer):
    """
    Crea la hoja "Planta" con las tres tablas especificadas
    """
    # Tabla 1: Usuarios
    datos_usuarios = [
        ["William Cabiativa", "1016019347", "Gerente"],
        ["Daniela Arias", "1015447386", "Gerente"],
        ["Yesid Espitia", "1013666144", "Gerente"],
        ["Luis Alzate", "1037648173", "Gerente"],
        ["Maria Acero", "1020798229", "Back"],
        ["Luis Aleman", "1003187750", "Team"],
        ["Nancy Rodriguez", "1014269628", "Team"],
        ["Danilo Rodriguez", "1016079907", "Team"],
        ["Camilo Arciniegas", "1003458325", "Team"],
        ["Brayan Murcia", "1010237867", "Team"],
        ["Edgar Parra", "1015456724", "Team"],
        ["Nancy Cruz", "1001296759", "Team"],
        ["Nicolas Briceno", "1013097231", "Back"],
        ["Kebin Bernal", "1033691778", "Back"],
        ["Natalia Quiceno", "1019108899", "Back"],
        ["Neverson Ulloa", "1003777394", "Back"],
        ["Zharik Jimenez", "1070590063", "Back"],
        ["Luisa Arevalo", "1031801240", "Back"],
        ["Brayan Sanchez", "1233904529", "Gerente"],
        ["Paula Rubio", "1007155877", "Team"],
        ["Joan Ruiz", "1018447274", "Gerente"],
        ["Beatriz Hao", "0", "Gerente"],
        ["Lizethe Rodriguez", "1105690146", "Team"],
        ["Andres Acevedo", "1000036873", "Back"]
    ]
    
    df_usuarios = pd.DataFrame(datos_usuarios, columns=["Usuario", "Cedula", "Cargo"])
    
    # Tabla 2: Dia Pago
    datos_dia_pago = [
        ["M0-PP", "37,0%"],
        ["M0-VP", "62,0%"],
        ["M1-1A", "9,0%"],
        ["M1-1B", "2,5%"],
        ["M0-PN", "52,0%"],
        ["M0-FRS", "52,0%"],
        ["M0-BT", "52,0%"],
        ["M0-1-PP", "10,0%"],
        ["M1-1A-FRS", "8,0%"],
        ["M1-1A-BT", "8,0%"],
        ["M1-1A-PN", "8,0%"]
    ]
    
    df_dia_pago = pd.DataFrame(datos_dia_pago, columns=["Dia Pago", "Meta"])
    
    # Tabla 3: Dia Normal
    datos_dia_normal = [
        ["M0-PP", "36,0%"],
        ["M0-VP", "58,0%"],
        ["M1-1A", "9,0%"],
        ["M1-1B", "2,5%"],
        ["M0-PN", "52,0%"],
        ["M0-FRS", "52,0%"],
        ["M0-BT", "52,0%"],
        ["M0-1-PP", "10,0%"],
        ["M1-1A-FRS", "8,0%"],
        ["M1-1A-BT", "8,0%"],
        ["M1-1A-PN", "8,0%"]
    ]
    
    df_dia_normal = pd.DataFrame(datos_dia_normal, columns=["Dia Normal", "Meta"])
    
    # Crear la hoja "Planta" directamente con pandas
    # Tabla 1: Usuarios (A1:C25)
    df_usuarios.to_excel(writer, sheet_name="Planta", startrow=0, startcol=0, index=False)
    
    # Obtener la hoja despues de crear la primera tabla
    worksheet = writer.sheets["Planta"]
    
    # Escribir las otras tablas manualmente para control de posicion
    # Tabla 2: Dia Pago (E1:F12)
    worksheet.cell(row=1, column=5, value="Dia Pago")
    worksheet.cell(row=1, column=6, value="Meta")
    for idx, row in enumerate(datos_dia_pago, start=2):
        worksheet.cell(row=idx, column=5, value=row[0])
        worksheet.cell(row=idx, column=6, value=row[1])
    
    # Tabla 3: Dia Normal (H1:I12)
    worksheet.cell(row=1, column=8, value="Dia Normal")
    worksheet.cell(row=1, column=9, value="Meta")
    for idx, row in enumerate(datos_dia_normal, start=2):
        worksheet.cell(row=idx, column=8, value=row[0])
        worksheet.cell(row=idx, column=9, value=row[1])
    
    # Crear tablas de Excel reales
    from openpyxl.worksheet.table import Table, TableStyleInfo
    
    # Tabla 1: Usuarios
    tabla_usuarios = Table(displayName="TablaUsuarios", ref="A1:C25")
    style_usuarios = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                  showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tabla_usuarios.tableStyleInfo = style_usuarios
    worksheet.add_table(tabla_usuarios)
    
    # Tabla 2: Dia Pago  
    tabla_dia_pago = Table(displayName="TablaDiaPago", ref="E1:F12")
    style_dia_pago = TableStyleInfo(name="TableStyleMedium15", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tabla_dia_pago.tableStyleInfo = style_dia_pago
    worksheet.add_table(tabla_dia_pago)
    
    # Tabla 3: Dia Normal
    tabla_dia_normal = Table(displayName="TablaDiaNormal", ref="H1:I12")
    style_dia_normal = TableStyleInfo(name="TableStyleMedium21", showFirstColumn=False,
                                     showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tabla_dia_normal.tableStyleInfo = style_dia_normal
    worksheet.add_table(tabla_dia_normal)
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 20  # Usuario
    worksheet.column_dimensions['B'].width = 12  # Cedula
    worksheet.column_dimensions['C'].width = 10  # Cargo
    worksheet.column_dimensions['E'].width = 15  # Dia Pago
    worksheet.column_dimensions['F'].width = 8   # Meta Pago
    worksheet.column_dimensions['H'].width = 15  # Dia Normal
    worksheet.column_dimensions['I'].width = 8   # Meta Normal

def crear_hoja_asistencia_lideres(writer):
    """
    Crea la hoja "Asistencia Lideres" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_asistencia = [
        "Codigo Aus", "Tipo Jornada", "Fecha", "Usuario", "Cedula", 
        "Cargo", "Ingreso", "Salida", "Horas laboradas", "Novedad Ingreso", "Drive"
    ]
    
    # Crear DataFrame con una fila de ejemplo para evitar errores de tabla
    datos_ejemplo = [["", "", "", "", "", "", "", "", "", "", ""]]
    df_asistencia = pd.DataFrame(datos_ejemplo, columns=columnas_asistencia)
    
    # Escribir a Excel
    df_asistencia.to_excel(writer, sheet_name="Asistencia Lideres", index=False)
    
    # Obtener la hoja y aplicar formato basico
    worksheet = writer.sheets["Asistencia Lideres"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, 12):  # A hasta K
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 12  # Codigo Aus
    worksheet.column_dimensions['B'].width = 15  # Tipo Jornada
    worksheet.column_dimensions['C'].width = 12  # Fecha
    worksheet.column_dimensions['D'].width = 20  # Usuario
    worksheet.column_dimensions['E'].width = 12  # Cedula
    worksheet.column_dimensions['F'].width = 10  # Cargo
    worksheet.column_dimensions['G'].width = 10  # Ingreso
    worksheet.column_dimensions['H'].width = 10  # Salida
    worksheet.column_dimensions['I'].width = 15  # Horas laboradas
    worksheet.column_dimensions['J'].width = 18  # Novedad Ingreso
    worksheet.column_dimensions['K'].width = 10  # Drive
    
    # Aplicar formato de tabla
    aplicar_formato_tabla(worksheet, df_asistencia, "TablaAsistenciaLideres")

def crear_hoja_ausentismo(writer):
    """
    Crea la hoja "Ausentismo" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_ausentismo = [
        "Codigo Aus", "Codigo", "Tipo Jornada", "Fecha", "Cedula", "ID",
        "Nombre", "Sede", "Ubicacion", "Logueo Admin", "Ingreso", 
        "Salida", "Horas laboradas", "Novedad Ingreso", "Validacion", "Drive"
    ]
    
    # Crear DataFrame con una fila de ejemplo para evitar errores de tabla
    datos_ejemplo = [["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]]
    df_ausentismo = pd.DataFrame(datos_ejemplo, columns=columnas_ausentismo)
    
    # Escribir a Excel
    df_ausentismo.to_excel(writer, sheet_name="Ausentismo", index=False)
    
    # Obtener la hoja y aplicar formato basico
    worksheet = writer.sheets["Ausentismo"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, 17):  # A hasta P
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 12  # Codigo Aus
    worksheet.column_dimensions['B'].width = 10  # Codigo
    worksheet.column_dimensions['C'].width = 15  # Tipo Jornada
    worksheet.column_dimensions['D'].width = 12  # Fecha
    worksheet.column_dimensions['E'].width = 12  # Cedula
    worksheet.column_dimensions['F'].width = 8   # ID
    worksheet.column_dimensions['G'].width = 20  # Nombre
    worksheet.column_dimensions['H'].width = 10  # Sede
    worksheet.column_dimensions['I'].width = 12  # Ubicacion
    worksheet.column_dimensions['J'].width = 15  # Logueo Admin
    worksheet.column_dimensions['K'].width = 10  # Ingreso
    worksheet.column_dimensions['L'].width = 10  # Salida
    worksheet.column_dimensions['M'].width = 15  # Horas laboradas
    worksheet.column_dimensions['N'].width = 18  # Novedad Ingreso
    worksheet.column_dimensions['O'].width = 12  # Validacion
    worksheet.column_dimensions['P'].width = 10  # Drive
    
    # Aplicar formato de tabla
    aplicar_formato_tabla(worksheet, df_ausentismo, "TablaAusentismo")

def leer_archivo_monitoreos():
    """
    Lee el archivo de monitoreos desde la carpeta 'Monitoreos'
    
    Returns:
        pd.DataFrame: DataFrame con los datos de monitoreos o DataFrame vacío si no encuentra el archivo
    """
    try:
        # Crear carpeta Monitoreos si no existe (con M mayúscula)
        if not os.path.exists('Monitoreos'):
            os.makedirs('Monitoreos')
            print("INFO: Carpeta 'Monitoreos' creada")
        
        # Buscar archivos en la carpeta Monitoreos
        patrones_archivos = [
            'Monitoreos/*.xlsx',
            'Monitoreos/*.xls', 
            'Monitoreos/*.csv'
        ]
        
        archivo_encontrado = None
        for patron in patrones_archivos:
            archivos = glob.glob(patron)
            if archivos:
                # Tomar el archivo más reciente
                archivo_encontrado = max(archivos, key=os.path.getmtime)
                break
        
        if not archivo_encontrado:
            print("ADVERTENCIA: No se encontró archivo de monitoreos en la carpeta 'Monitoreos'")
            return pd.DataFrame()
        
        print(f"INFO: Cargando archivo de monitoreos: {archivo_encontrado}")
        
        # Leer el archivo según su extensión
        if archivo_encontrado.endswith('.csv'):
            df_monitoreos = pd.read_csv(archivo_encontrado)
        else:
            df_monitoreos = pd.read_excel(archivo_encontrado)
        
        # Limpiar columnas unnamed (columnas vacías)
        df_monitoreos = df_monitoreos.loc[:, ~df_monitoreos.columns.str.contains('^Unnamed')]
        
        # Verificar que tenga las columnas requeridas
        columnas_requeridas = ['Fecha Monitoreo', 'ID Asesor', 'VOZ', 'SMS', 'TERCERO', 'Nota Total', 'Total Monitoreos']
        columnas_disponibles = [col for col in columnas_requeridas if col in df_monitoreos.columns]
        columnas_faltantes = [col for col in columnas_requeridas if col not in df_monitoreos.columns]
        
        if columnas_faltantes:
            print(f"ADVERTENCIA: Columnas faltantes en archivo de monitoreos: {columnas_faltantes}")
            
        print(f"INFO: Columnas encontradas: {columnas_disponibles}")
        print(f"INFO: Archivo de monitoreos cargado exitosamente: {len(df_monitoreos)} registros")
        
        return df_monitoreos
        
    except Exception as e:
        print(f"ERROR: Error al leer archivo de monitoreos: {str(e)}")
        return pd.DataFrame()


def crear_hoja_calidad(writer):
    """
    Crea la hoja "Calidad" con datos del archivo de monitoreos
    """
    # Leer datos de monitoreos
    df_monitoreos = leer_archivo_monitoreos()
    
    # Definir columnas esperadas (agregando Codigo como primera columna)
    columnas_calidad = [
        "Codigo", "Fecha Monitoreo", "ID Asesor", "VOZ", "SMS", 
        "TERCERO", "Nota Total", "Total Monitoreos"
    ]
    
    # Si tenemos datos de monitoreos, usarlos; si no, crear DataFrame vacío con las columnas
    if not df_monitoreos.empty:
        # Filtrar y reorganizar columnas según lo requerido (sin la columna Codigo del archivo original)
        columnas_disponibles = [col for col in columnas_calidad[1:] if col in df_monitoreos.columns]  # Excluir "Codigo" de la búsqueda
        df_calidad = df_monitoreos[columnas_disponibles].copy()
        
        # Agregar columna Codigo como primera columna (vacía inicialmente, se llenará con fórmulas)
        df_calidad.insert(0, 'Codigo', '')
        
        # Procesar fecha para que se muestre correctamente
        if 'Fecha Monitoreo' in df_calidad.columns:
            df_calidad['Fecha Monitoreo'] = pd.to_datetime(df_calidad['Fecha Monitoreo'], errors='coerce').dt.strftime('%d/%m/%Y')
        
        # Asegurar que ID Asesor sea string sin decimales
        if 'ID Asesor' in df_calidad.columns:
            df_calidad['ID Asesor'] = df_calidad['ID Asesor'].fillna(0).astype(int).astype(str)
        
        # Procesar columnas de porcentaje - convertir a decimal para formato de Excel
        columnas_porcentaje = ['VOZ', 'SMS', 'TERCERO', 'Nota Total']
        for col in columnas_porcentaje:
            if col in df_calidad.columns:
                # Convertir a numérico y mantener como decimal (ya están en formato 0-1)
                df_calidad[col] = pd.to_numeric(df_calidad[col], errors='coerce')
        
        # Agregar columnas faltantes como vacías
        for col in columnas_calidad:
            if col not in df_calidad.columns:
                df_calidad[col] = ""
        
        # Reorganizar columnas en el orden correcto
        df_calidad = df_calidad[columnas_calidad]
        
        # FILTRAR FILAS CON VALORES PROBLEMÁTICOS
        # Eliminar filas que tengan exactamente el número 0 o un solo número en columnas importantes
        filas_antes = len(df_calidad)
        
        def es_fila_problematica(row):
            """Verifica si una fila tiene solo 0s o números únicos problemáticos"""
            # Verificar columnas numéricas importantes
            columnas_verificar = ['ID Asesor', 'VOZ', 'SMS', 'TERCERO', 'Nota Total', 'Total Monitoreos']
            
            for col in columnas_verificar:
                if col in row.index:
                    valor = row[col]
                    # Si el valor es string y es exactamente "0"
                    if isinstance(valor, str) and valor.strip() == '0':
                        return True
                    # Si el valor es numérico y es exactamente 0
                    elif pd.notna(valor) and valor == 0:
                        # Solo eliminar si TODAS las columnas importantes son 0 o vacías
                        valores_importantes = [row[c] for c in columnas_verificar if c in row.index]
                        valores_validos = [v for v in valores_importantes if pd.notna(v) and v != 0 and str(v).strip() != '0']
                        if len(valores_validos) == 0:
                            return True
            
            return False
        
        # Aplicar filtro
        mask_mantener = ~df_calidad.apply(es_fila_problematica, axis=1)
        df_calidad = df_calidad[mask_mantener].reset_index(drop=True)
        
        filas_despues = len(df_calidad)
        filas_eliminadas = filas_antes - filas_despues
        
        if filas_eliminadas > 0:
            print(f"INFO: Eliminadas {filas_eliminadas} filas con valores problemáticos (0s o números únicos)")
        
        # GENERAR CÓDIGOS antes de escribir a Excel
        print("INFO: Generando códigos en hoja Calidad...")
        
        def generar_codigo(row):
            try:
                id_asesor = str(row['ID Asesor']).strip()
                fecha_str = str(row['Fecha Monitoreo']).strip()
                
                if id_asesor and fecha_str and id_asesor != '0' and fecha_str != '':
                    # Parsear la fecha en formato DD/MM/YYYY
                    if '/' in fecha_str:
                        parts = fecha_str.split('/')
                        if len(parts) == 3:
                            dia = parts[0].zfill(2)
                            mes = parts[1].zfill(2)
                            return f"{id_asesor}{dia}{mes}"
                return ""
            except:
                return ""
        
        # Aplicar la función para generar códigos
        df_calidad['Codigo'] = df_calidad.apply(generar_codigo, axis=1)
        
        print(f"INFO: Códigos generados para {len(df_calidad[df_calidad['Codigo'] != ''])} registros")
        print(f"INFO: Hoja Calidad creada con {len(df_calidad)} registros de monitoreos")
    else:
        # Crear DataFrame vacío con las columnas requeridas
        df_calidad = pd.DataFrame(columns=columnas_calidad)
        print("INFO: Hoja Calidad creada sin datos (archivo de monitoreos no encontrado)")
    
    # Escribir a Excel
    df_calidad.to_excel(writer, sheet_name="Calidad", index=False)
    
    # Obtener la hoja y aplicar formato básico
    worksheet = writer.sheets["Calidad"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, len(columnas_calidad) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 15  # Codigo
    worksheet.column_dimensions['B'].width = 18  # Fecha Monitoreo
    worksheet.column_dimensions['C'].width = 12  # ID Asesor
    worksheet.column_dimensions['D'].width = 8   # VOZ
    worksheet.column_dimensions['E'].width = 8   # SMS
    worksheet.column_dimensions['F'].width = 10  # TERCERO
    worksheet.column_dimensions['G'].width = 12  # Nota Total
    worksheet.column_dimensions['H'].width = 18  # Total Monitoreos
    
    # Aplicar formato de porcentaje a las columnas especificadas
    porcentaje_format = '0.00%'
    columnas_porcentaje = ['VOZ', 'SMS', 'TERCERO', 'Nota Total']
    
    # Obtener las letras de columna para las columnas de porcentaje
    columnas_porcentaje_indices = {}
    for idx, col_name in enumerate(columnas_calidad, 1):
        if col_name in columnas_porcentaje:
            columnas_porcentaje_indices[col_name] = idx
    
    # Aplicar formato de porcentaje a todas las celdas de esas columnas
    if not df_calidad.empty:
        for col_name, col_idx in columnas_porcentaje_indices.items():
            col_letter = chr(64 + col_idx)  # Convertir número a letra (A=1, B=2, etc.)
            for row_num in range(2, len(df_calidad) + 2):  # Empezar desde fila 2 (después del encabezado)
                cell = worksheet[f"{col_letter}{row_num}"]
                cell.number_format = porcentaje_format
    
    # Aplicar formato de tabla si hay datos
    if not df_calidad.empty:
        aplicar_formato_tabla(worksheet, df_calidad, "TablaCalidad")

def verificar_integridad_datos(df_operativo, mapeo_columnas):
    """
    Verifica la integridad de los datos transferidos del Reporte 3 al Operativo
    
    Args:
        df_operativo (pd.DataFrame): DataFrame del operativo con datos mapeados
        mapeo_columnas (dict): Diccionario de mapeo de columnas
    """
    print("\nINFO: Verificando integridad de datos:")
    
    # Verificar columnas clave
    columnas_clave = ['ID', 'EXT', 'VOIP', 'Nombre', 'Fecha']
    for col in columnas_clave:
        if col in df_operativo.columns and col in mapeo_columnas:
            valores_vacios = df_operativo[col].isna().sum() + (df_operativo[col] == '').sum()
            total = len(df_operativo)
            pct_completo = (total - valores_vacios) / total * 100 if total > 0 else 0
            
            if pct_completo >= 90:
                status = "OK:"
            elif pct_completo >= 70:
                status = "WARN: "
            else:
                status = "ERROR:"
                
            print(f"  {status} {col}: {pct_completo:.1f}% completo ({total - valores_vacios}/{total})")
            
            # Mostrar muestra de datos para columnas importantes
            if col in ['VOIP', 'EXT'] and valores_vacios < total:
                muestra = df_operativo[col].dropna().head(3).tolist()
                print(f"    Muestra: {muestra}")
    
    # Verificar datos numericos
    columnas_numericas = ['Asignacion', 'PAGOS', 'Total toques', 'Llamadas Microsip', 'Total Llamadas']
    for col in columnas_numericas:
        if col in df_operativo.columns and col in mapeo_columnas:
            try:
                suma_total = pd.to_numeric(df_operativo[col], errors='coerce').sum()
                print(f"  INFO: {col}: Suma total = {suma_total:,.0f}")
            except:
                print(f"  WARN:  {col}: Error al procesar valores numericos")

def crear_hoja_operativo(writer, df_reporte3=None):
    """
    Crea la hoja "Operativo" con la tabla especificada e integra datos del Reporte 3
    """
    # Definir las columnas de la hoja Operativo
    columnas_operativo = [
        "CODIGO", "Tipo Jornada", "Fecha", "Cedula", "ID", "EXT", "VOIP", "Nombre", "Sede", "Ubicacion",
        "Logueo", "Mora", "Asignacion", "Clientes gestionados 11 am", "Capital Asignado", "Capital Recuperado",
        "PAGOS", "% Recuperado", "% Cuentas", "Total toques", "Ultimo Toque", "Llamadas Microsip",
        "Llamadas VOIP", "Total Llamadas", "Gerencia", "Team", "Meta", "Ejecucion", "Ind Logueo",
        "Ind Ultimo", "Ind Ges Medio", "Ind Llamadas", "Indicador Toques", "Ind Pausa", "Total Infracciones", "Total Operativo"
    ]
    
    # Crear DataFrame con datos del Reporte 3 si existe
    if df_reporte3 is not None and not df_reporte3.empty:
        print("Integrando datos del Reporte 3 en hoja Operativo...")
        print(f"Columnas disponibles en Reporte 3: {list(df_reporte3.columns)}")
        
        # Mapeo explicito y mejorado de columnas entre Reporte 3 y Operativo
        mapeo_explicito = {
            # Mapeos directos (nombres exactos)
            'Fecha': 'Fecha',
            'Cedula': 'Cedula', 
            'ID': 'ID',
            'EXT': 'EXT',
            'VOIP': 'VOIP',
            'Nombre': 'Nombre',
            'Sede': 'Sede',
            'Ubicacion': 'Ubicación',  # CORREGIDO: mapear a la version con acento
            'Logueo': 'Logueo',
            'Mora': 'Mora',
            'Asignacion': 'Asignación',  # CORREGIDO: mapear a la version con acento
            'Clientes gestionados 11 am': 'Clientes gestionados 11 am',
            'Capital Asignado': 'Capital Asignado',
            'Capital Recuperado': 'Capital Recuperado',
            'PAGOS': 'PAGOS',
            '% Recuperado': '% Recuperado',
            '% Cuentas': '% Cuentas',
            'Total toques': 'Total toques',
            'Ultimo Toque': 'Ultimo Toque',
            'Llamadas Microsip': 'Llamadas Microsip',
            'Llamadas VOIP': 'Llamadas VOIP',
            'Total Llamadas': 'Total Llamadas',
            'Gerencia': 'Gerencia',
            'Team': 'Team'
        }
        
        
        # Mapeo automatico inteligente
        mapeo_columnas = {}
        
        # Primero aplicar mapeos explicitos
        for col_operativo, col_reporte3_esperado in mapeo_explicito.items():
            if col_reporte3_esperado in df_reporte3.columns:
                mapeo_columnas[col_operativo] = col_reporte3_esperado
                print(f"OK: Mapeo directo: {col_operativo} ← {col_reporte3_esperado}")
        
        # Función auxiliar para normalizar texto y mejorar coincidencias
        def normalizar_texto(texto):
            """Normaliza texto para mejorar coincidencias de mapeo"""
            import unicodedata
            # Quitar acentos y convertir a minúsculas
            texto_normalizado = unicodedata.normalize('NFD', texto)
            texto_sin_acentos = ''.join(c for c in texto_normalizado if unicodedata.category(c) != 'Mn')
            return texto_sin_acentos.lower().strip()
        
        # Luego buscar coincidencias por similitud para columnas no mapeadas
        for col_reporte3 in df_reporte3.columns:
            col_r3_lower = col_reporte3.lower().strip()
            col_r3_normalizado = normalizar_texto(col_reporte3)
            
            for col_operativo in columnas_operativo:
                if col_operativo in mapeo_columnas:
                    continue  # Ya esta mapeado
                    
                col_op_lower = col_operativo.lower().strip()
                col_op_normalizado = normalizar_texto(col_operativo)
                
                # Coincidencias exactas (sin espacios/case)
                if col_r3_lower == col_op_lower:
                    mapeo_columnas[col_operativo] = col_reporte3
                    print(f"OK: Mapeo exacto: {col_operativo} <- {col_reporte3}")
                    continue
                
                # Coincidencias normalizadas (sin acentos)
                if col_r3_normalizado == col_op_normalizado:
                    mapeo_columnas[col_operativo] = col_reporte3
                    print(f"OK: Mapeo normalizado: {col_operativo} <- {col_reporte3}")
                    continue
                
                # Coincidencias parciales inteligentes
                palabras_operativo = set(col_op_normalizado.replace('%', 'porcentaje').split())
                palabras_reporte3 = set(col_r3_normalizado.replace('%', 'porcentaje').split())
                
                # Si al menos 80% de palabras coinciden
                if len(palabras_operativo) > 0:
                    coincidencia = len(palabras_operativo & palabras_reporte3) / len(palabras_operativo)
                    if coincidencia >= 0.8:
                        mapeo_columnas[col_operativo] = col_reporte3
                        print(f"OK: Mapeo por similitud ({coincidencia:.1%}): {col_operativo} <- {col_reporte3}")
        
        print(f"\nMapeo final encontrado ({len(mapeo_columnas)} columnas):")
        for k, v in mapeo_columnas.items():
            print(f"  {k} <- {v}")
        
        # Identificar columnas no mapeadas
        no_mapeadas = [col for col in columnas_operativo if col not in mapeo_columnas]
        if no_mapeadas:
            print(f"\nWARN:  Columnas no mapeadas ({len(no_mapeadas)}): {no_mapeadas}")
        
        # Identificar columnas del Reporte 3 no utilizadas
        utilizadas = set(mapeo_columnas.values())
        no_utilizadas = [col for col in df_reporte3.columns if col not in utilizadas]
        if no_utilizadas:
            print(f"\nINFO: Columnas del Reporte 3 no utilizadas ({len(no_utilizadas)}): {no_utilizadas}")
        
        print(f"\nINFO: Resumen del mapeo:")
        print(f"  - Columnas del Reporte 3: {len(df_reporte3.columns)}")
        print(f"  - Columnas del Operativo: {len(columnas_operativo)}")
        print(f"  - Columnas mapeadas: {len(mapeo_columnas)}")
        print(f"  - Eficiencia de mapeo: {len(mapeo_columnas)/len(columnas_operativo)*100:.1f}%")
        
        # Crear DataFrame con los datos mapeados - MEJORADO
        print("\nINFO: Procesando datos del Reporte 3...")
        print(f"INFO: Registros originales en Reporte 3: {len(df_reporte3)}")
        
        # Crear DataFrame directamente con mapeo de columnas (mas eficiente)
        df_operativo = pd.DataFrame()
        
        for col_operativo in columnas_operativo:
            if col_operativo in mapeo_columnas:
                col_reporte3 = mapeo_columnas[col_operativo]
                # Preservar el tipo de datos original, pero manejar fechas especialmente
                if col_operativo in ['Fecha']:
                    # Para columnas de fecha, copiar directamente (ya estan en formato DD/MM/YYYY)
                    df_operativo[col_operativo] = df_reporte3[col_reporte3].copy()
                    print(f"  INFO: Columna {col_operativo}: copiada directamente con formato DD/MM/YYYY")
                else:
                    df_operativo[col_operativo] = df_reporte3[col_reporte3].copy()
                
                # Verificar si hay valores nulos y reportar
                nulos = df_operativo[col_operativo].isna().sum()
                if nulos > 0:
                    print(f"  WARN:  {col_operativo}: {nulos} valores nulos de {len(df_operativo)}")
            else:
                # Columna no mapeada - inicializar con valores vacios
                df_operativo[col_operativo] = ""
        
        print(f"INFO: Registros procesados en hoja Operativo: {len(df_operativo)}")
        
        # Verificar que no se hayan perdido registros
        if len(df_operativo) != len(df_reporte3):
            print(f"INFO: ADVERTENCIA: Se perdieron {len(df_reporte3) - len(df_operativo)} registros durante el procesamiento")
        else:
            print(f"OK: Todos los registros fueron procesados correctamente")
        
        print(f"OK: Datos integrados: {len(df_operativo)} registros, {len([c for c in columnas_operativo if c in mapeo_columnas])} columnas con datos")
        
        # Verificar integridad de datos clave
        verificar_integridad_datos(df_operativo, mapeo_columnas)
        
    else:
        # Crear DataFrame con una fila de ejemplo si no hay datos del Reporte 3
        datos_ejemplo = [[""] * len(columnas_operativo)]
        df_operativo = pd.DataFrame(datos_ejemplo, columns=columnas_operativo)
        print("No hay datos del Reporte 3, creando hoja con estructura vacia")
    
    # Verificacion final de formatos antes de escribir a Excel
    # Verificación final de datos de fecha
    for col in ['Fecha']:
        if col in df_operativo.columns:
            # Asegurar que las fechas sean strings y estén en formato DD/MM/YYYY
            df_operativo[col] = df_operativo[col].astype(str)
            # Limpiar valores que no sean fechas válidas
            df_operativo[col] = df_operativo[col].replace('NaT', '').replace('nan', '')
            print(f"INFO: Formato de fechas verificado para columna {col}")
            df_operativo[col] = df_operativo[col].replace(['nan', 'NaT', 'None'], '')
            
            # Validar y corregir formato DD/MM/YYYY si es necesario
            import re
            for idx, fecha_val in enumerate(df_operativo[col]):
                if fecha_val and fecha_val != '':
                    # Verificar si esta en formato MM/DD/YYYY y convertir a DD/MM/YYYY
                    if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', fecha_val):
                        try:
                            # Intentar parsear como DD/MM/YYYY primero
                            fecha_dt = pd.to_datetime(fecha_val, format='%d/%m/%Y', errors='raise')
                            df_operativo.iloc[idx, df_operativo.columns.get_loc(col)] = fecha_dt.strftime('%d/%m/%Y')
                        except:
                            try:
                                # Si falla, intentar como MM/DD/YYYY y convertir a DD/MM/YYYY
                                fecha_dt = pd.to_datetime(fecha_val, format='%m/%d/%Y', errors='raise')
                                df_operativo.iloc[idx, df_operativo.columns.get_loc(col)] = fecha_dt.strftime('%d/%m/%Y')
                                print(f"    - Corregida fecha {fecha_val} → {fecha_dt.strftime('%d/%m/%Y')}")
                            except:
                                print(f"    - No se pudo parsear fecha: {fecha_val}")
            
            print(f"  - Despues de correccion - Valores vacios: {df_operativo[col].eq('').sum()}")
            print(f"  - Muestra final: {df_operativo[col].head(3).tolist()}")
            print(f"OK: Columna {col} convertida y validada en formato DD/MM/YYYY")
    
    # Escribir a Excel
    df_operativo.to_excel(writer, sheet_name="Operativo", index=False)
    
    # Obtener la hoja y aplicar formato basico
    worksheet = writer.sheets["Operativo"]
    
    # AGREGAR FORMULAS EN LA COLUMNA CODIGO (primera columna)
    # Encontrar las posiciones de las columnas ID y Fecha
    id_col = None
    fecha_col = None
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "ID":
            id_col = i + 1  # Excel columnas empiezan en 1
        elif col_name == "Fecha":
            fecha_col = i + 1
    
    if id_col and fecha_col:
        from openpyxl.utils import get_column_letter
        id_letter = get_column_letter(id_col)
        fecha_letter = get_column_letter(fecha_col)
        
        # Agregar formulas en la columna CODIGO para cada fila de datos
        # Formato: ID + DDMM (ej: 2020 + 0908 = 20200908)
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            # Formula Excel para extraer dia y mes de fecha DD/MM/YYYY con validacion
            # Si fecha es 09/08/2025, resultado debe ser: ID + "0908"
            formula = f'=IF(AND({id_letter}{row}<>"",{fecha_letter}{row}<>""),{id_letter}{row}&TEXT(DAY(DATEVALUE({fecha_letter}{row})),"00")&TEXT(MONTH(DATEVALUE({fecha_letter}{row})),"00"),"")'
            cell = worksheet.cell(row=row, column=1)  # Columna A (CODIGO)
            cell.value = formula
        print(f"Formulas agregadas en columna CODIGO: {len(df_operativo)} formulas con DATEVALUE, validacion de celdas vacias y formato ID+DDMM")
        
        # AGREGAR FORMULAS EN LA COLUMNA TIPO JORNADA (segunda columna)
        # Encontrar la posicion de la columna Tipo Jornada
        tipo_jornada_col = None
        for i, col_name in enumerate(columnas_operativo):
            if col_name == "Tipo Jornada":
                tipo_jornada_col = i + 1  # Excel columnas empiezan en 1
                break
        
        if tipo_jornada_col and fecha_col:
            tipo_jornada_letter = get_column_letter(tipo_jornada_col)
            
            # Agregar formulas en la columna Tipo Jornada para cada fila de datos
            # Si dia es 30, 31, 1, 2, 15, 16, 17 → "Pago", sino → "Normal"
            for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
                formula = f'=IF(OR(DAY({fecha_letter}{row})=30,DAY({fecha_letter}{row})=31,DAY({fecha_letter}{row})=1,DAY({fecha_letter}{row})=2,DAY({fecha_letter}{row})=15,DAY({fecha_letter}{row})=16,DAY({fecha_letter}{row})=17),"Pago","Normal")'
                cell = worksheet.cell(row=row, column=tipo_jornada_col)
                cell.value = formula
            print(f"Formulas agregadas en columna Tipo Jornada: {len(df_operativo)} formulas con logica Pago/Normal (dias: 30,31,1,2,15,16,17)")
            
            # Agregar validacion de datos (dropdown) en la columna Tipo Jornada
            from openpyxl.worksheet.datavalidation import DataValidation
            
            # Crear validacion con lista desplegable
            dv = DataValidation(type="list", formula1='"Pago,Normal"', allow_blank=False)
            dv.error = 'Seleccione: Pago o Normal'
            dv.errorTitle = 'Valor invalido'
            dv.prompt = 'Seleccione el tipo de jornada'
            dv.promptTitle = 'Tipo Jornada'
            
            # Aplicar validacion a todas las filas de datos
            range_validation = f"{tipo_jornada_letter}2:{tipo_jornada_letter}{len(df_operativo) + 1}"
            dv.add(range_validation)
            worksheet.add_data_validation(dv)
            print(f"Validacion de datos agregada en columna Tipo Jornada: rango {range_validation}")
        else:
            print("ERROR: No se encontro columna Tipo Jornada para las formulas")
            
        # AGREGAR FORMULAS EN LA COLUMNA META
        # Encontrar las posiciones de las columnas necesarias
        meta_col = None
        mora_col = None
        capital_asignado_col = None
        for i, col_name in enumerate(columnas_operativo):
            if col_name == "Meta":
                meta_col = i + 1  # Excel columnas empiezan en 1
            elif col_name == "Mora":
                mora_col = i + 1
            elif col_name == "Capital Asignado":
                capital_asignado_col = i + 1
        
        if meta_col and mora_col and tipo_jornada_col and capital_asignado_col:
            meta_letter = get_column_letter(meta_col)
            mora_letter = get_column_letter(mora_col)
            capital_asignado_letter = get_column_letter(capital_asignado_col)
            
            # Agregar formulas en la columna Meta para cada fila de datos
            for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
                # Formula con validacion de celdas vacias:
                # 1. Verificar que Mora, Capital Asignado y Tipo Jornada no esten vacios
                # 2. Si Mora = "M1-2" → calcular 500000 / Capital Asignado como porcentaje
                # 3. Si Tipo Jornada = "Pago" → buscar en Planta!E:F (Dia Pago)
                # 4. Si Tipo Jornada = "Normal" → buscar en Planta!H:I (Dia Normal)
                formula = f'=IF(AND({mora_letter}{row}<>"",{capital_asignado_letter}{row}<>"",{tipo_jornada_letter}{row}<>""),IF({mora_letter}{row}="M1-2",ROUND(500000/{capital_asignado_letter}{row},4),IF({tipo_jornada_letter}{row}="Pago",VLOOKUP({mora_letter}{row},Planta!E:F,2,FALSE),IF({tipo_jornada_letter}{row}="Normal",VLOOKUP({mora_letter}{row},Planta!H:I,2,FALSE),""))),"")'
                cell = worksheet.cell(row=row, column=meta_col)
                cell.value = formula
            print(f"Formulas agregadas en columna Meta: {len(df_operativo)} formulas con validacion de celdas vacias, VLOOKUP a tablas de Planta y calculo especial para M1-2")
            
            # AGREGAR FORMULAS EN LA COLUMNA EJECUCION
            # Encontrar las posiciones de las columnas necesarias para Ejecucion
            ejecucion_col = None
            percent_recuperado_col = None
            percent_cuentas_col = None
            for i, col_name in enumerate(columnas_operativo):
                if col_name == "Ejecucion":
                    ejecucion_col = i + 1
                elif col_name == "% Recuperado":
                    percent_recuperado_col = i + 1
                elif col_name == "% Cuentas":
                    percent_cuentas_col = i + 1
            
            if ejecucion_col and percent_recuperado_col and percent_cuentas_col:
                ejecucion_letter = get_column_letter(ejecucion_col)
                percent_recuperado_letter = get_column_letter(percent_recuperado_col)
                percent_cuentas_letter = get_column_letter(percent_cuentas_col)
                
                # Agregar formulas en la columna Ejecucion para cada fila de datos
                for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
                    # Formula con validacion de celdas vacias:
                    # 1. Verificar que Mora, Meta, % Recuperado y % Cuentas no esten vacios
                    # 2. Si Mora="M0-PP" → % Cuentas / Meta, sino → % Recuperado / Meta
                    formula = f'=IF(AND({mora_letter}{row}<>"",{meta_letter}{row}<>"",{percent_recuperado_letter}{row}<>"",{percent_cuentas_letter}{row}<>""),IF({mora_letter}{row}="M0-PP",{percent_cuentas_letter}{row}/{meta_letter}{row},{percent_recuperado_letter}{row}/{meta_letter}{row}),"")'
                    cell = worksheet.cell(row=row, column=ejecucion_col)
                    cell.value = formula
                print(f"Formulas agregadas en columna Ejecucion: {len(df_operativo)} formulas con validacion de celdas vacias y logica M0-PP")
                
                # AGREGAR FORMULAS EN LA COLUMNA IND LOGUEO
                # Encontrar las posiciones de las columnas necesarias para Ind Logueo
                ind_logueo_col = None
                logueo_col = None
                for i, col_name in enumerate(columnas_operativo):
                    if col_name == "Ind Logueo":
                        ind_logueo_col = i + 1
                    elif col_name == "Logueo":
                        logueo_col = i + 1
                
                if ind_logueo_col and logueo_col:
                    ind_logueo_letter = get_column_letter(ind_logueo_col)
                    logueo_letter = get_column_letter(logueo_col)
                    
                    # Agregar formulas en la columna Ind Logueo para cada fila de datos
                    for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
                        # Formula con validacion de celdas vacias:
                        # 1. Verificar que Logueo y Tipo Jornada no esten vacios
                        # 2. Si Tipo Jornada="Pago" y Logueo≤7:30 AM → 0.15
                        # 3. Si Tipo Jornada="Normal" y Logueo≤8:00 AM → 0.15, sino → 0
                        formula = f'=IF(AND({logueo_letter}{row}<>"",{tipo_jornada_letter}{row}<>""),IF({tipo_jornada_letter}{row}="Pago",IF(TIMEVALUE({logueo_letter}{row})<=TIMEVALUE("7:30:00 AM"),0.15,0),IF({tipo_jornada_letter}{row}="Normal",IF(TIMEVALUE({logueo_letter}{row})<=TIMEVALUE("8:00:00 AM"),0.15,0),0)),0)'
                        cell = worksheet.cell(row=row, column=ind_logueo_col)
                        cell.value = formula
                    print(f"Formulas agregadas en columna Ind Logueo: {len(df_operativo)} formulas con TIMEVALUE para formato AM/PM")
                    
                    # AGREGAR FORMULAS EN LA COLUMNA IND ULTIMO
                    # Encontrar las posiciones de las columnas necesarias para Ind Ultimo
                    ind_ultimo_col = None
                    ultimo_toque_col = None
                    fecha_col = None
                    for i, col_name in enumerate(columnas_operativo):
                        if col_name == "Ind Ultimo":
                            ind_ultimo_col = i + 1
                        elif col_name == "Ultimo Toque":
                            ultimo_toque_col = i + 1
                        elif col_name == "Fecha":
                            fecha_col = i + 1
                    
                    if ind_ultimo_col and ultimo_toque_col and fecha_col:
                        ind_ultimo_letter = get_column_letter(ind_ultimo_col)
                        ultimo_toque_letter = get_column_letter(ultimo_toque_col)
                        fecha_letter = get_column_letter(fecha_col)
                        
                        # Agregar formulas en la columna Ind Ultimo para cada fila de datos
                        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
                            # Formula con validacion de celdas vacias y conversion de fecha:
                            # 1. Verificar que Fecha, Ultimo Toque y Tipo Jornada no esten vacios
                            # 2. Convertir fecha DD/MM/YYYY a formato que Excel entienda para WEEKDAY
                            # 3. Si es sabado: Ultimo Toque >= 12:20 PM → 0.15
                            # 4. Si es pago (no sabado): Ultimo Toque >= 6:50 PM → 0.15  
                            # 5. Si es normal (no sabado): Ultimo Toque >= 5:20 PM → 0.15
                            # 6. Sino → 0
                            formula = f'=IF(AND({fecha_letter}{row}<>"",{ultimo_toque_letter}{row}<>"",{tipo_jornada_letter}{row}<>""),IF(WEEKDAY(DATEVALUE({fecha_letter}{row}))=7,IF(TIMEVALUE({ultimo_toque_letter}{row})>=TIMEVALUE("12:20:00 PM"),0.15,0),IF({tipo_jornada_letter}{row}="Pago",IF(TIMEVALUE({ultimo_toque_letter}{row})>=TIMEVALUE("6:50:00 PM"),0.15,0),IF({tipo_jornada_letter}{row}="Normal",IF(TIMEVALUE({ultimo_toque_letter}{row})>=TIMEVALUE("5:20:00 PM"),0.15,0),0))),0)'
                            cell = worksheet.cell(row=row, column=ind_ultimo_col)
                            cell.value = formula
                        print(f"Formulas agregadas en columna Ind Ultimo: {len(df_operativo)} formulas con DATEVALUE para conversion de fecha, validacion de celdas vacias, sabados (12:20 PM) y horarios PM")
                    else:
                        print("ERROR: No se encontraron columnas Ind Ultimo, Ultimo Toque y/o Fecha para las formulas")
                else:
                    print("ERROR: No se encontraron columnas Ind Logueo y/o Logueo para las formulas")
            else:
                print("ERROR: No se encontraron columnas Ejecucion, % Recuperado y/o % Cuentas para las formulas")
        else:
            print("ERROR: No se encontraron columnas Meta, Mora, Tipo Jornada y/o Capital Asignado para las formulas")
    else:
        print("ERROR: No se encontraron columnas ID y/o Fecha para las formulas")
    
    # AGREGAR VALORES FIJOS EN LA COLUMNA IND PAUSA
    # Encontrar la posicion de la columna Ind Pausa
    ind_pausa_col = None
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "Ind Pausa":
            ind_pausa_col = i + 1
            break
    
    if ind_pausa_col:
        # Agregar valor fijo 0.15 en la columna Ind Pausa para cada fila de datos
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            cell = worksheet.cell(row=row, column=ind_pausa_col)
            cell.value = 0.15
        print(f"Valores fijos 0.15 agregados en columna Ind Pausa: {len(df_operativo)} filas")
    else:
        print("ERROR: No se encontro columna Ind Pausa")
    
    # AGREGAR FORMULAS EN LA COLUMNA IND GES MEDIO
    # Encontrar las posiciones de las columnas necesarias para Ind Ges Medio
    ind_ges_medio_col = None
    asignacion_col = None
    clientes_gestionados_col = None
    
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "Ind Ges Medio":
            ind_ges_medio_col = i + 1
        elif col_name == "Asignacion":
            asignacion_col = i + 1
        elif col_name == "Clientes gestionados 11 am":
            clientes_gestionados_col = i + 1
    
    if ind_ges_medio_col and asignacion_col and clientes_gestionados_col:
        ind_ges_medio_letter = get_column_letter(ind_ges_medio_col)
        asignacion_letter = get_column_letter(asignacion_col)
        clientes_gestionados_letter = get_column_letter(clientes_gestionados_col)
        
        # Agregar formulas en la columna Ind Ges Medio para cada fila de datos
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            # Formula con validacion de celdas vacias y logica condicional:
            # 1. Verificar que Asignacion y Clientes gestionados no esten vacios
            # 2. Si Asignacion < 45: Clientes gestionados >= ROUND(Asignacion*0.9,0) → 0.15, sino → 0
            # 3. Si Asignacion >= 45: Clientes gestionados >= 45 → 0.15, sino → 0
            # Aplicar redondeo al 90% de asignacion: 35.1→35, 35.6→36
            formula = f'=IF(AND({asignacion_letter}{row}<>"",{clientes_gestionados_letter}{row}<>""),IF({asignacion_letter}{row}<45,IF({clientes_gestionados_letter}{row}>=ROUND({asignacion_letter}{row}*0.9,0),0.15,0),IF({clientes_gestionados_letter}{row}>=45,0.15,0)),0)'
            cell = worksheet.cell(row=row, column=ind_ges_medio_col)
            cell.value = formula
        print(f"Formulas agregadas en columna Ind Ges Medio: {len(df_operativo)} formulas con validacion de gestion media (<45: >=90% Asignacion redondeado, >=45: >=45)")
    else:
        print("ERROR: No se encontraron columnas Ind Ges Medio, Asignacion y/o Clientes gestionados 11 am")
    
    # AGREGAR FORMULAS EN LA COLUMNA INDICADOR TOQUES
    # Encontrar las posiciones de las columnas necesarias para Indicador Toques
    indicador_toques_col = None
    mora_col = None
    total_toques_col = None
    
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "Indicador Toques":
            indicador_toques_col = i + 1
        elif col_name == "Mora":
            mora_col = i + 1
        elif col_name == "Total toques":
            total_toques_col = i + 1
    
    if indicador_toques_col and mora_col and total_toques_col:
        mora_letter = get_column_letter(mora_col)
        total_toques_letter = get_column_letter(total_toques_col)
        
        # Agregar formulas en la columna Indicador Toques para cada fila de datos
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            # Formula con validacion de cartera que CONTENGA M0 vs otras carteras:
            # 1. Verificar que Mora y Total toques no esten vacios
            # 2. Si Mora CONTIENE "M0": Total toques >= 120 → 0.2, sino → 0
            # 3. Si Mora NO CONTIENE "M0": Total toques >= 160 → 0.2, sino → 0
            # Usar ISNUMBER(SEARCH()) para buscar "M0" dentro del texto
            formula = f'=IF(AND({mora_letter}{row}<>"",{total_toques_letter}{row}<>""),IF(ISNUMBER(SEARCH("M0",{mora_letter}{row})),IF({total_toques_letter}{row}>=120,0.2,0),IF({total_toques_letter}{row}>=160,0.2,0)),0)'
            cell = worksheet.cell(row=row, column=indicador_toques_col)
            cell.value = formula
        print(f"Formulas agregadas en columna Indicador Toques: {len(df_operativo)} formulas con validacion de toques (CONTIENE M0: >=120, Otras: >=160)")
    else:
        print("ERROR: No se encontraron columnas Indicador Toques, Mora y/o Total toques")
    
    # AGREGAR FORMULAS EN LA COLUMNA IND LLAMADAS
    # Encontrar las posiciones de las columnas necesarias para Ind Llamadas
    ind_llamadas_col = None
    total_llamadas_col = None
    
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "Ind Llamadas":
            ind_llamadas_col = i + 1
        elif col_name == "Total Llamadas":
            total_llamadas_col = i + 1
    
    if ind_llamadas_col and total_llamadas_col:
        total_llamadas_letter = get_column_letter(total_llamadas_col)
        
        # Agregar formulas en la columna Ind Llamadas para cada fila de datos
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            # Formula con validacion de Total Llamadas >= 150:
            # 1. Verificar que Total Llamadas no este vacio
            # 2. Si Total Llamadas >= 150 → 0.2, sino → 0
            formula = f'=IF({total_llamadas_letter}{row}<>"",IF({total_llamadas_letter}{row}>=150,0.2,0),0)'
            cell = worksheet.cell(row=row, column=ind_llamadas_col)
            cell.value = formula
        print(f"Formulas agregadas en columna Ind Llamadas: {len(df_operativo)} formulas con validacion de llamadas (>=150)")
    else:
        print("ERROR: No se encontraron columnas Ind Llamadas y/o Total Llamadas")
    
    # AGREGAR FORMULAS EN LA COLUMNA TOTAL INFRACCIONES
    # Encontrar las posiciones de las columnas necesarias
    total_infracciones_col = None
    ind_logueo_col = None
    ind_pausa_col = None
    
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "Total Infracciones":
            total_infracciones_col = i + 1
        elif col_name == "Ind Logueo":
            ind_logueo_col = i + 1
        elif col_name == "Ind Pausa":
            ind_pausa_col = i + 1
    
    if total_infracciones_col and ind_logueo_col and ind_pausa_col:
        ind_logueo_letter = get_column_letter(ind_logueo_col)
        ind_pausa_letter = get_column_letter(ind_pausa_col)
        
        # Agregar formulas en la columna Total Infracciones para cada fila de datos
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            # Formula: COUNTIF del rango Ind Logueo hasta Ind Pausa donde valor = 0
            formula = f'=COUNTIF({ind_logueo_letter}{row}:{ind_pausa_letter}{row},0)'
            cell = worksheet.cell(row=row, column=total_infracciones_col)
            cell.value = formula
        print(f"Formulas agregadas en columna Total Infracciones: {len(df_operativo)} formulas COUNTIF para contar ceros")
    else:
        print("ERROR: No se encontraron columnas Total Infracciones, Ind Logueo y/o Ind Pausa")
    
    # AGREGAR FORMULAS EN LA COLUMNA TOTAL OPERATIVO
    # Encontrar la posicion de la columna Total Operativo
    total_operativo_col = None
    for i, col_name in enumerate(columnas_operativo):
        if col_name == "Total Operativo":
            total_operativo_col = i + 1
            break
    
    if total_operativo_col and ind_logueo_col and ind_pausa_col:
        total_operativo_letter = get_column_letter(total_operativo_col)
        
        # Agregar formulas en la columna Total Operativo para cada fila de datos
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (despues del header)
            # Formula: SUM del rango Ind Logueo hasta Ind Pausa (resultado en porcentaje)
            formula = f'=SUM({ind_logueo_letter}{row}:{ind_pausa_letter}{row})'
            cell = worksheet.cell(row=row, column=total_operativo_col)
            cell.value = formula
        print(f"Formulas agregadas en columna Total Operativo: {len(df_operativo)} formulas SUM (formato porcentaje)")
    else:
        print("ERROR: No se encontraron columnas Total Operativo, Ind Logueo y/o Ind Pausa")
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, len(columnas_operativo) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Aplicar formato de porcentaje a las columnas % Recuperado y % Cuentas
    from openpyxl.utils import get_column_letter
    
    # Crear diccionario de columnas para encontrar las posiciones
    column_letters = {}
    for col_num, col_name in enumerate(columnas_operativo, 1):
        column_letters[col_name] = get_column_letter(col_num)
    
    # Formato de porcentaje para las columnas de porcentaje
    if '% Recuperado' in column_letters and '% Cuentas' in column_letters:
        percent_format = '0.00%'
        # Aplicar formato a todas las filas de datos existentes
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            # Columna % Recuperado
            percent_recup_cell = f"{column_letters['% Recuperado']}{row_num}"
            worksheet[percent_recup_cell].number_format = percent_format
            
            # Columna % Cuentas
            percent_cuentas_cell = f"{column_letters['% Cuentas']}{row_num}"
            worksheet[percent_cuentas_cell].number_format = percent_format
        print(f"OK: Formato de porcentaje aplicado a {max_row-1} filas en columnas % Recuperado y % Cuentas")
    
    # Aplicar formato de porcentaje a la columna Meta
    if 'Meta' in column_letters:
        percent_format = '0.00%'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            meta_cell = f"{column_letters['Meta']}{row_num}"
            worksheet[meta_cell].number_format = percent_format
        print(f"OK: Formato de porcentaje aplicado a {max_row-1} filas en columna Meta")
    
    # Aplicar formato de porcentaje a la columna Ejecucion
    if 'Ejecucion' in column_letters:
        percent_format = '0.00%'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            ejecucion_cell = f"{column_letters['Ejecucion']}{row_num}"
            worksheet[ejecucion_cell].number_format = percent_format
        print(f"OK: Formato de porcentaje aplicado a {max_row-1} filas en columna Ejecucion")
    
    # Aplicar formato de porcentaje a la columna Total Operativo
    if 'Total Operativo' in column_letters:
        percent_format = '0.00%'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            total_operativo_cell = f"{column_letters['Total Operativo']}{row_num}"
            worksheet[total_operativo_cell].number_format = percent_format
        print(f"OK: Formato de porcentaje aplicado a {max_row-1} filas en columna Total Operativo")
    
    # Aplicar formato numerico a la columna Ind Logueo
    if 'Ind Logueo' in column_letters:
        decimal_format = '0.00'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            ind_logueo_cell = f"{column_letters['Ind Logueo']}{row_num}"
            worksheet[ind_logueo_cell].number_format = decimal_format
        print(f"OK: Formato decimal aplicado a {max_row-1} filas en columna Ind Logueo")
    
    # Aplicar formato numerico a la columna Ind Ultimo
    if 'Ind Ultimo' in column_letters:
        decimal_format = '0.00'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            ind_ultimo_cell = f"{column_letters['Ind Ultimo']}{row_num}"
            worksheet[ind_ultimo_cell].number_format = decimal_format
        print(f"OK: Formato decimal aplicado a {max_row-1} filas en columna Ind Ultimo")
    
    # Aplicar formato numerico a la columna Ind Pausa
    if 'Ind Pausa' in column_letters:
        decimal_format = '0.00'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            ind_pausa_cell = f"{column_letters['Ind Pausa']}{row_num}"
            worksheet[ind_pausa_cell].number_format = decimal_format
        print(f"OK: Formato decimal aplicado a {max_row-1} filas en columna Ind Pausa")
    
    # Aplicar formato numerico a la columna Ind Ges Medio
    if 'Ind Ges Medio' in column_letters:
        decimal_format = '0.00'
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            ind_ges_medio_cell = f"{column_letters['Ind Ges Medio']}{row_num}"
            worksheet[ind_ges_medio_cell].number_format = decimal_format
        print(f"OK: Formato decimal aplicado a {max_row-1} filas en columna Ind Ges Medio")
    
    # Aplicar formato numerico a la columna Total Infracciones
    if 'Total Infracciones' in column_letters:
        integer_format = '0'  # Formato entero sin decimales
        max_row = max(len(df_operativo) + 1, 100)  # Minimo 100 filas para permitir datos manuales
        for row_num in range(2, max_row + 1):  # Empezar en fila 2 (despues del header)
            total_infracciones_cell = f"{column_letters['Total Infracciones']}{row_num}"
            worksheet[total_infracciones_cell].number_format = integer_format
        print(f"OK: Formato entero aplicado a {max_row-1} filas en columna Total Infracciones")
    # Ajustar ancho de columnas
    from openpyxl.utils import get_column_letter
    anchos_operativo = [10, 12, 10, 12, 8, 8, 8, 20, 10, 12, 10, 8, 12, 18, 15, 15, 10, 12, 12, 12, 15, 15, 15, 15, 12, 10, 8, 12, 12, 12, 12, 12, 15, 12, 18, 15]
    for i, ancho in enumerate(anchos_operativo, 1):
        worksheet.column_dimensions[get_column_letter(i)].width = ancho
    
    # Aplicar formato de tabla
    aplicar_formato_tabla(worksheet, df_operativo, "TablaOperativo")

def crear_hoja_team(writer):
    """
    Crea la hoja "Team" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_team = [
        "Codigo Aus", "Fecha", "Usuario", "Cedula", "Asistencia", "Asesores", 
        "Monitoreos", "% Calidad", "Infracciones", "% Operativo", "Cargo"
    ]
    
    # Crear DataFrame con una fila de ejemplo
    datos_ejemplo = [[""] * len(columnas_team)]
    df_team = pd.DataFrame(datos_ejemplo, columns=columnas_team)
    
    # Escribir a Excel
    df_team.to_excel(writer, sheet_name="Team", index=False)
    
    # Obtener la hoja y aplicar formato basico
    worksheet = writer.sheets["Team"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="9B59B6", end_color="9B59B6", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, len(columnas_team) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 12  # Codigo Aus
    worksheet.column_dimensions['B'].width = 12  # Fecha
    worksheet.column_dimensions['C'].width = 20  # Usuario
    worksheet.column_dimensions['D'].width = 12  # Cedula
    worksheet.column_dimensions['E'].width = 12  # Asistencia
    worksheet.column_dimensions['F'].width = 10  # Asesores
    worksheet.column_dimensions['G'].width = 12  # Monitoreos
    worksheet.column_dimensions['H'].width = 12  # % Calidad
    worksheet.column_dimensions['I'].width = 15  # Infracciones
    worksheet.column_dimensions['J'].width = 15  # % Operativo
    worksheet.column_dimensions['K'].width = 10  # Cargo
    
    # Aplicar formato de tabla
    aplicar_formato_tabla(worksheet, df_team, "TablaTeam")

def crear_hoja_gerente(writer):
    """
    Crea la hoja "Gerente" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_gerente = [
        "Codigo Aus", "Fecha", "Usuario", "Cedula", "Asistencia", "Asesores", 
        "Monitoreos", "% Calidad", "Infracciones", "% Operativo", "Cargo"
    ]
    
    # Crear DataFrame con una fila de ejemplo
    datos_ejemplo = [[""] * len(columnas_gerente)]
    df_gerente = pd.DataFrame(datos_ejemplo, columns=columnas_gerente)
    
    # Escribir a Excel
    df_gerente.to_excel(writer, sheet_name="Gerente", index=False)
    
    # Obtener la hoja y aplicar formato basico
    worksheet = writer.sheets["Gerente"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2E86AB", end_color="2E86AB", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, len(columnas_gerente) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 12  # Codigo Aus
    worksheet.column_dimensions['B'].width = 12  # Fecha
    worksheet.column_dimensions['C'].width = 20  # Usuario
    worksheet.column_dimensions['D'].width = 12  # Cedula
    worksheet.column_dimensions['E'].width = 12  # Asistencia
    worksheet.column_dimensions['F'].width = 10  # Asesores
    worksheet.column_dimensions['G'].width = 12  # Monitoreos
    worksheet.column_dimensions['H'].width = 12  # % Calidad
    worksheet.column_dimensions['I'].width = 15  # Infracciones
    worksheet.column_dimensions['J'].width = 15  # % Operativo
    worksheet.column_dimensions['K'].width = 10  # Cargo
    
    # Aplicar formato de tabla
    aplicar_formato_tabla(worksheet, df_gerente, "TablaGerente")

def crear_hoja_consolidado(writer):
    """
    Crea la hoja "Consolidado" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_consolidado = [
        "Codigo_Asis", "CODIGO", "Tipo Jornada", "Fecha", "Cedula", "ID", "Nombre", "Sede", "Ubicacion",
        "Asistencia", "Mora", "Monitoreos", "Nota Calidad", "Ejecucion", "# Infracciones", 
        "% Operativo", "Team", "Gerente"
    ]
    
    # Crear DataFrame con una fila de ejemplo
    datos_ejemplo = [[""] * len(columnas_consolidado)]
    df_consolidado = pd.DataFrame(datos_ejemplo, columns=columnas_consolidado)
    
    # Escribir a Excel
    df_consolidado.to_excel(writer, sheet_name="Consolidado", index=False)
    
    # Obtener la hoja y aplicar formato basico
    worksheet = writer.sheets["Consolidado"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="34495E", end_color="34495E", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, len(columnas_consolidado) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 12  # Codigo_Asis
    worksheet.column_dimensions['B'].width = 10  # CODIGO
    worksheet.column_dimensions['C'].width = 15  # Tipo Jornada
    worksheet.column_dimensions['D'].width = 12  # Fecha
    worksheet.column_dimensions['E'].width = 12  # Cedula
    worksheet.column_dimensions['F'].width = 8   # ID
    worksheet.column_dimensions['G'].width = 20  # Nombre
    worksheet.column_dimensions['H'].width = 10  # Sede
    worksheet.column_dimensions['I'].width = 12  # Ubicacion
    worksheet.column_dimensions['J'].width = 12  # Asistencia
    worksheet.column_dimensions['K'].width = 8   # Mora
    worksheet.column_dimensions['L'].width = 12  # Monitoreos
    worksheet.column_dimensions['M'].width = 15  # Nota Calidad
    worksheet.column_dimensions['N'].width = 12  # Ejecucion
    worksheet.column_dimensions['O'].width = 15  # # Infracciones
    worksheet.column_dimensions['P'].width = 15  # % Operativo
    worksheet.column_dimensions['Q'].width = 10  # Team
    worksheet.column_dimensions['R'].width = 12  # Gerente
    
    # Aplicar formato de tabla
    aplicar_formato_tabla(worksheet, df_consolidado, "TablaConsolidado")
