"""
Módulo para el procesamiento del 4_Reporte_Calidad - NUEVA VERSIÓN
Sistema simplificado para el reporte de calidad
"""

import pandas as pd
import io
import os
from datetime import datetime
from flask import request, jsonify, send_file
from utils.file_utils import allowed_file

def procesar_reporte_calidad():
    """
    Función principal para procesar el reporte de calidad
    """
    try:
        # Verificar si hay archivo automático del paso 3
        reporte3_auto_file = request.form.get('reporte3_auto_file')
        archivo_reporte3_content = None
        
        if reporte3_auto_file:
            # Usar archivo temporal del paso 3
            temp_filepath = os.path.join('temp_files', reporte3_auto_file)
            if os.path.exists(temp_filepath):
                with open(temp_filepath, 'rb') as f:
                    file_content = f.read()
                archivo_reporte3_content = io.BytesIO(file_content)
            else:
                return jsonify({
                    'success': False,
                    'message': 'Archivo temporal del Paso 3 no encontrado'
                }), 400
        else:
            # Usar archivo manual
            if 'excelFileReporte3' not in request.files:
                return jsonify({
                    'success': False,
                    'message': 'Se requiere el archivo de Reporte 3'
                }), 400
            
            archivo_reporte3 = request.files['excelFileReporte3']
            if archivo_reporte3.filename == '':
                return jsonify({
                    'success': False,
                    'message': 'Archivo de Reporte 3 sin nombre válido'
                }), 400
            
            archivo_reporte3_content = archivo_reporte3
        
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
        return jsonify({
            'success': False,
            'message': f'Error procesando reporte de calidad: {str(e)}'
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
                df_reporte3 = pd.read_excel(archivo_reporte3, sheet_name=0)
                print(f"Datos del Reporte 3 cargados: {len(df_reporte3)} registros")
                print(f"Columnas disponibles: {list(df_reporte3.columns)}")
            except Exception as e:
                print(f"Error leyendo Reporte 3: {str(e)}")
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
            
            # Si hay archivo del reporte 3, agregarlo como hoja adicional
            if archivo_reporte3:
                try:
                    df_reporte3 = pd.read_excel(archivo_reporte3, sheet_name=0)
                    df_reporte3.to_excel(writer, sheet_name='Datos_Reporte3', index=False)
                except Exception:
                    pass  # Si no se puede leer, continuar sin esta hoja
        
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
                'hojas_creadas': ['Consolidado', 'Gerente', 'Team', 'Operativo', 'Calidad', 'Ausentismo', 'Asistencia Lideres', 'Planta'] + (['Datos_Reporte3'] if archivo_reporte3 else []),
                'tablas_planta': 3,
                'total_hojas': 9 if archivo_reporte3 else 8
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
        ["Nicolas Briceño", "1013097231", "Back"],
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
    
    # Tabla 2: Día Pago
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
    
    # Tabla 3: Día Normal
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
    
    # Obtener la hoja después de crear la primera tabla
    worksheet = writer.sheets["Planta"]
    
    # Escribir las otras tablas manualmente para control de posición
    # Tabla 2: Día Pago (E1:F12)
    worksheet.cell(row=1, column=5, value="Dia Pago")
    worksheet.cell(row=1, column=6, value="Meta")
    for idx, row in enumerate(datos_dia_pago, start=2):
        worksheet.cell(row=idx, column=5, value=row[0])
        worksheet.cell(row=idx, column=6, value=row[1])
    
    # Tabla 3: Día Normal (H1:I12)
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
    
    # Tabla 2: Día Pago  
    tabla_dia_pago = Table(displayName="TablaDiaPago", ref="E1:F12")
    style_dia_pago = TableStyleInfo(name="TableStyleMedium15", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tabla_dia_pago.tableStyleInfo = style_dia_pago
    worksheet.add_table(tabla_dia_pago)
    
    # Tabla 3: Día Normal
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
    
    # Obtener la hoja y aplicar formato básico
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

def crear_hoja_ausentismo(writer):
    """
    Crea la hoja "Ausentismo" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_ausentismo = [
        "Codigo Aus", "Codigo", "Tipo Jornada", "Fecha", "Cedula", "ID",
        "Nombre", "Sede", "Ubicación", "Logueo Admin", "Ingreso", 
        "Salida", "Horas laboradas", "Novedad Ingreso", "Validación", "Drive"
    ]
    
    # Crear DataFrame con una fila de ejemplo para evitar errores de tabla
    datos_ejemplo = [["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]]
    df_ausentismo = pd.DataFrame(datos_ejemplo, columns=columnas_ausentismo)
    
    # Escribir a Excel
    df_ausentismo.to_excel(writer, sheet_name="Ausentismo", index=False)
    
    # Obtener la hoja y aplicar formato básico
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
    worksheet.column_dimensions['I'].width = 12  # Ubicación
    worksheet.column_dimensions['J'].width = 15  # Logueo Admin
    worksheet.column_dimensions['K'].width = 10  # Ingreso
    worksheet.column_dimensions['L'].width = 10  # Salida
    worksheet.column_dimensions['M'].width = 15  # Horas laboradas
    worksheet.column_dimensions['N'].width = 18  # Novedad Ingreso
    worksheet.column_dimensions['O'].width = 12  # Validación
    worksheet.column_dimensions['P'].width = 10  # Drive

def crear_hoja_calidad(writer):
    """
    Crea la hoja "Calidad" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_calidad = [
        "Codigo", "Fecha Monitoreo", "ID Asesor", "VOZ", "SMS", 
        "TERCERO", "Nota Total", "Total Monitoreos"
    ]
    
    # Crear DataFrame con una fila de ejemplo para evitar errores de tabla
    datos_ejemplo = [["", "", "", "", "", "", "", ""]]
    df_calidad = pd.DataFrame(datos_ejemplo, columns=columnas_calidad)
    
    # Escribir a Excel
    df_calidad.to_excel(writer, sheet_name="Calidad", index=False)
    
    # Obtener la hoja y aplicar formato básico
    worksheet = writer.sheets["Calidad"]
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, 9):  # A hasta H
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas
    worksheet.column_dimensions['A'].width = 10  # Codigo
    worksheet.column_dimensions['B'].width = 18  # Fecha Monitoreo
    worksheet.column_dimensions['C'].width = 12  # ID Asesor
    worksheet.column_dimensions['D'].width = 8   # VOZ
    worksheet.column_dimensions['E'].width = 8   # SMS
    worksheet.column_dimensions['F'].width = 10  # TERCERO
    worksheet.column_dimensions['G'].width = 12  # Nota Total
    worksheet.column_dimensions['H'].width = 18  # Total Monitoreos

def crear_hoja_operativo(writer, df_reporte3=None):
    """
    Crea la hoja "Operativo" con la tabla especificada e integra datos del Reporte 3
    """
    # Definir las columnas de la hoja Operativo
    columnas_operativo = [
        "CODIGO", "Tipo Jornada", "Fecha", "Cedula", "ID", "EXT", "VOIP", "Nombre", "Sede", "Ubicación",
        "Logueo", "Mora", "Asignación", "Clientes gestionados 11 am", "Capital Asignado", "Capital Recuperado",
        "PAGOS", "% Recuperado", "% Cuentas", "Total toques", "Ultimo Toque", "Llamadas Microsip",
        "Llamadas VOIP", "Total Llamadas", "Gerencia", "Team", "Meta", "Ejecución", "Ind Logueo",
        "Ind Ultimo", "Ind Ges Medio", "Ind Llamadas", "Indicador Toques", "Ind Pausa", "Total Infracciones", "Total Operativo"
    ]
    
    # Crear DataFrame con datos del Reporte 3 si existe
    if df_reporte3 is not None and not df_reporte3.empty:
        print("Integrando datos del Reporte 3 en hoja Operativo...")
        
        # Mapeo de columnas entre Reporte 3 y Operativo
        # Necesitamos ajustar esto según las columnas reales del Reporte 3
        mapeo_columnas = {}
        
        # Buscar coincidencias - PRIORIZAR EXACTAS primero
        for col_reporte3 in df_reporte3.columns:
            col_lower = col_reporte3.lower().strip()
            for col_operativo in columnas_operativo:
                col_op_lower = col_operativo.lower().strip()
                # Primero buscar coincidencias exactas
                if col_lower == col_op_lower:
                    mapeo_columnas[col_operativo] = col_reporte3
                    break
        
        # Luego buscar coincidencias parciales para columnas no mapeadas
        for col_reporte3 in df_reporte3.columns:
            col_lower = col_reporte3.lower().strip()
            for col_operativo in columnas_operativo:
                col_op_lower = col_operativo.lower().strip()
                # Solo si no se mapeó exactamente antes
                if col_operativo not in mapeo_columnas:
                    if col_lower in col_op_lower or col_op_lower in col_lower:
                        mapeo_columnas[col_operativo] = col_reporte3
                        break
        
        print(f"Mapeo de columnas encontrado: {mapeo_columnas}")
        
        # Debug: Mostrar valores específicos para VOIP
        if 'VOIP' in mapeo_columnas:
            print(f"Columna VOIP mapeada a: {mapeo_columnas['VOIP']}")
            print(f"Primeros 5 valores VOIP del Reporte 3: {df_reporte3[mapeo_columnas['VOIP']].head().tolist()}")
        else:
            print("ERROR: Columna VOIP no encontrada en el mapeo!")
            print(f"Columnas disponibles en Reporte 3: {list(df_reporte3.columns)}")
        
        # Crear DataFrame con los datos mapeados
        datos_operativo = []
        for _, fila_reporte3 in df_reporte3.iterrows():
            fila_operativo = []
            for col_operativo in columnas_operativo:
                if col_operativo in mapeo_columnas:
                    valor = fila_reporte3[mapeo_columnas[col_operativo]]
                    fila_operativo.append(valor)
                else:
                    fila_operativo.append("")  # Columna vacía si no hay mapeo
            datos_operativo.append(fila_operativo)
        
        df_operativo = pd.DataFrame(datos_operativo, columns=columnas_operativo)
        print(f"Datos integrados: {len(df_operativo)} registros")
        
        # Debug: Verificar valores VOIP después del mapeo
        if 'VOIP' in df_operativo.columns:
            print(f"Primeros 5 valores VOIP en resultado: {df_operativo['VOIP'].head().tolist()}")
        
    else:
        # Crear DataFrame con una fila de ejemplo si no hay datos del Reporte 3
        datos_ejemplo = [[""] * len(columnas_operativo)]
        df_operativo = pd.DataFrame(datos_ejemplo, columns=columnas_operativo)
        print("No hay datos del Reporte 3, creando hoja con estructura vacía")
    
    # Escribir a Excel
    df_operativo.to_excel(writer, sheet_name="Operativo", index=False)
    
    # Obtener la hoja y aplicar formato básico
    worksheet = writer.sheets["Operativo"]
    
    # AGREGAR FÓRMULAS EN LA COLUMNA CODIGO (primera columna)
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
        
        # Agregar fórmulas en la columna CODIGO para cada fila de datos
        # Formato: ID + DDMM (ej: 2020 + 0908 = 20200908)
        for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (después del header)
            # Fórmula Excel para extraer día y mes de fecha DD/MM/YYYY
            # Si fecha es 09/08/2025, resultado debe ser: ID + "0908"
            formula = f'={id_letter}{row}&TEXT(DAY({fecha_letter}{row}),"00")&TEXT(MONTH({fecha_letter}{row}),"00")'
            cell = worksheet.cell(row=row, column=1)  # Columna A (CODIGO)
            cell.value = formula
        print(f"Fórmulas agregadas en columna CODIGO: {len(df_operativo)} fórmulas con formato ID+DDMM")
        
        # AGREGAR FÓRMULAS EN LA COLUMNA TIPO JORNADA (segunda columna)
        # Encontrar la posición de la columna Tipo Jornada
        tipo_jornada_col = None
        for i, col_name in enumerate(columnas_operativo):
            if col_name == "Tipo Jornada":
                tipo_jornada_col = i + 1  # Excel columnas empiezan en 1
                break
        
        if tipo_jornada_col and fecha_col:
            tipo_jornada_letter = get_column_letter(tipo_jornada_col)
            
            # Agregar fórmulas en la columna Tipo Jornada para cada fila de datos
            # Si día es 30, 31, 1, 2, 15, 16, 17 → "Pago", sino → "Normal"
            for row in range(2, len(df_operativo) + 2):  # Empezar en fila 2 (después del header)
                formula = f'=IF(OR(DAY({fecha_letter}{row})=30,DAY({fecha_letter}{row})=31,DAY({fecha_letter}{row})=1,DAY({fecha_letter}{row})=2,DAY({fecha_letter}{row})=15,DAY({fecha_letter}{row})=16,DAY({fecha_letter}{row})=17),"Pago","Normal")'
                cell = worksheet.cell(row=row, column=tipo_jornada_col)
                cell.value = formula
            print(f"Fórmulas agregadas en columna Tipo Jornada: {len(df_operativo)} fórmulas con lógica Pago/Normal (días: 30,31,1,2,15,16,17)")
            
            # Agregar validación de datos (dropdown) en la columna Tipo Jornada
            from openpyxl.worksheet.datavalidation import DataValidation
            
            # Crear validación con lista desplegable
            dv = DataValidation(type="list", formula1='"Pago,Normal"', allow_blank=False)
            dv.error = 'Seleccione: Pago o Normal'
            dv.errorTitle = 'Valor inválido'
            dv.prompt = 'Seleccione el tipo de jornada'
            dv.promptTitle = 'Tipo Jornada'
            
            # Aplicar validación a todas las filas de datos
            range_validation = f"{tipo_jornada_letter}2:{tipo_jornada_letter}{len(df_operativo) + 1}"
            dv.add(range_validation)
            worksheet.add_data_validation(dv)
            print(f"Validación de datos agregada en columna Tipo Jornada: rango {range_validation}")
        else:
            print("ERROR: No se encontró columna Tipo Jornada para las fórmulas")
    else:
        print("ERROR: No se encontraron columnas ID y/o Fecha para las fórmulas")
    
    # Aplicar formato de encabezados
    from openpyxl.styles import Font, PatternFill
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="E74C3C", end_color="E74C3C", fill_type="solid")
    
    # Formatear encabezados
    for col in range(1, len(columnas_operativo) + 1):
        cell = worksheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
    
    # Ajustar ancho de columnas - usando get_column_letter para manejar columnas más allá de Z
    from openpyxl.utils import get_column_letter
    anchos_operativo = [10, 12, 10, 12, 8, 8, 8, 20, 10, 12, 10, 8, 12, 18, 15, 15, 10, 12, 12, 12, 15, 15, 15, 15, 12, 10, 8, 12, 12, 12, 12, 12, 15, 12, 18, 15]
    for i, ancho in enumerate(anchos_operativo, 1):
        worksheet.column_dimensions[get_column_letter(i)].width = ancho

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
    
    # Obtener la hoja y aplicar formato básico
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
    
    # Obtener la hoja y aplicar formato básico
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

def crear_hoja_consolidado(writer):
    """
    Crea la hoja "Consolidado" con la tabla especificada
    """
    # Crear DataFrame con las columnas especificadas y una fila de ejemplo
    columnas_consolidado = [
        "Codigo_Asis", "CODIGO", "Tipo Jornada", "Fecha", "Cedula", "ID", "Nombre", "Sede", "Ubicación",
        "Asistencia", "Mora", "Monitoreos", "Nota Calidad", "Ejecución", "# Infracciones", 
        "% Operativo", "Team", "Gerente"
    ]
    
    # Crear DataFrame con una fila de ejemplo
    datos_ejemplo = [[""] * len(columnas_consolidado)]
    df_consolidado = pd.DataFrame(datos_ejemplo, columns=columnas_consolidado)
    
    # Escribir a Excel
    df_consolidado.to_excel(writer, sheet_name="Consolidado", index=False)
    
    # Obtener la hoja y aplicar formato básico
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
    worksheet.column_dimensions['I'].width = 12  # Ubicación
    worksheet.column_dimensions['J'].width = 12  # Asistencia
    worksheet.column_dimensions['K'].width = 8   # Mora
    worksheet.column_dimensions['L'].width = 12  # Monitoreos
    worksheet.column_dimensions['M'].width = 15  # Nota Calidad
    worksheet.column_dimensions['N'].width = 12  # Ejecución
    worksheet.column_dimensions['O'].width = 15  # # Infracciones
    worksheet.column_dimensions['P'].width = 15  # % Operativo
    worksheet.column_dimensions['Q'].width = 10  # Team
    worksheet.column_dimensions['R'].width = 12  # Gerente
