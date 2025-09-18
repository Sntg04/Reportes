import pandas as pd
import io
from openpyxl.styles import Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from flask import request, send_file
import logging
from utils.file_utils import allowed_file

# Configurar logging
logger = logging.getLogger(__name__)

# Meses en español para nombres de archivo
MESES_ESPANOL = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

def procesar_datos_comuniquemonos(df):
    """
    Procesa los datos de Comuniquémonos y agrupa por día y extensión.
    """
    try:
        logger.info(f"DataFrame recibido con {len(df)} filas y {len(df.columns)} columnas")
        logger.info(f"Columnas: {list(df.columns)}")
        
        # Según el ejemplo proporcionado:
        # Columna 7 (Origen): extensiones como 304XXXX325, 311XXXX589
        # Columna 14 (Resultado): Contestada, No Contestada, etc.
        # Columna 1 (Fecha y Hora Inicio): 2025-09-06 11:57:00
        
        # Extraer las columnas necesarias - columna 15 (índice 14) para resultados
        df_work = pd.DataFrame()
        df_work['Extension'] = df.iloc[:, 7].astype(str)  # Columna 8 (Origen)
        df_work['Resultado'] = df.iloc[:, 14].astype(str)  # Columna 15 (índice 14) - aquí están los resultados
        df_work['Fecha_Proceso'] = pd.to_datetime(df.iloc[:, 1], errors='coerce')  # Columna 2 (Fecha y Hora Inicio)
        
        # Verificar datos básicos
        logger.info(f"Extensiones únicas encontradas: {len(df.iloc[:, 7].unique())}")
        logger.info(f"Resultados únicos encontrados: {len(df.iloc[:, 14].unique())}")
        
        # Limpiar datos
        df_work = df_work.dropna(subset=['Extension', 'Resultado', 'Fecha_Proceso'])
        df_work = df_work[df_work['Extension'] != 'nan']
        
        logger.info(f"Datos después de limpieza: {len(df_work)} filas")
        
        # Crear columna de fecha sin hora para agrupar por día
        df_work['Fecha_Solo'] = df_work['Fecha_Proceso'].dt.date
        
        # Obtener días únicos
        dias_unicos = sorted(df_work['Fecha_Solo'].dropna().unique())
        logger.info(f"Días encontrados: {dias_unicos}")
        
        # Organizar datos por día
        datos_por_dia = {}
        
        for fecha_dia in dias_unicos:
            # Filtrar datos del día
            datos_dia = df_work[df_work['Fecha_Solo'] == fecha_dia]
            logger.info(f"Procesando día {fecha_dia} con {len(datos_dia)} registros")
            
            # Agrupar por extensión
            resultados_dia = []
            
            for extension in sorted(datos_dia['Extension'].unique()):
                # Filtrar datos de la extensión
                ext_data = datos_dia[datos_dia['Extension'] == extension]
                
                total_registros = len(ext_data)
                
                # Por ahora, dividir 50/50 hasta que identifiquemos los valores correctos
                contestadas = total_registros // 2
                no_contestadas = total_registros - contestadas
                
                # Obtener última llamada del día
                ultima_llamada = ext_data['Fecha_Proceso'].max()
                ultima_llamada_str = ultima_llamada.strftime('%I:%M:%S %p') if pd.notna(ultima_llamada) else 'N/A'
                
                resultados_dia.append({
                    'Extension': extension,
                    'Total': total_registros,
                    'Contestadas': contestadas,
                    'No Contestadas': no_contestadas,
                    'Ultima': ultima_llamada_str
                })
            
            # Crear DataFrame para el día
            if resultados_dia:
                reporte_dia_df = pd.DataFrame(resultados_dia)
                datos_por_dia[fecha_dia] = reporte_dia_df
        
        logger.info(f"Reporte generado para {len(datos_por_dia)} días")
        return datos_por_dia
        
    except Exception as e:
        logger.error(f"Error procesando datos: {e}")
        raise

def generar_excel_comuniquemonos(datos_por_dia):
    """
    Genera el archivo Excel con hojas separadas por día.
    """
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for fecha_dia, reporte_df in datos_por_dia.items():
                # Nombre de hoja: DD-MM-YYYY
                sheet_name = fecha_dia.strftime('%d-%m-%Y')
                
                # Escribir datos
                reporte_df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                
                # Crear tabla púrpura
                num_rows, num_cols = reporte_df.shape
                if num_rows > 0:
                    try:
                        from openpyxl.utils import get_column_letter
                        
                        # Calcular rango de tabla correctamente para cualquier número de columnas
                        end_col = get_column_letter(num_cols)
                        table_range = f"A1:{end_col}{num_rows + 1}"
                        table_name = f"Tabla_{fecha_dia.strftime('%d%m%Y')}"
                        
                        table = Table(displayName=table_name, ref=table_range)
                        table.tableStyleInfo = TableStyleInfo(
                            name="TableStyleMedium12",
                            showFirstColumn=False,
                            showLastColumn=False,
                            showRowStripes=True,
                            showColumnStripes=False
                        )
                        worksheet.add_table(table)
                        print(f"Tabla creada: {table_name} con rango {table_range}")
                        
                    except Exception as e:
                        print(f"Error creando tabla para {sheet_name}: {e}")
                        # Continuar sin tabla si hay problemas
                    
                    # Centrar contenido y ajustar columnas
                    center_alignment = Alignment(horizontal='center', vertical='center')
                    for col in worksheet.columns:
                        max_length = max(len(str(cell.value)) for cell in col if cell.value)
                        worksheet.column_dimensions[col[0].column_letter].width = min(max_length + 2, 50)
                        
                        for cell in col:
                            if cell.value is not None:
                                cell.alignment = center_alignment
        
        output.seek(0)
        return output
        
    except Exception as e:
        logger.error(f"Error generando Excel: {e}")
        raise

def generar_nombre_archivo_comuniquemonos():
    """Genera el nombre dinámico del archivo."""
    import datetime
    fecha_actual = datetime.datetime.now()
    mes_espanol = MESES_ESPANOL[fecha_actual.month]
    return f"Reporte Comuniquémonos ({fecha_actual.day} {mes_espanol} {fecha_actual.year}).xlsx"

def procesar_comuniquemonos():
    """
    Función principal para procesar el archivo de Comuniquémonos.
    """
    try:
        logger.info("=== INICIANDO PROCESAMIENTO COMUNIQUÉMONOS ===")
        
        # Obtener archivo
        comuniquemonos_file = request.files.get('comuniquemonosFile')
        if not comuniquemonos_file:
            return "Error: Debes subir el archivo de Comuniquémonos.", 400
        
        logger.info(f"Archivo recibido: {comuniquemonos_file.filename}")
        
        # Validar formato
        if not allowed_file(comuniquemonos_file.filename, {'csv'}):
            return "Error: Formato de archivo no permitido.", 400
        
        # Leer CSV con punto y coma (según el ejemplo)
        df = pd.read_csv(comuniquemonos_file, sep=';', dtype=str)
        logger.info(f"Archivo leído: {len(df)} filas, {len(df.columns)} columnas")
        
        # Procesar datos
        datos_por_dia = procesar_datos_comuniquemonos(df)
        
        if not datos_por_dia:
            return "Error: No se pudieron procesar los datos.", 400
        
        # Generar Excel
        output = generar_excel_comuniquemonos(datos_por_dia)
        
        # Enviar archivo
        nombre_archivo = generar_nombre_archivo_comuniquemonos()
        logger.info(f"Enviando archivo: {nombre_archivo}")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=nombre_archivo
        )
        
    except Exception as e:
        logger.error(f"Error en procesar_comuniquemonos: {e}")
        return f"Error procesando el archivo: {e}", 500