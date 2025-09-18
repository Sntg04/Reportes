"""
Módulo para el procesamiento del 3_Reporte_Reporteria
"""
import io
import json
import os
import re
import shutil
from datetime import datetime
from typing import Any, Dict, List, Optional, Union

import pandas as pd
from flask import request, send_file, jsonify
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from utils.file_utils import allowed_file

# ============================================================================
# CONSTANTES
# ============================================================================
BASE_ASESORES_FILE = 'base_asesores.json'
EXCEL_EPOCH_THRESHOLD = 40000
DEFAULT_SHEET_NAME = 'Información'
TABLE_STYLE = "TableStyleMedium12"
MAX_COLUMN_WIDTH = 50
FORMATO_HORA = 'h:mm:ss AM/PM'  # Formato de hora con segundos y AM/PM

# Columnas esperadas en archivos
REQUIRED_ASESOR_COLUMNS = [
    'Fecha Ingreso', 'Fecha', 'Cedula', 'ID', 'EXT', 
    'VOIP', 'Nombre', 'Sede', 'Ubicación'
]

ADMIN_COLUMNS = {
    'ID': 'ID',
    'NOMBRE': 'Nombre',
    'LOGUEO': 'Logueo',
    'CARTERA': 'CARTERA',
    'ASIGNACION': 'ASIGNACION',
    'TOCADAS_11AM': 'TOCADAS 11 AM',
    'ASIGNADO': 'ASIGNADO',
    'RECUPERADO': 'RECUPERADO',
    'PAGOS': 'PAGOS',
    'PORCENTAJE_RECUPERADO': '% RECUPERADO',
    'PORCENTAJE_CUENTAS': '% CUENTAS',
    'TOQUES': 'TOQUES',
    'ULTIMO_TOQUE': 'ULTIMO TOQUE',
    'GERENTE': 'Gerente',
    'TEAM_LEADER': 'Team Leader',
    'UBICACION': 'Ubicación'
}

LLAMADAS_COLUMNS = {
    'EXTENSION': 'Número Extensión',  # Exactamente como lo genera el Reporte 1
    'TOTAL_LLAMADAS': 'Total Llamadas'
}

# Mapeo de correcciones de valores de mora
MORA_CORRECTIONS = {
    'M0-1 PP': 'M0-1-PP',
    'M1-1A FRS': 'M1-1A-FRS',
    'M1-1A BT': 'M1-1A-BT',
    'M1-1A PN': 'M1-1A-PN'
}

# Diccionario para nombres de meses en español
MESES_ESPANOL = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 
    5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
    9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

# ============================================================================
# FUNCIONES DE UTILIDAD
# ============================================================================
def formatear_nombre(nombre: str) -> str:
    """
    Formatea nombres poniendo la primera letra de cada palabra en mayúscula
    y el resto en minúscula (formato título)
    
    Args:
        nombre (str): Nombre a formatear
        
    Returns:
        str: Nombre formateado
        
    Example:
        formatear_nombre("JUAN CARLOS pérez") -> "Juan Carlos Pérez"
    """
    if not nombre or pd.isna(nombre):
        return ''
    
    return str(nombre).title()

def generar_nombre_archivo_reporteria(hojas_procesadas: int, admin_sheets: Dict[str, Any]) -> str:
    """
    Genera nombre dinámico para el archivo Excel del reporte de reportería
    
    Args:
        hojas_procesadas (int): Número de hojas procesadas exitosamente
        admin_sheets (Dict[str, Any]): Diccionario con las hojas del archivo admin
        
    Returns:
        str: Nombre del archivo generado con formato de fecha español
        
    Example:
        generar_nombre_archivo_reporteria(3, sheets) -> "Reporte Reporteria (1-15 Septiembre 2025).xlsx"
    """
    try:
        # Extraer fechas de los nombres de las hojas
        fechas = []
        for sheet_name in admin_sheets.keys():
            try:
                # Convertir el formato de fecha de la hoja
                fecha_formateada = convertir_fecha_formato(sheet_name)
                if fecha_formateada and '/' in fecha_formateada:
                    # Convertir DD/MM/YYYY a datetime
                    fecha_obj = datetime.strptime(fecha_formateada, '%d/%m/%Y')
                    fechas.append(fecha_obj)
            except (ValueError, AttributeError):
                continue
        
        if not fechas:
            # Si no hay fechas válidas, usar fecha actual
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            return f"Reporte Reporteria ({timestamp}).xlsx"
        
        # Ordenar fechas y obtener rango
        fechas.sort()
        fecha_min = fechas[0]
        fecha_max = fechas[-1]
        
        if fecha_min.date() == fecha_max.date():
            # Un solo día
            mes_nombre = MESES_ESPANOL[fecha_min.month]
            return f"Reporte Reporteria ({fecha_min.day} {mes_nombre} {fecha_min.year}).xlsx"
        elif fecha_min.month == fecha_max.month and fecha_min.year == fecha_max.year:
            # Mismo mes y año
            mes_nombre = MESES_ESPANOL[fecha_min.month]
            return f"Reporte Reporteria ({fecha_min.day}-{fecha_max.day} {mes_nombre} {fecha_min.year}).xlsx"
        else:
            # Diferentes meses o años
            mes_min = MESES_ESPANOL[fecha_min.month]
            mes_max = MESES_ESPANOL[fecha_max.month]
            if fecha_min.year == fecha_max.year:
                return f"Reporte Reporteria ({fecha_min.day} {mes_min} - {fecha_max.day} {mes_max} {fecha_min.year}).xlsx"
            else:
                return f"Reporte Reporteria ({fecha_min.day} {mes_min} {fecha_min.year} - {fecha_max.day} {mes_max} {fecha_max.year}).xlsx"
                
    except Exception as e:
        # En caso de error, usar timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"Reporte Reporteria (Error_{timestamp}).xlsx"

def safe_get_value(row: Optional[Dict[str, Any]], key: str, default: Any = '') -> Any:
    """Obtiene un valor de una fila de manera segura"""
    if not isinstance(row, dict) or not row:
        return default
    return row.get(key, default)

def safe_float_conversion(value: Any, default: float = 0) -> float:
    """Convierte un valor a float de manera segura"""
    if pd.isna(value) or value == '' or value is None:
        return default
    try:
        return float(str(value).replace('$', '').replace(',', ''))
    except (ValueError, TypeError):
        return default

def limpiar_fecha(fecha_str: str) -> str:
    """Convierte fechas en formato ISO a DD/MM/YYYY"""
    if not fecha_str or pd.isna(fecha_str):
        return ''
    
    fecha_str = str(fecha_str).strip()
    
    # Si ya está en formato DD/MM/YYYY, devolverlo
    if len(fecha_str) == 10 and fecha_str.count('/') == 2:
        parts = fecha_str.split('/')
        if len(parts) == 3 and len(parts[0]) <= 2 and len(parts[1]) <= 2 and len(parts[2]) == 4:
            return fecha_str
    
    # Probar varios formatos de fecha posibles
    formatos = [
        '%Y-%m-%d',      # 2025-08-04
        '%Y-%m-%d %H:%M:%S',  # 2025-08-04 12:00:00 (datetime completo)
        '%d/%m/%Y',      # 04/08/2025
        '%d-%m-%Y',      # 04-08-2025
        '%Y/%m/%d',      # 2025/08/04
        '%m/%d/%Y',      # 08/04/2025
        '%d.%m.%Y',      # 04.08.2025
    ]
    
    for formato in formatos:
        try:
            fecha_obj = datetime.strptime(fecha_str, formato)
            return fecha_obj.strftime('%d/%m/%Y')
        except ValueError:
            continue
    
    # Intentar con pandas como última opción
    try:
        fecha_obj = pd.to_datetime(fecha_str, errors='raise')
        return fecha_obj.strftime('%d/%m/%Y')
    except:
        pass
    
    # Si no se puede parsear, devolver el valor original
    print(f"ADVERTENCIA: No se pudo convertir fecha: '{fecha_str}'")
    return fecha_str

def convertir_fecha_formato(fecha_str: str) -> str:
    """Convierte fecha de formato largo español (ej: '8 de Septiembre de 2025') a DD/MM/YYYY"""
    if not fecha_str or pd.isna(fecha_str):
        return ''
    
    fecha_str = str(fecha_str).strip()
    
    # Mapeo de meses en español
    meses = {
        'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
        'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
        'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
    }
    
    try:
        # Buscar patrón: "día de mes de año"
        patron = r'(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})'
        match = re.search(patron, fecha_str.lower())
        
        if match:
            dia = match.group(1).zfill(2)  # Agregar cero al inicio si es necesario
            mes_texto = match.group(2).lower()
            año = match.group(3)
            
            mes_numero = meses.get(mes_texto, '01')
            return f"{dia}/{mes_numero}/{año}"
        
        # Intentar formato ISO: YYYY-MM-DD
        if re.match(r'\d{4}-\d{2}-\d{2}', fecha_str):
            fecha_obj = datetime.strptime(fecha_str, '%Y-%m-%d')
            return fecha_obj.strftime('%d/%m/%Y')
            
        # Intentar formato: DD-MM-YYYY
        if re.match(r'\d{2}-\d{2}-\d{4}', fecha_str):
            fecha_obj = datetime.strptime(fecha_str, '%d-%m-%Y')
            return fecha_obj.strftime('%d/%m/%Y')
            
    except Exception:
        pass
    
    # Extraer solo la parte de fecha si hay espacios
    if ' ' in fecha_str:
        fecha_str = fecha_str.split(' ')[0]
    
    # Eliminar decimales de Excel
    fecha_str = fecha_str.replace('.0', '')
    
    # Convertir timestamp de Excel si es necesario
    if fecha_str.isdigit() and len(fecha_str) > 6:
        try:
            fecha_num = float(fecha_str)
            if fecha_num > EXCEL_EPOCH_THRESHOLD:  # Fechas después del año 2009
                from datetime import timedelta
                excel_epoch = datetime(1899, 12, 30)
                fecha_convertida = excel_epoch + timedelta(days=fecha_num)
                return fecha_convertida.strftime('%d/%m/%Y')
        except (ValueError, TypeError):
            pass
    
    return fecha_str

def formatear_valor_monetario(valor: Any) -> str:
    """Formatea valores monetarios agregando el signo peso ($) usando punto como separador de miles"""
    valor_num = safe_float_conversion(valor)
    
    if valor_num == 0:
        return '$0'
    
    # Formatear manualmente con puntos como separador de miles
    valor_str = f"{valor_num:.0f}"
    if len(valor_str) <= 3:
        return f"${valor_str}"
    
    # Agregar puntos cada 3 dígitos desde la derecha
    valor_formateado = ""
    for i, digit in enumerate(reversed(valor_str)):
        if i > 0 and i % 3 == 0:
            valor_formateado = "." + valor_formateado
        valor_formateado = digit + valor_formateado
    
    return f"${valor_formateado}"

def convertir_porcentaje_a_decimal(valor_porcentaje: Any) -> float:
    """
    Convierte un valor de porcentaje (string con %) a decimal para Excel
    
    Args:
        valor_porcentaje: String como "25%" o "0%" o valor numérico
        
    Returns:
        float: Valor decimal (25% -> 0.25)
    """
    if pd.isna(valor_porcentaje) or valor_porcentaje == '' or valor_porcentaje is None:
        return 0.0
    
    try:
        # Si ya es un número, asumir que está en formato decimal
        if isinstance(valor_porcentaje, (int, float)):
            return float(valor_porcentaje)
        
        # Si es string, convertir
        valor_str = str(valor_porcentaje).strip()
        
        # Si tiene símbolo %, quitarlo y convertir a decimal
        if '%' in valor_str:
            numero = valor_str.replace('%', '').strip()
            return float(numero) / 100.0
        
        # Si no tiene %, asumir que ya está en decimal
        return float(valor_str)
        
    except (ValueError, TypeError):
        return 0.0

def normalizar_valor_mora(valor_mora: Any) -> Any:
    """Normaliza los valores de mora para usar guiones consistentes"""
    if not valor_mora or pd.isna(valor_mora):
        return valor_mora
    
    valor_str = str(valor_mora).strip()
    
    # Aplicar corrección si existe mapeo directo
    if valor_str in MORA_CORRECTIONS:
        return MORA_CORRECTIONS[valor_str]
    
    # Si no hay mapeo directo, aplicar regla general: reemplazar espacios con guiones
    return valor_str.replace(' ', '-')

def convertir_hora_formato(hora_str: Any) -> str:
    """Convierte horas al formato de 12 horas con segundos y AM/PM"""
    if not hora_str or pd.isna(hora_str) or hora_str == '' or hora_str is None:
        return ''
    
    hora_str = str(hora_str).strip()
    
    # Si ya está en formato AM/PM, devolverlo tal como está
    if 'AM' in hora_str.upper() or 'PM' in hora_str.upper():
        return hora_str
    
    try:
        # Intentar parsear diferentes formatos de hora
        formatos_hora = [
            '%H:%M:%S',     # 14:30:00
            '%H:%M',        # 14:30
            '%I:%M:%S %p',  # 2:30:00 PM
            '%I:%M %p',     # 2:30 PM
        ]
        
        for formato in formatos_hora:
            try:
                hora_obj = datetime.strptime(hora_str, formato)
                return hora_obj.strftime('%I:%M:%S %p').lstrip('0')  # Formato 12 horas con segundos y AM/PM
            except ValueError:
                continue
        
        # Si ningún formato funciona, intentar con pandas
        hora_obj = pd.to_datetime(hora_str, errors='coerce')
        if not pd.isna(hora_obj):
            return hora_obj.strftime('%I:%M:%S %p').lstrip('0')
            
        return hora_str
        
    except Exception:
        return hora_str

def load_base_asesores() -> List[Dict[str, Any]]:
    """
    Carga la base de datos de asesores desde archivo JSON
    
    Returns:
        List[Dict[str, Any]]: Lista de diccionarios con datos de asesores,
                             lista vacía si el archivo no existe o tiene errores
                             
    Note:
        - Maneja automáticamente FileNotFoundError y JSONDecodeError
        - Retorna lista vacía en caso de error para evitar interrupciones
    """
    try:
        with open(BASE_ASESORES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return []
    except json.JSONDecodeError:
        return []

def create_data_map(data: Optional[pd.DataFrame], key_column: str) -> Dict[str, Dict[str, Any]]:
    """
    Crea un diccionario de mapeo usando una columna como clave única
    
    Args:
        data (Optional[pd.DataFrame]): DataFrame fuente con los datos
        key_column (str): Nombre de la columna que servirá como clave
        
    Returns:
        Dict[str, Dict[str, Any]]: Diccionario donde las claves son los valores
                                  de key_column y los valores son filas completas
                                  
    Note:
        - Convierte las claves a string para evitar problemas de tipo
        - Retorna diccionario vacío si la columna no existe o hay errores
        - Útil para búsquedas rápidas por ID o extensión
    """
    if data is None or len(data) == 0:
        return {}
    
    try:
        # Verificar si la columna clave existe
        if key_column not in data.columns:
            print(f"❌ Columna '{key_column}' no encontrada")
            return {}
        
        # Crear el mapa
        data_map = {
            str(row[key_column]): row.to_dict() 
            for _, row in data.iterrows() 
            if pd.notna(row.get(key_column))
        }
        
        print(f"✅ Mapa creado: {len(data_map)} registros para '{key_column}'")
        return data_map
        
    except Exception as e:
        print(f"❌ Error creando mapa: {e}")
        return {}

def apply_excel_formatting(worksheet: Any, dataframe: pd.DataFrame, sheet_name: str) -> None:
    """
    Aplica formato profesional de tabla y estilos a una hoja de Excel
    
    Args:
        worksheet (Any): Hoja de Excel de openpyxl donde aplicar el formato
        dataframe (pd.DataFrame): DataFrame fuente para determinar dimensiones
        sheet_name (str): Nombre de la hoja para generar nombre único de tabla
        
    Returns:
        None: Modifica la hoja directamente
        
    Note:
        - Crea tabla con estilo TABLE_STYLE (púrpura con rayas)
        - Ajusta ancho automático de columnas (máximo MAX_COLUMN_WIDTH)
        - Aplica formato de porcentaje a columnas específicas
        - Centra el contenido de todas las celdas
    """
    num_rows, num_cols = dataframe.shape
    
    if num_rows == 0 or num_cols == 0:
        return
    
    try:
        # Crear nombre de tabla válido (solo letras, números y underscore)
        clean_name = re.sub(r'[^a-zA-Z0-9_]', '_', sheet_name)
        table_name = f"Tabla_{clean_name}"
        
        # Calcular rango de tabla correctamente para cualquier número de columnas
        end_col = get_column_letter(num_cols)
        table_range = f"A1:{end_col}{num_rows + 1}"
        
        # Crear tabla púrpura
        table = Table(displayName=table_name, ref=table_range)
        table.tableStyleInfo = TableStyleInfo(
            name=TABLE_STYLE, 
            showFirstColumn=False, 
            showLastColumn=False,
            showRowStripes=True, 
            showColumnStripes=False
        )
        worksheet.add_table(table)
        print(f"✅ Tabla creada: {table_name}")
        
    except Exception as e:
        print(f"❌ Error creando tabla: {e}")
        # Continuar sin tabla si hay problemas
    
    try:
        # Ajustar ancho de columnas
        for col_obj in worksheet.columns:
            column_letter = col_obj[0].column_letter
            max_length = max(
                len(str(cell.value)) for cell in col_obj 
                if cell.value is not None
            )
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, MAX_COLUMN_WIDTH)
        
        # Aplicar formato de porcentaje a columnas específicas
        column_letters = {}
        for col_num, col_name in enumerate(dataframe.columns, 1):
            column_letters[col_name] = get_column_letter(col_num)
        
        # Formato de porcentaje para las columnas de porcentaje
        if '% Recuperado' in column_letters and '% Cuentas' in column_letters:
            percent_format = '0.00%'
            for row_num in range(2, worksheet.max_row + 1):
                # Columna % Recuperado
                percent_recup_cell = f"{column_letters['% Recuperado']}{row_num}"
                worksheet[percent_recup_cell].number_format = percent_format
                
                # Columna % Cuentas
                percent_cuentas_cell = f"{column_letters['% Cuentas']}{row_num}"
                worksheet[percent_cuentas_cell].number_format = percent_format
        
        # Formato de hora para las columnas de tiempo
        time_columns = ['Logueo', 'Ultimo Toque']
        for time_col in time_columns:
            if time_col in column_letters:
                for row_num in range(2, worksheet.max_row + 1):
                    time_cell = f"{column_letters[time_col]}{row_num}"
                    worksheet[time_cell].number_format = FORMATO_HORA
        
        # Centrar contenido en todas las celdas
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
    except Exception as e:
        print(f"Error aplicando formato a {sheet_name}: {e}")
        # Continuar sin formato si hay problemas

# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================
def procesar_reporteria_cobranza():
    """
    Procesa los archivos de reportería de cobranza y genera el 3_Reporte_Reporteria
    """
    
    # Validar archivos de entrada
    admin_file = request.files.get('reporteAdminFile')
    llamadas_file = request.files.get('reporteLlamadasFile')

    if not admin_file or not llamadas_file:
        return "Error: Debes subir ambos archivos de reporte (Admin y Llamadas).", 400
    
    if not (allowed_file(admin_file.filename, {'xlsx'}) and 
            allowed_file(llamadas_file.filename, {'xlsx'})):
        return "Error: Formato de archivo no permitido (solo .xlsx).", 400
    
    try:
        # Cargar datos de Excel
        admin_sheets = pd.read_excel(admin_file, sheet_name=None, dtype=str)
        llamadas_sheets = pd.read_excel(llamadas_file, sheet_name=None, dtype=str)
        
        print(f"✅ Archivos cargados - Admin: {len(admin_sheets)} hojas, Llamadas: {len(llamadas_sheets)} hojas")
        
        # Cargar base de asesores
        base_asesores = load_base_asesores()
        if not base_asesores:
            return "Error: No hay asesores en la base de datos.", 400
        
        output = io.BytesIO()
        hojas_procesadas = 0
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Procesar todas las hojas disponibles
            for sheet_name in admin_sheets.keys():
                try:
                    admin_data = admin_sheets.get(sheet_name)
                    llamadas_data = llamadas_sheets.get(sheet_name)
                    
                    # Crear mapas de datos
                    admin_map = create_data_map(admin_data, ADMIN_COLUMNS['ID'])
                    
                    # Debug: verificar datos de admin
                    if admin_map:
                        print(f"   🗂️  Mapa de admin creado: {len(admin_map)} registros")
                        sample_id = list(admin_map.keys())[0]
                        sample_admin = admin_map[sample_id]
                        print(f"   📋 Ejemplo ID {sample_id} - columnas disponibles: {list(sample_admin.keys())}")
                        print(f"   💼 Valores ejemplo: CARTERA={sample_admin.get('CARTERA')}, ASIGNACION={sample_admin.get('ASIGNACION')}, RECUPERADO={sample_admin.get('RECUPERADO')}")
                    else:
                        print(f"   ❌ No se pudo crear mapa de admin para hoja '{sheet_name}'")
                    
                    # Verificar si existe el archivo de llamadas para esta hoja
                    if llamadas_data is not None and not llamadas_data.empty:
                        llamadas_map = create_data_map(llamadas_data, LLAMADAS_COLUMNS['EXTENSION'])
                        print(f"✅ Datos de llamadas cargados para hoja '{sheet_name}': {len(llamadas_data)} registros")
                        
                        # Debug: mostrar información de mapeo de llamadas
                        if llamadas_map:
                            print(f"   📞 Mapa de llamadas creado con {len(llamadas_map)} extensiones")
                            sample_extensions = list(llamadas_map.keys())[:5]
                            print(f"   📝 Primeras 5 extensiones mapeadas: {sample_extensions}")
                            
                            # Verificar valores de llamadas en el mapa
                            for ext in sample_extensions:
                                llamadas_value = safe_get_value(llamadas_map[ext], LLAMADAS_COLUMNS['TOTAL_LLAMADAS'], 0)
                                print(f"   🔢 Ext {ext}: {llamadas_value} llamadas")
                        else:
                            print(f"   ⚠️  No se pudo crear mapa de llamadas para hoja '{sheet_name}'")
                    else:
                        llamadas_map = {}
                        print(f"❌ No hay datos de llamadas para hoja '{sheet_name}' - todas las llamadas serán 0")
                    
                    # Generar filas del reporte
                    report_rows = generate_report_rows(base_asesores, admin_map, llamadas_map, sheet_name)
                    
                    if report_rows:
                        # Crear nombre de hoja válido para Excel (reemplazar / con -)
                        excel_sheet_name = convertir_fecha_formato(sheet_name).replace('/', '-')
                        
                        # Crear DataFrame y escribir a Excel
                        df_report = pd.DataFrame(report_rows)
                        
                        # Eliminar filas que no tienen logueo (valores vacíos, nulos o 'N/A')
                        df_report = df_report[
                            (df_report['Logueo'].notna()) & 
                            (df_report['Logueo'] != '') & 
                            (df_report['Logueo'] != 'N/A')
                        ]
                        
                        df_report.to_excel(writer, sheet_name=excel_sheet_name, index=False)
                        
                        # Aplicar formato
                        worksheet = writer.sheets[excel_sheet_name]
                        apply_excel_formatting(worksheet, df_report, excel_sheet_name)
                        
                        hojas_procesadas += 1
                    
                except Exception:
                    continue

            # Si no se procesó ninguna hoja, crear una hoja por defecto
            if hojas_procesadas == 0:
                create_default_sheet(writer, admin_sheets.keys())

        output.seek(0)
        
        # Guardar archivo temporalmente para continuar al paso 4
        temp_filename = f"temp_reporte3_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        temp_filepath = os.path.join('temp_files', temp_filename)
        
        # Crear directorio temp_files si no existe
        os.makedirs('temp_files', exist_ok=True)
        
        # Guardar archivo temporal
        with open(temp_filepath, 'wb') as f:
            f.write(output.getvalue())
        
        # Generar nombre dinámico del archivo basado en las fechas procesadas
        nombre_archivo_dinamico = generar_nombre_archivo_reporteria(hojas_procesadas, admin_sheets)
        
        # Guardar el nombre dinámico en un archivo de metadata
        metadata_filename = temp_filename.replace('.xlsx', '_metadata.json')
        metadata_filepath = os.path.join('temp_files', metadata_filename)
        metadata = {
            'download_name': nombre_archivo_dinamico,
            'hojas_procesadas': hojas_procesadas,
            'created_at': datetime.now().isoformat()
        }
        with open(metadata_filepath, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
        
        # Retornar JSON con opciones
        return jsonify({
            'success': True,
            'message': 'Reporte 3 procesado exitosamente',
            'temp_file': temp_filename,
            'download_name': nombre_archivo_dinamico,
            'hojas_procesadas': hojas_procesadas
        })

    except pd.errors.EmptyDataError:
        return "Error: Uno de los archivos está vacío o no contiene datos válidos.", 400
    except pd.errors.ParserError:
        return "Error: No se pudo leer el archivo Excel. Verifica el formato.", 400
    except Exception as e:
        return f"Ocurrió un error procesando los archivos: {e}", 500

def generate_report_rows(
    base_asesores: List[Dict[str, Any]], 
    admin_map: Dict[str, Dict[str, Any]], 
    llamadas_map: Dict[str, Dict[str, Any]], 
    sheet_name: str
) -> List[Dict[str, Any]]:
    """Genera las filas del reporte para una hoja específica"""
    report_rows = []
    
    for asesor in base_asesores:
        asesor_id = str(asesor.get('ID', ''))
        asesor_ext = str(asesor.get('EXT', ''))
        
        admin_row = admin_map.get(asesor_id, {})
        llamadas_row = llamadas_map.get(asesor_ext, {})
        
        # Usar la columna correcta para el total de llamadas
        llamadas_microsip = safe_get_value(llamadas_row, LLAMADAS_COLUMNS['TOTAL_LLAMADAS'], 0)
        llamadas_microsip = safe_float_conversion(llamadas_microsip)
        
        new_row = {
            'Fecha Ingreso': limpiar_fecha(asesor.get('Fecha Ingreso', '')), 
            'Fecha': convertir_fecha_formato(sheet_name), 
            'Cedula': asesor.get('Cedula', ''),
            'ID': asesor_id, 
            'EXT': asesor_ext, 
            'VOIP': asesor.get('VOIP', ''),
            'Nombre': formatear_nombre(asesor.get('Nombre', '')), 
            'Sede': asesor.get('Sede', ''), 
            'Ubicación': asesor.get('Ubicación', ''),
            'Logueo': convertir_hora_formato(safe_get_value(admin_row, ADMIN_COLUMNS['LOGUEO'])), 
            'Mora': normalizar_valor_mora(safe_get_value(admin_row, ADMIN_COLUMNS['CARTERA'])),
            'Asignación': safe_get_value(admin_row, ADMIN_COLUMNS['ASIGNACION'], 0), 
            'Clientes gestionados 11 am': safe_get_value(admin_row, ADMIN_COLUMNS['TOCADAS_11AM'], 0),
            'Capital Asignado': formatear_valor_monetario(safe_get_value(admin_row, ADMIN_COLUMNS['ASIGNADO'], 0)), 
            'Capital Recuperado': formatear_valor_monetario(safe_get_value(admin_row, ADMIN_COLUMNS['RECUPERADO'], 0)),
            'PAGOS': safe_get_value(admin_row, ADMIN_COLUMNS['PAGOS'], 0), 
            '% Recuperado': convertir_porcentaje_a_decimal(safe_get_value(admin_row, ADMIN_COLUMNS['PORCENTAJE_RECUPERADO'], '0%')),
            '% Cuentas': convertir_porcentaje_a_decimal(safe_get_value(admin_row, ADMIN_COLUMNS['PORCENTAJE_CUENTAS'], '0%')), 
            'Total toques': safe_get_value(admin_row, ADMIN_COLUMNS['TOQUES'], 0),
            'Ultimo Toque': convertir_hora_formato(safe_get_value(admin_row, ADMIN_COLUMNS['ULTIMO_TOQUE'])), 
            'Llamadas Microsip': llamadas_microsip,
            'Llamadas VOIP': 0, 
            'Total Llamadas': llamadas_microsip,
            'Gerencia': formatear_nombre(safe_get_value(admin_row, ADMIN_COLUMNS['GERENTE'])), 
            'Team': formatear_nombre(safe_get_value(admin_row, ADMIN_COLUMNS['TEAM_LEADER']))
        }
        report_rows.append(new_row)
    
    return report_rows

def create_default_sheet(writer: Any, available_sheets: Dict[str, Any]) -> None:
    """
    Crea una hoja por defecto cuando no se procesaron hojas válidas
    
    Args:
        writer (Any): Objeto ExcelWriter de pandas
        available_sheets (Dict[str, Any]): Diccionario con hojas disponibles
        
    Returns:
        None: Modifica el writer directamente
    """
    default_data = pd.DataFrame([{
        'Mensaje': 'No se encontraron hojas válidas para procesar',
        'Hojas disponibles': ', '.join(list(available_sheets)),
        'Fecha': datetime.now().strftime('%Y-%m-%d')
    }])
    default_data.to_excel(writer, sheet_name=DEFAULT_SHEET_NAME, index=False)

def descargar_reporte3():
    """Descarga el archivo temporal del Reporte 3"""
    try:
        temp_filename = request.json.get('temp_file')
        if not temp_filename:
            return jsonify({'success': False, 'message': 'Nombre de archivo temporal no proporcionado'}), 400
        
        temp_filepath = os.path.join('temp_files', temp_filename)
        
        if not os.path.exists(temp_filepath):
            return jsonify({'success': False, 'message': 'Archivo temporal no encontrado'}), 404
        
        # Obtener nombre dinámico desde metadata
        metadata_filename = temp_filename.replace('.xlsx', '_metadata.json')
        metadata_filepath = os.path.join('temp_files', metadata_filename)
        download_name = '3_Reporte_Reporteria.xlsx'  # Nombre por defecto
        
        if os.path.exists(metadata_filepath):
            try:
                with open(metadata_filepath, 'r', encoding='utf-8') as f:
                    metadata = json.load(f)
                    download_name = metadata.get('download_name', download_name)
            except Exception:
                pass  # Usar nombre por defecto si hay error
        
        return send_file(
            temp_filepath,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=download_name
        )
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al descargar archivo: {str(e)}'}), 500

def continuar_a_paso4():
    """Prepara los datos para continuar al paso 4"""
    try:
        temp_filename = request.json.get('temp_file')
        if not temp_filename:
            return jsonify({'success': False, 'message': 'Nombre de archivo temporal no proporcionado'}), 400
        
        temp_filepath = os.path.join('temp_files', temp_filename)
        
        if not os.path.exists(temp_filepath):
            return jsonify({'success': False, 'message': 'Archivo temporal no encontrado'}), 404
        
        # Verificar que el archivo sea válido
        try:
            pd.read_excel(temp_filepath, sheet_name=0, nrows=1)
        except Exception:
            return jsonify({'success': False, 'message': 'Archivo temporal corrupto'}), 500
        
        return jsonify({
            'success': True,
            'message': 'Listo para continuar al Paso 4',
            'temp_file': temp_filename,
            'redirect_url': '/paso4'
        })
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error al preparar paso 4: {str(e)}'}), 500

# ============================================================================
# FUNCIÓN DE ACTUALIZACIÓN DE BASE DE ASESORES
# ============================================================================
def actualizar_base_asesores():
    """
    Actualiza la base de asesores desde un archivo Excel/CSV subido
    Columnas esperadas: Fecha Ingreso, Fecha, Cedula, ID, EXT, VOIP, Nombre, Sede, Ubicación
    """
    
    base_file = request.files.get('baseAsesorFile')
    
    if not base_file:
        return "Error: Debes subir el archivo de base de asesores.", 400
    
    if not allowed_file(base_file.filename, {'xlsx', 'csv'}):
        return "Error: Formato de archivo no permitido. Usa Excel (.xlsx) o CSV (.csv).", 400
    
    try:
        # Leer archivo según su tipo
        if base_file.filename.lower().endswith('.csv'):
            df = pd.read_csv(base_file)
        else:
            df = pd.read_excel(base_file)
        
        # Validar que existan las columnas requeridas
        missing_columns = [col for col in REQUIRED_ASESOR_COLUMNS if col not in df.columns]
        
        if missing_columns:
            error_msg = f"Error: Faltan las siguientes columnas en el archivo: {', '.join(missing_columns)}"
            return error_msg, 400
        
        # Limpiar y procesar los datos
        df = df.dropna(subset=['ID', 'EXT'])  # Eliminar filas sin ID o EXT
        
        if len(df) == 0:
            return "Error: No se encontraron registros válidos con ID y EXT.", 400
        
        # Convertir a string los campos ID y EXT para evitar problemas de tipo
        df['ID'] = df['ID'].astype(str)
        df['EXT'] = df['EXT'].astype(str)
        
        # Crear lista de diccionarios con las columnas correctas
        asesores_list = []
        for _, row in df.iterrows():
            asesor = {}
            for campo in REQUIRED_ASESOR_COLUMNS:
                if campo == 'Fecha Ingreso':
                    asesor[campo] = limpiar_fecha(row[campo])
                elif campo == 'Nombre':
                    asesor[campo] = formatear_nombre(str(row[campo]) if pd.notna(row[campo]) else '')
                else:
                    asesor[campo] = str(row[campo]) if pd.notna(row[campo]) else ''
            asesores_list.append(asesor)
        
        if not asesores_list:
            return "Error: No se pudieron procesar los registros del archivo.", 400
        
        # Guardar como JSON con backup
        try:
            # Crear backup del archivo actual si existe
            if os.path.exists(BASE_ASESORES_FILE):
                backup_name = f"{BASE_ASESORES_FILE}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copy2(BASE_ASESORES_FILE, backup_name)
        except Exception:
            pass
        
        # Guardar nueva base
        with open(BASE_ASESORES_FILE, 'w', encoding='utf-8') as f:
            json.dump(asesores_list, f, ensure_ascii=False, indent=2)
        
        success_msg = f"✅ Base de asesores actualizada exitosamente. Total: {len(asesores_list)} asesores registrados."
        return success_msg, 200
        
    except pd.errors.EmptyDataError:
        error_msg = "Error: El archivo está vacío o no contiene datos válidos."
        return error_msg, 400
    except pd.errors.ParserError:
        error_msg = "Error: No se pudo leer el archivo. Verifica el formato."
        return error_msg, 400
    except PermissionError:
        error_msg = "Error: No se puede escribir el archivo. Verifica permisos."
        return error_msg, 500
    except Exception as e:
        error_msg = f"Error actualizando la base de asesores: {e}"
        return error_msg, 500