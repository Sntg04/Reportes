"""
Sistema de Reportes Flask - Aplicación Principal
Aplicación modular que integra múltiples procesadores de reportes
"""

import logging
import os
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, send_file, request, jsonify
from utils.file_utils import setup_logging

# Importar procesadores de reportes
from reportes import (
    procesar_admin_cobranza,
    procesar_llamadas_isabel,
    procesar_reporte_agentes,
    procesar_reporteria_cobranza,
    descargar_reporte3,
    continuar_a_paso4,
    procesar_reporte_calidad,
    descargar_reporte4,
    generar_prueba_reporte4,
    actualizar_base_asesores
)

# --- Configurar logging ---
setup_logging()
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ==============================================================================
# CONFIGURACIÓN DE LA APLICACIÓN
# ==============================================================================
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB máximo para archivos

# ==============================================================================
# DECORADOR PARA MANEJO DE ERRORES
# ==============================================================================
def handle_errors(f):
    """Decorador para manejo común de errores en endpoints"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error en {f.__name__}: {e}")
            return f"Error interno del servidor: {e}", 500
    return decorated_function

# ==============================================================================
# RUTAS DE VISTA (Templates HTML) - FLUJO PASO A PASO
# ==============================================================================
@app.route('/')
def vista_bienvenida():
    """Página de bienvenida con redirección automática"""
    return render_template('bienvenida.html')

@app.route('/paso1')
def vista_paso1():
    """Paso 1: Reportes de Llamadas"""
    return render_template('paso1_llamadas.html')

@app.route('/paso2')
def vista_paso2():
    """Paso 2: Reporte Admin Cobranza"""
    return render_template('paso2_admin_cobranza.html')

@app.route('/paso3')
def vista_paso3():
    """Paso 3: Reporte Reportería"""
    return render_template('paso3_reporteria.html')

@app.route('/paso4')
def vista_paso4():
    """Paso 4: Reporte Calidad"""
    return render_template('paso4_calidad.html')

# Paso 5 eliminado - solo mantenemos reportes 1, 2, 3, 4

# ==============================================================================
# RUTAS DE PROCESAMIENTO (Endpoints de API)  
# =============================================================================="

# ==============================================================================
# RUTAS DE PROCESAMIENTO (Endpoints de API)
# ==============================================================================
@app.route('/procesar-admin-cobranza', methods=['POST'])
@handle_errors
def endpoint_admin_cobranza():
    """Endpoint para procesar 2_Reporte_Admin_Cobranza"""
    return procesar_admin_cobranza()

@app.route('/procesar-llamadas-isabel', methods=['POST'])
@handle_errors
def endpoint_llamadas_isabel():
    """Endpoint para procesar 1_Reporte_Llamadas"""
    return procesar_llamadas_isabel()

@app.route('/procesar-reporte-agentes', methods=['POST'])
@handle_errors
def endpoint_reporte_agentes():
    """Endpoint para procesar Reporte de Agentes"""
    return procesar_reporte_agentes()

@app.route('/procesar-reporteria-cobranza', methods=['POST'])
@handle_errors
def endpoint_reporteria_cobranza():
    """Endpoint para procesar 3_Reporte_Reporteria"""
    return procesar_reporteria_cobranza()

@app.route('/descargar-reporte3', methods=['POST'])
@handle_errors
def endpoint_descargar_reporte3():
    """Endpoint para descargar el archivo temporal del Reporte 3"""
    return descargar_reporte3()

@app.route('/continuar-a-paso4', methods=['POST'])
@handle_errors
def endpoint_continuar_paso4():
    """Endpoint para continuar al paso 4 con el archivo del Reporte 3"""
    return continuar_a_paso4()

@app.route('/test-connection', methods=['GET', 'POST'])
def test_connection():
    """Endpoint de prueba para verificar conectividad"""
    return jsonify({
        'status': 'success',
        'message': 'Conexión OK',
        'method': request.method,
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    })

@app.route('/procesar-reporte-calidad', methods=['POST'])
@handle_errors
def endpoint_reporte_calidad():
    """Endpoint para procesar 4_Reporte_Calidad"""
    return procesar_reporte_calidad()

@app.route('/procesar-y-descargar-reporte4', methods=['POST'])
def endpoint_procesar_y_descargar_reporte4():
    """Endpoint que procesa y descarga directamente el Reporte 4 CON procesamiento de biométricos"""
    try:
        # Usar la función corregida que incluye procesamiento de biométricos
        response = procesar_reporte_calidad()
        
        # Si la respuesta es exitosa, extraer el filename y descargar
        if hasattr(response, 'json') and response.json.get('success'):
            filename = response.json.get('filename')
            temp_filepath = os.path.join('temp_files', filename)
            
            return send_file(
                temp_filepath,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            # Si hay error, devolver la respuesta JSON original
            return response
        
    except Exception as e:
        from flask import jsonify
        return jsonify({
            'success': False,
            'message': f'Error procesando reporte de calidad: {str(e)}'
        }), 500

@app.route('/descargar-reporte4', methods=['POST'])
@handle_errors
def endpoint_descargar_reporte4():
    """Endpoint para descargar el archivo temporal del Reporte 4"""
    return descargar_reporte4()

@app.route('/descargar-reporte4-directo/<filename>')
def endpoint_descargar_reporte4_directo(filename):
    """Endpoint para descarga directa del archivo del Reporte 4"""
    try:
        temp_filepath = os.path.join('temp_files', filename)
        if not os.path.exists(temp_filepath):
            return "Archivo no encontrado", 404
        
        return send_file(
            temp_filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return f"Error al descargar archivo: {str(e)}", 500

@app.route('/generar-prueba-reporte4', methods=['POST'])
@handle_errors
def endpoint_generar_prueba_reporte4():
    """Endpoint para generar archivo de prueba del Reporte 4"""
    return generar_prueba_reporte4()

# Endpoint de comuniquémonos eliminado - solo reportes 1, 2, 3, 4

@app.route('/actualizar-base-asesores', methods=['POST'])
@handle_errors
def endpoint_actualizar_base_asesores():
    """Endpoint para actualizar base de asesores desde archivo Excel/CSV"""
    return actualizar_base_asesores()

# ==============================================================================
# MAIN - EJECUTAR APLICACIÓN
# ==============================================================================
if __name__ == '__main__':
    logger.info("🚀 Iniciando Sistema de Reportes Flask")
    logger.info("📊 Módulos de reportes cargados: 4 procesadores principales (1-Llamadas, 2-Admin/Cobranza, 3-Reportería, 4-Calidad)")
    app.run(debug=True, port=5000, host='0.0.0.0')
