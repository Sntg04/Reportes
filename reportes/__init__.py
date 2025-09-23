"""
Módulo de reportes - Sistema de Reportes Flask
Contiene los 4 procesadores de reportes principales:
- 1_Reporte_Llamadas
- 2_Reporte_Admin_Cobranza
- 3_Reporte_Reporteria
- 4_Reporte_Calidad
"""

import importlib

# Importar módulos con nombres que empiezan con números
_1_Reporte_Llamadas = importlib.import_module('reportes.1_Reporte_Llamadas')
_2_Reporte_Admin_Cobranza = importlib.import_module('reportes.2_Reporte_Admin_Cobranza')
_3_Reporte_Reporteria = importlib.import_module('reportes.3_Reporte_Reporteria')
_4_Reporte_Calidad = importlib.import_module('reportes.4_Reporte_Calidad')

# Módulo comuniquemonos eliminado - solo reportes 1, 2, 3, 4

# Exportar las funciones
procesar_llamadas_isabel = _1_Reporte_Llamadas.procesar_llamadas_isabel
procesar_reporte_agentes = _1_Reporte_Llamadas.procesar_reporte_agentes
procesar_admin_cobranza = _2_Reporte_Admin_Cobranza.procesar_admin_cobranza
procesar_reporteria_cobranza = _3_Reporte_Reporteria.procesar_reporteria_cobranza
descargar_reporte3 = _3_Reporte_Reporteria.descargar_reporte3
continuar_a_paso4 = _3_Reporte_Reporteria.continuar_a_paso4
procesar_reporte_calidad = _4_Reporte_Calidad.procesar_reporte_calidad
descargar_reporte4 = _4_Reporte_Calidad.descargar_reporte4
generar_prueba_reporte4 = _4_Reporte_Calidad.generar_prueba_reporte4
actualizar_base_asesores = _3_Reporte_Reporteria.actualizar_base_asesores

__all__ = [
    'procesar_admin_cobranza',
    'procesar_llamadas_isabel',
    'procesar_reporte_agentes',
    'procesar_reporteria_cobranza',
    'descargar_reporte3',
    'continuar_a_paso4',
    'procesar_reporte_calidad',
    'descargar_reporte4',
    'generar_prueba_reporte4',
    'actualizar_base_asesores'
]
