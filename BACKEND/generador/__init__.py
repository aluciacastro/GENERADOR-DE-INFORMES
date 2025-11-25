"""
Módulo generador de informes ICBF
"""

from .analizador import (
    analizar_columna,
    generar_analisis_resultados,
    generar_oportunidades_mejora,
    extraer_tema_pregunta
)

from .graficas import crear_grafica_circular

from .documento import generar_informe_word

from .utils import convertir_a_png

__version__ = '1.0.0'
__all__ = [
    'analizar_columna',
    'generar_analisis_resultados',
    'generar_oportunidades_mejora',
    'extraer_tema_pregunta',
    'crear_grafica_circular',
    'generar_informe_word',
    'convertir_a_png'
]