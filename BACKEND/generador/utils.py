"""
Utilidades para el generador de informes
"""

import os
from PIL import Image


def convertir_a_png(archivo):
    """
    Convierte una imagen a PNG si es necesario (ej: WebP)
    
    Args:
        archivo: Ruta del archivo de imagen
        
    Returns:
        str: Ruta del archivo PNG
    """
    try:
        img = Image.open(archivo)
        # Si es WebP u otro formato, convertir a PNG
        if img.format != 'PNG':
            nombre_base = os.path.splitext(archivo)[0]
            archivo_png = f"{nombre_base}_converted.png"
            img = img.convert('RGB')
            img.save(archivo_png, 'PNG')
            return archivo_png
        return archivo
    except Exception as e:
        print(f"  ⚠️ Error al convertir imagen: {e}")
        return archivo


def extraer_nombre_uds(archivo_excel):
    """
    Extrae el nombre del UDS del nombre del archivo Excel
    
    Args:
        archivo_excel: Ruta del archivo Excel
        
    Returns:
        str: Nombre del UDS formateado
    """
    nombre_archivo = os.path.basename(archivo_excel)
    nombre_sin_ext = os.path.splitext(nombre_archivo)[0]
    # Reemplazar guiones bajos por espacios y capitalizar
    nombre_uds = nombre_sin_ext.replace('_', ' ').title()
    return nombre_uds


def generar_nombre_salida(archivo_excel):
    """
    Genera el nombre del archivo de salida basado en el Excel
    Formato: "informe NOMBRE_UDS.docx"
    
    Args:
        archivo_excel: Ruta del archivo Excel
        
    Returns:
        str: Nombre del archivo de salida
    """
    nombre_uds = extraer_nombre_uds(archivo_excel)
    return f"informe {nombre_uds}.docx"


# Configuración del encabezado y pie de página
ENCABEZADO = {
    'linea1': 'ASOCIACION DE PADRES DE FAMILIA DEL HOGAR INFANTIL GUATAPURI',
    'linea2': 'NIT: 892301280-4',
    'linea3': 'Resolución personería jurídica N°10597 del 20 de septiembre de 1983'
}

PIE_PAGINA = {
    'linea1': 'Dirección: Manzana 34 casa 1 Garupal segunda etapa',
    'linea2': 'Teléfono: 5878818-3178209014',
    'linea3': 'Correo: higuatapuri@gmail.com'
}

# Columnas a excluir del análisis
COLUMNAS_EXCLUIR = [
    'marca temporal',
    'dirección de correo electrónico',
    'direccion de correo electronico',
    'nombre padre/madre del menor- gestante',
    'nombre padre madre del menor gestante',
    'timestamp',
    'email',
    'correo'
]