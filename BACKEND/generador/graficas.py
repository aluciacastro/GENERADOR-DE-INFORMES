"""
Módulo para generación de gráficas circulares
"""

import matplotlib
matplotlib.use('Agg')  # Usar backend sin GUI
import matplotlib.pyplot as plt
from io import BytesIO


def crear_grafica_circular(datos_dict, titulo):
    """
    Crea una gráfica circular con el título de la pregunta
    Tamaño: 9.28cm ancho x 5.74cm alto
    
    Args:
        datos_dict: Diccionario con los datos {etiqueta: valor}
        titulo: Título de la gráfica
        
    Returns:
        BytesIO: Buffer con la imagen de la gráfica en formato PNG
    """
    labels = list(datos_dict.keys())
    sizes = list(datos_dict.values())
    
    colors = ['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6', 
              '#ec4899', '#14b8a6', '#f97316', '#06b6d4', '#84cc16']
    
    # Crear figura con dimensiones específicas (9.28cm ancho x 5.74cm alto)
    # Convertir cm a pulgadas: 1 cm = 0.393701 pulgadas
    ancho_pulgadas = 9.28 * 0.393701  # ≈ 3.65 pulgadas
    alto_pulgadas = 5.74 * 0.393701   # ≈ 2.26 pulgadas
    
    fig, ax = plt.subplots(figsize=(ancho_pulgadas, alto_pulgadas))
    
    # Título de la gráfica
    plt.title(titulo, fontsize=10, fontweight='bold', pad=8, wrap=True)
    
    # Gráfica circular
    wedges, texts, autotexts = ax.pie(
        sizes, 
        labels=labels, 
        colors=colors[:len(labels)],
        autopct='%1.1f%%',
        startangle=90,
        textprops={'fontsize': 9}
    )
    
    # Formatear textos
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(10)
    
    for text in texts:
        text.set_fontsize(8)
    
    ax.axis('equal')
    
    # Guardar en buffer
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
    plt.close()
    img_buffer.seek(0)
    
    return img_buffer