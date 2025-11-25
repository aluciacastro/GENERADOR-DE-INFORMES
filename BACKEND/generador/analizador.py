"""
Módulo para análisis de datos de encuestas
Genera análisis inteligente y oportunidades de mejora
"""

import pandas as pd


def analizar_columna(df, columna):
    """
    Analiza una columna del DataFrame
    
    Args:
        df: DataFrame de pandas con los datos
        columna: Nombre de la columna a analizar
        
    Returns:
        dict: Diccionario con resultados del análisis o None si está vacía
    """
    datos_limpios = df[columna].dropna()
    
    if len(datos_limpios) == 0:
        return None
    
    frecuencias = datos_limpios.value_counts()
    total = len(datos_limpios)
    
    # Calcular porcentajes
    porcentajes = {}
    porcentajes_exactos = {}
    
    for valor, conteo in frecuencias.items():
        porcentaje_exacto = (conteo / total) * 100
        porcentajes[str(valor)] = f"{porcentaje_exacto:.1f}"
        porcentajes_exactos[str(valor)] = porcentaje_exacto
    
    return {
        'pregunta': columna,
        'frecuencias': frecuencias.to_dict(),
        'porcentajes': porcentajes,
        'porcentajes_exactos': porcentajes_exactos,
        'total': total
    }


def extraer_tema_pregunta(pregunta):
    """
    Extrae el tema principal de una pregunta para el análisis
    
    Args:
        pregunta: Texto de la pregunta
        
    Returns:
        str: Tema identificado
    """
    pregunta_lower = pregunta.lower()
    
    # Diccionario de temas comunes
    temas = {
        ('calidad', 'servicio'): 'la calidad del servicio',
        ('atención', 'niño', 'hijo'): 'la atención a los niños',
        ('espacio', 'ambiente', 'infraestructura'): 'la organización del espacio',
        ('talento', 'equipo', 'agente', 'personal', 'docente'): 'el compromiso del talento humano',
        ('aliment', 'comida', 'menú', 'complemento'): 'los complementos alimentarios',
        ('comunicación', 'información'): 'la comunicación con las familias',
        ('pedagógic', 'actividad', 'enseñanza'): 'las actividades pedagógicas',
        ('familia', 'participación'): 'la participación familiar',
        ('higiene', 'limpieza', 'aseo'): 'las condiciones de higiene',
        ('seguridad', 'protección'): 'las medidas de seguridad',
        ('queja', 'reclamo'): 'la atención a quejas y reclamos',
        ('material', 'recurso'): 'los materiales y recursos',
        ('horario', 'tiempo'): 'la organización de horarios',
        ('relaciones', 'interpersonal'): 'las relaciones interpersonales',
    }
    
    # Buscar coincidencias
    for palabras_clave, tema in temas.items():
        if any(palabra in pregunta_lower for palabra in palabras_clave):
            return tema
    
    # Si no encuentra tema específico, extraer de la pregunta
    if len(pregunta) > 50:
        return f"los aspectos relacionados con {pregunta[:47].lower()}..."
    else:
        return f"los aspectos de {pregunta.lower()}"


def generar_analisis_resultados(resultados_preguntas, nombre_uds):
    """
    Genera análisis de resultados con lenguaje técnico e institucional para el ICBF
    Un solo párrafo, profesional y objetivo
    
    Args:
        resultados_preguntas: Lista de resultados de análisis
        nombre_uds: Nombre de la Unidad de Servicio
        
    Returns:
        str: Texto del análisis
    """
    if not resultados_preguntas:
        return "No se encontraron resultados para analizar."
    
    # Calcular estadísticas generales
    total_preguntas = len(resultados_preguntas)
    porcentajes_todos = [float(list(r['porcentajes'].values())[0]) for r in resultados_preguntas]
    promedio_general = sum(porcentajes_todos) / len(porcentajes_todos) if porcentajes_todos else 0
    
    porcentajes_altos = sum(1 for p in porcentajes_todos if p >= 90)
    porcentajes_excelentes = sum(1 for p in porcentajes_todos if p >= 95)
    porcentajes_buenos = sum(1 for p in porcentajes_todos if 80 <= p < 90)
    porcentajes_regulares = sum(1 for p in porcentajes_todos if p < 80)
    
    # Identificar las 3 fortalezas principales (porcentajes más altos)
    preguntas_ordenadas = sorted(
        resultados_preguntas,
        key=lambda r: float(list(r['porcentajes'].values())[0]),
        reverse=True
    )
    
    fortalezas = []
    for resultado in preguntas_ordenadas[:3]:
        porcentaje = float(list(resultado['porcentajes'].values())[0])
        if porcentaje >= 85:
            tema = extraer_tema_pregunta(resultado['pregunta'])
            fortalezas.append((tema, porcentaje))
    
    # Identificar áreas con valoraciones menores
    areas_menores = []
    preguntas_bajas = sorted(
        [r for r in resultados_preguntas if float(list(r['porcentajes'].values())[0]) < 90],
        key=lambda r: float(list(r['porcentajes'].values())[0])
    )
    
    for resultado in preguntas_bajas[:3]:
        porcentaje = float(list(resultado['porcentajes'].values())[0])
        tema = extraer_tema_pregunta(resultado['pregunta'])
        areas_menores.append((tema, porcentaje))
    
    # Construir análisis con lenguaje técnico e institucional
    partes_analisis = []
    
    # Introducción con valoración global
    if promedio_general >= 90:
        valoracion = "índice de satisfacción altamente favorable"
    elif promedio_general >= 85:
        valoracion = "índice de satisfacción favorable"
    elif promedio_general >= 80:
        valoracion = "índice de satisfacción satisfactorio"
    else:
        valoracion = "índice de satisfacción que evidencia oportunidades de mejora"
    
    partes_analisis.append(
        f"El análisis de los resultados obtenidos en la Unidad de Servicio {nombre_uds} evidencia un {valoracion}, "
        f"con un promedio general de {promedio_general:.1f}% en los {total_preguntas} ítems evaluados."
    )
    
    # Distribución de resultados
    if porcentajes_excelentes > 0:
        partes_analisis.append(
            f"Se destaca que {porcentajes_excelentes} de los {total_preguntas} aspectos consultados ({(porcentajes_excelentes/total_preguntas*100):.0f}%) "
            f"registran valoraciones superiores al 95%, lo cual refleja un alto nivel de percepción positiva por parte de las familias usuarias."
        )
    elif porcentajes_altos > 0:
        partes_analisis.append(
            f"Del total de aspectos evaluados, {porcentajes_altos} ({(porcentajes_altos/total_preguntas*100):.0f}%) "
            f"presentan valoraciones superiores al 90%, evidenciando una percepción favorable del servicio prestado."
        )
    
    # Fortalezas identificadas
    if fortalezas:
        if len(fortalezas) == 1:
            fortalezas_texto = f"{fortalezas[0][0]} ({fortalezas[0][1]:.1f}%)"
        elif len(fortalezas) == 2:
            fortalezas_texto = f"{fortalezas[0][0]} ({fortalezas[0][1]:.1f}%) y {fortalezas[1][0]} ({fortalezas[1][1]:.1f}%)"
        else:
            fortalezas_texto = f"{fortalezas[0][0]} ({fortalezas[0][1]:.1f}%), {fortalezas[1][0]} ({fortalezas[1][1]:.1f}%) y {fortalezas[2][0]} ({fortalezas[2][1]:.1f}%)"
        
        partes_analisis.append(
            f"Las principales fortalezas identificadas corresponden a {fortalezas_texto}, "
            f"aspectos que demuestran el compromiso institucional con la calidad de la atención integral a la primera infancia."
        )
    
    # Áreas con valoraciones menores (si existen)
    if areas_menores:
        if len(areas_menores) == 1:
            areas_texto = f"{areas_menores[0][0]} ({areas_menores[0][1]:.1f}%)"
        elif len(areas_menores) == 2:
            areas_texto = f"{areas_menores[0][0]} ({areas_menores[0][1]:.1f}%) y {areas_menores[1][0]} ({areas_menores[1][1]:.1f}%)"
        else:
            areas_texto = f"{areas_menores[0][0]} ({areas_menores[0][1]:.1f}%), {areas_menores[1][0]} ({areas_menores[1][1]:.1f}%) y {areas_menores[2][0]} ({areas_menores[2][1]:.1f}%)"
        
        if porcentajes_regulares > 0:
            partes_analisis.append(
                f"No obstante, se identifican oportunidades de mejora en aspectos como {areas_texto}, "
                f"los cuales requieren acciones de fortalecimiento para alcanzar los estándares de excelencia esperados."
            )
        else:
            partes_analisis.append(
                f"Si bien los resultados son favorables en su conjunto, aspectos como {areas_texto} "
                f"presentan valoraciones ligeramente menores que podrían optimizarse mediante estrategias focalizadas de mejoramiento."
            )
    else:
        partes_analisis.append(
            f"Los resultados evidencian una gestión integral que responde satisfactoriamente a las expectativas "
            f"de las familias beneficiarias, consolidando la Unidad de Servicio como un referente de calidad "
            f"en la atención a la primera infancia."
        )
    
    return " ".join(partes_analisis)


def generar_oportunidades_mejora(resultados_preguntas):
    """
    Genera máximo 3 oportunidades de mejora con lenguaje técnico e institucional
    Propuestas concretas, viables y directamente relacionadas con los ítems de menor calificación
    
    Args:
        resultados_preguntas: Lista de resultados de análisis
        
    Returns:
        list: Lista con máximo 3 oportunidades de mejora
    """
    oportunidades = []
    
    # Ordenar preguntas por porcentaje (de menor a mayor)
    preguntas_ordenadas = sorted(
        resultados_preguntas,
        key=lambda r: float(list(r['porcentajes'].values())[0])
    )
    
    # Generar oportunidades para las 3 preguntas con porcentajes más bajos
    for resultado in preguntas_ordenadas[:3]:
        primera_respuesta = list(resultado['porcentajes'].items())[0]
        porcentaje = float(primera_respuesta[1])
        pregunta = resultado['pregunta']
        
        # Solo generar oportunidad si el porcentaje es menor a 95% (hay margen de mejora)
        if porcentaje < 95:
            oportunidad = _generar_oportunidad_institucional(pregunta, porcentaje)
            if oportunidad and oportunidad not in oportunidades:
                oportunidades.append(oportunidad)
    
    # Si no hay suficientes oportunidades específicas, agregar generales
    if len(oportunidades) < 3:
        oportunidades_generales = [
            "Implementar mecanismos de seguimiento y evaluación continua del servicio mediante instrumentos estandarizados que permitan identificar oportunamente aspectos susceptibles de mejora.",
            "Fortalecer los procesos de formación y acompañamiento a familias, incorporando metodologías participativas y contenidos pertinentes según las necesidades identificadas.",
            "Establecer protocolos de aseguramiento de la calidad que garanticen el mantenimiento de los estándares en todos los componentes del servicio de atención integral."
        ]
        
        for oportunidad_general in oportunidades_generales:
            if len(oportunidades) < 3:
                oportunidades.append(oportunidad_general)
    
    # Retornar máximo 3
    return oportunidades[:3]


def _generar_oportunidad_institucional(pregunta, porcentaje):
    """
    Genera oportunidades de mejora con lenguaje técnico e institucional apropiado para el ICBF
    Propuestas concretas, viables y específicas
    
    Args:
        pregunta: Texto de la pregunta
        porcentaje: Porcentaje obtenido
        
    Returns:
        str: Texto de la oportunidad de mejora
    """
    pregunta_lower = pregunta.lower()
    
    # Diccionario ampliado de temas con oportunidades redactadas institucionalmente
    temas_oportunidades = {
        # Alimentación y nutrición
        ('aliment', 'comida', 'menú', 'complemento', 'nutrición'): 
            "Realizar evaluaciones sensoriales y nutricionales periódicas de los complementos alimentarios, incorporando la retroalimentación de las familias para ajustar menús y garantizar su aceptabilidad y aporte nutricional.",
        
        # Comunicación e información
        ('comunicación', 'información', 'mensaje', 'notificación'): 
            "Fortalecer las estrategias de comunicación institucional mediante la diversificación de canales (digitales y presenciales), estableciendo protocolos de información clara, oportuna y pertinente sobre las actividades y procesos pedagógicos.",
        
        # Atención y servicio al usuario
        ('atención', 'atender', 'trato', 'servicio al usuario'): 
            "Implementar un plan de mejoramiento del servicio al usuario que incluya capacitación en atención humanizada, protocolos de respuesta oportuna y mecanismos de verificación de la satisfacción en cada punto de contacto.",
        
        # Espacios físicos e infraestructura
        ('espacio', 'ambiente', 'infraestructura', 'instalacion', 'área'): 
            "Desarrollar un plan de adecuación y mantenimiento de espacios físicos que garantice condiciones óptimas de seguridad, funcionalidad y ambientación pedagógica, conforme a los estándares técnicos establecidos por el ICBF.",
        
        # Calidad del servicio
        ('calidad', 'servicio', 'prestación'): 
            "Implementar un sistema de gestión de calidad que incluya indicadores de desempeño, auditorías internas periódicas y planes de mejoramiento continuo en todos los componentes de la atención integral.",
        
        # Talento humano y personal
        ('personal', 'talento', 'equipo', 'agente', 'docente', 'maestr', 'profesional'): 
            "Diseñar e implementar un plan de desarrollo del talento humano que contemple formación continua, acompañamiento técnico y estrategias de bienestar laboral para fortalecer las competencias del equipo interdisciplinario.",
        
        # Pedagogía y procesos educativos
        ('pedagógic', 'actividad', 'enseñanza', 'aprendizaje', 'educativ', 'didáctic'): 
            "Enriquecer las prácticas pedagógicas mediante la implementación de metodologías innovadoras, incorporación de recursos didácticos pertinentes y evaluación sistemática del desarrollo infantil conforme a los referentes técnicos.",
        
        # Participación y vinculación familiar
        ('familia', 'padre', 'madre', 'participación', 'acudiente'): 
            "Fortalecer la vinculación de las familias mediante estrategias diferenciadas de participación que promuevan su rol como agentes educadores y corresponsables en el desarrollo integral de los niños y niñas.",
        
        # Sistema de quejas y sugerencias
        ('queja', 'reclamo', 'sugerencia', 'PQRS'): 
            "Optimizar el sistema de atención a Peticiones, Quejas, Reclamos y Sugerencias (PQRS), garantizando tiempos de respuesta oportunos, seguimiento efectivo y análisis de tendencias para la mejora continua.",
        
        # Seguridad y protección
        ('seguridad', 'protección', 'riesgo', 'prevención'): 
            "Fortalecer los protocolos de seguridad y protección integral mediante la actualización de rutas de atención, capacitación permanente del personal y realización de simulacros periódicos conforme a la normativa vigente.",
        
        # Higiene y saneamiento
        ('higiene', 'limpieza', 'aseo', 'saneamiento', 'desinfección'): 
            "Reforzar los protocolos de higiene, limpieza y desinfección mediante cronogramas estructurados, listas de verificación diarias y capacitación continua al personal de servicios generales conforme a normativa sanitaria.",
        
        # Organización de tiempos y horarios
        ('horario', 'tiempo', 'puntualidad', 'jornada'): 
            "Optimizar la distribución de tiempos pedagógicos y rutinas diarias, garantizando el cumplimiento de la programación establecida y el aprovechamiento efectivo de las jornadas de atención.",
        
        # Recursos y materiales didácticos
        ('material', 'recurso', 'dotación', 'juguete', 'didáctico'): 
            "Fortalecer la dotación de materiales didácticos mediante la evaluación de necesidades, selección de recursos pertinentes al desarrollo infantil y establecimiento de protocolos de mantenimiento y renovación.",
        
        # Salud y bienestar
        ('salud', 'enfermedad', 'vacuna', 'control'): 
            "Fortalecer el componente de salud mediante el seguimiento sistemático del estado de salud de los niños y niñas, articulación con el sector salud y promoción de hábitos saludables con las familias.",
        
        # Proceso de valoración del desarrollo
        ('valoración', 'evaluación', 'desarrollo', 'seguimiento'): 
            "Mejorar los procesos de valoración y seguimiento al desarrollo infantil mediante la aplicación rigurosa de instrumentos estandarizados y la socialización oportuna de resultados con las familias.",
    }
    
    # Buscar el tema en la pregunta
    for palabras_clave, oportunidad in temas_oportunidades.items():
        if any(palabra in pregunta_lower for palabra in palabras_clave):
            return oportunidad
    
    # Si no encuentra tema específico, generar oportunidad institucional basada en la pregunta
    return f"Implementar acciones de mejoramiento específicas relacionadas con el aspecto evaluado en el ítem '{pregunta}', mediante la identificación de causas, establecimiento de metas claras y seguimiento periódico a los resultados."