"""
API Flask para generador de informes ICBF
Backend RESTful para la aplicación web
Adaptado para Vercel Serverless
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import tempfile
from datetime import datetime
import sys

# Agregar el directorio del generador al path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'generador'))

from generador import generar_informe_word
from generador.utils import extraer_nombre_uds, generar_nombre_salida

app = Flask(__name__)
CORS(app)  # Habilitar CORS para permitir peticiones del frontend

# Configuración
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg'}

# Usar /tmp en Vercel (único directorio escribible)
TEMP_DIR = '/tmp' if os.path.exists('/tmp') else tempfile.gettempdir()


def allowed_file(filename, allowed_extensions):
    """Verifica si el archivo tiene una extensión permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


@app.route('/')
def index():
    """Endpoint de bienvenida"""
    return jsonify({
        'message': 'API Generador de Informes ICBF',
        'version': '1.0.0',
        'endpoints': {
            'health': '/api/health',
            'generate': '/api/generate (POST)'
        }
    })


@app.route('/api/health')
def health():
    """Endpoint de salud del servicio"""
    return jsonify({
        'status': 'healthy',
        'timestamp': datetime.now().isoformat(),
        'temp_dir': TEMP_DIR
    })


@app.route('/api/generate', methods=['POST'])
def generate_report():
    """
    Genera un informe a partir de un archivo Excel
    
    Parámetros:
        - excel_file: Archivo Excel con los datos (requerido)
        - nombre_uds: Nombre de la UDS (opcional)
        - encabezado: Imagen del encabezado (opcional)
        - pie: Imagen del pie (opcional)
    
    Returns:
        Archivo .docx generado
    """
    temp_work_dir = None
    try:
        # Validar que se envió el archivo Excel
        if 'excel_file' not in request.files:
            return jsonify({'error': 'No se envió el archivo Excel'}), 400
        
        excel_file = request.files['excel_file']
        
        if excel_file.filename == '':
            return jsonify({'error': 'Nombre de archivo vacío'}), 400
        
        if not allowed_file(excel_file.filename, ALLOWED_EXTENSIONS):
            return jsonify({'error': 'Formato de archivo no válido. Use .xlsx o .xls'}), 400
        
        # Crear directorio temporal único para este request
        temp_work_dir = tempfile.mkdtemp(dir=TEMP_DIR)
        
        # Guardar archivo Excel
        excel_filename = secure_filename(excel_file.filename)
        excel_path = os.path.join(temp_work_dir, excel_filename)
        excel_file.save(excel_path)
        
        # Obtener nombre UDS (del formulario o del archivo)
        nombre_uds = request.form.get('nombre_uds', '')
        if not nombre_uds:
            nombre_uds = extraer_nombre_uds(excel_path)
        
        # Guardar imágenes si se enviaron
        if 'encabezado' in request.files:
            encabezado_file = request.files['encabezado']
            if encabezado_file.filename != '' and allowed_file(encabezado_file.filename, ALLOWED_IMAGE_EXTENSIONS):
                ext = encabezado_file.filename.rsplit('.', 1)[1].lower()
                encabezado_path = os.path.join(temp_work_dir, f'encabezado.{ext}')
                encabezado_file.save(encabezado_path)
        
        if 'pie' in request.files:
            pie_file = request.files['pie']
            if pie_file.filename != '' and allowed_file(pie_file.filename, ALLOWED_IMAGE_EXTENSIONS):
                ext = pie_file.filename.rsplit('.', 1)[1].lower()
                pie_path = os.path.join(temp_work_dir, f'pie.{ext}')
                pie_file.save(pie_path)
        
        # Generar nombre del archivo de salida
        output_filename = generar_nombre_salida(excel_path)
        output_path = os.path.join(temp_work_dir, output_filename)
        
        # Generar el informe
        resultado = generar_informe_word(
            archivo_excel=excel_path,
            archivo_salida=output_path,
            nombre_uds=nombre_uds,
            directorio_trabajo=temp_work_dir
        )
        
        # Leer el archivo generado en memoria
        with open(output_path, 'rb') as f:
            file_data = f.read()
        
        # Limpiar archivos temporales
        import shutil
        if temp_work_dir and os.path.exists(temp_work_dir):
            shutil.rmtree(temp_work_dir)
        
        # Crear respuesta con el archivo en memoria
        from io import BytesIO
        file_stream = BytesIO(file_data)
        file_stream.seek(0)
        
        response = send_file(
            file_stream,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
        response.headers['Content-Disposition'] = f'attachment; filename="{output_filename}"'
        
        return response
        
    except Exception as e:
        # Limpiar en caso de error
        if temp_work_dir and os.path.exists(temp_work_dir):
            import shutil
            shutil.rmtree(temp_work_dir, ignore_errors=True)
        
        print(f"Error al generar informe: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/validate', methods=['POST'])
def validate_file():
    """
    Valida un archivo Excel antes de generar el informe
    
    Returns:
        Información sobre el archivo (número de columnas, respuestas, etc.)
    """
    temp_path = None
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'No se envió el archivo Excel'}), 400
        
        excel_file = request.files['excel_file']
        
        if not allowed_file(excel_file.filename, ALLOWED_EXTENSIONS):
            return jsonify({'error': 'Formato no válido'}), 400
        
        # Guardar temporalmente
        temp_path = os.path.join(TEMP_DIR, secure_filename(excel_file.filename))
        excel_file.save(temp_path)
        
        # Leer y validar
        import pandas as pd
        df = pd.read_excel(temp_path)
        
        # Limpiar archivo temporal
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
        
        return jsonify({
            'valid': True,
            'total_respuestas': len(df),
            'total_columnas': len(df.columns),
            'columnas': list(df.columns),
            'nombre_uds_detectado': extraer_nombre_uds(excel_file.filename)
        })
        
    except Exception as e:
        # Limpiar en caso de error
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
        return jsonify({'error': str(e)}), 500


# Para desarrollo local
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

# Para Vercel serverless (IMPORTANTE)
# Vercel busca una variable llamada 'app' o 'application'
application = app