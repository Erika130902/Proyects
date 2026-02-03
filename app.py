from flask import Flask, render_template, request, send_file
import io
from word import crear_documento_word
from word2 import crear_documento_word2

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('menu.html')

@app.route('/plantilla1')
def plantilla1():
    return render_template('plantilla1.html')

@app.route('/plantilla2')
def plantilla2():
    return render_template('plantilla2.html')

@app.route('/descargar', methods=['POST'])
def descargar():
    # Obtener datos del formulario
    datos = {
        'nombre3': request.form.get('nombre3', ''),
        'telefono3': request.form.get('telefono3', ''),
        'correo3': request.form.get('correo3', ''),
        'nombre4': request.form.get('nombre4', ''),
        'telefono4': request.form.get('telefono4', ''),
        'correo4': request.form.get('correo4', '')
    }
    
    # Crear documento
    doc = crear_documento_word(datos)
    
    # Guardar en memoria
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return send_file(
        file_stream,
        as_attachment=True,
        download_name='Escalamiento_Sin_Residentes.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

@app.route('/descargar2', methods=['POST'])
def descargar2():
    # ========== OBTENER ARRAYS DE RESIDENTES ==========
    residentes_nombres = request.form.getlist('residente_nombre[]')  # Array
    residentes_telefonos = request.form.getlist('residente_telefono[]')  # Array
    
    # Obtener datos del formulario
    datos = {
        # Residentes (Fila 1)
        'residentes_nombres': residentes_nombres,
        'residentes_telefonos': residentes_telefonos,
        'correo_residentes': request.form.get('correo_residentes', ''),
        
        # Fila 3 - Gestor (nombre2 en tu HTML)
        'nombre2': request.form.get('nombre2', ''),
        'telefono2': request.form.get('telefono2', ''),
        'correo2': request.form.get('correo2', ''),
        
        # Fila 4 - Gerente (nombre3 en tu HTML)
        'nombre3': request.form.get('nombre3', ''),
        'telefono3': request.form.get('telefono3', ''),
        'correo3': request.form.get('correo3', '')
    }
    
    # Crear documento con la función correcta
    doc2 = crear_documento_word2(datos)  # Llamar a la función correcta
    
    # Guardar en memoria
    file_stream = io.BytesIO()
    doc2.save(file_stream)
    file_stream.seek(0)
    
    return send_file(
        file_stream,
        as_attachment=True,
        download_name='Escalamiento_Con_Residentes.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True)