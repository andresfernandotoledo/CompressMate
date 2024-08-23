import os
import subprocess
import tempfile
import pdfplumber
from flask import Flask, render_template, request, send_file, abort
from io import BytesIO
from PIL import Image, UnidentifiedImageError
import moviepy.editor as mp
from pdf2image import convert_from_path
from docx2pdf import convert as docx_to_pdf
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from docx import Document
import openpyxl
from openpyxl import Workbook, load_workbook
import pdfkit
from docx2pdf import convert as docx2pdf_convert
from pdf2docx import Converter as PDF2DOCXConverter
from werkzeug.utils import secure_filename
import magic  # Asegúrate de que esta línea esté presente

app = Flask(__name__)

# Funciones de compresión
def compress_image(input_stream, quality=85):
    try:
        output = BytesIO()
        with Image.open(input_stream) as img:
            img = img.convert("RGB")  # Convertir a RGB si la imagen no está en RGB
            img.save(output, format='JPEG', quality=quality)
        output.seek(0)
        return output
    except UnidentifiedImageError:
        raise ValueError("El archivo no es una imagen válida.")
    except Exception as e:
        raise ValueError(f"Ocurrió un error al procesar la imagen: {str(e)}")

def compress_pdf(input_stream, target_size_kb):
    temp_input_path = tempfile.mktemp(suffix='.pdf')
    temp_output_path = tempfile.mktemp(suffix='.pdf')

    with open(temp_input_path, 'wb') as f:
        f.write(input_stream.read())

    quality_settings = ['/screen', '/ebook', '/printer', '/prepress']
    
    for quality in quality_settings:
        gs_command = [
            "gs",
            "-sDEVICE=pdfwrite",
            f"-dPDFSETTINGS={quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={temp_output_path}",
            temp_input_path
        ]
        subprocess.run(gs_command, check=True)
        output_size_kb = os.path.getsize(temp_output_path) / 1024

        if output_size_kb <= target_size_kb:
            break

    with open(temp_output_path, 'rb') as f:
        output_data = f.read()

    os.remove(temp_input_path)
    os.remove(temp_output_path)

    output_stream = BytesIO(output_data)
    output_stream.seek(0)
    return output_stream

def compress_docx(input_stream, target_size_kb):
    temp_input_path = tempfile.mktemp(suffix='.docx')
    temp_output_path = tempfile.mktemp(suffix='.docx')

    with open(temp_input_path, 'wb') as f:
        f.write(input_stream.read())

    doc = Document(temp_input_path)
    doc.save(temp_output_path)

    output_size_kb = os.path.getsize(temp_output_path) / 1024
    if output_size_kb > target_size_kb:
        print("No se pudo comprimir a menos del tamaño deseado.")
    
    with open(temp_output_path, 'rb') as f:
        output_data = f.read()

    os.remove(temp_input_path)
    os.remove(temp_output_path)

    output_stream = BytesIO(output_data)
    output_stream.seek(0)
    return output_stream

def compress_pptx(input_stream, target_size_kb):
    temp_input_path = tempfile.mktemp(suffix='.pptx')
    temp_output_path = tempfile.mktemp(suffix='.pptx')

    with open(temp_input_path, 'wb') as f:
        f.write(input_stream.read())

    ppt = Presentation(temp_input_path)
    ppt.save(temp_output_path)

    output_size_kb = os.path.getsize(temp_output_path) / 1024
    if output_size_kb > target_size_kb:
        print("No se pudo comprimir a menos del tamaño deseado.")
    
    with open(temp_output_path, 'rb') as f:
        output_data = f.read()

    os.remove(temp_input_path)
    os.remove(temp_output_path)

    output_stream = BytesIO(output_data)
    output_stream.seek(0)
    return output_stream

def compress_xlsx(input_stream, target_size_kb):
    temp_input_path = tempfile.mktemp(suffix='.xlsx')
    temp_output_path = tempfile.mktemp(suffix='.xlsx')

    with open(temp_input_path, 'wb') as f:
        f.write(input_stream.read())

    workbook = load_workbook(temp_input_path)
    workbook.save(temp_output_path)

    output_size_kb = os.path.getsize(temp_output_path) / 1024
    if output_size_kb > target_size_kb:
        print("No se pudo comprimir a menos del tamaño deseado.")
    
    with open(temp_output_path, 'rb') as f:
        output_data = f.read()

    os.remove(temp_input_path)
    os.remove(temp_output_path)

    output_stream = BytesIO(output_data)
    output_stream.seek(0)
    return output_stream

def compress_video(input_stream, bitrate="500k"):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.tmp') as temp_input_file:
            temp_input_path = temp_input_file.name
            temp_input_file.write(input_stream.read())
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.mp4') as temp_output_file:
            temp_output_path = temp_output_file.name
        
        video = mp.VideoFileClip(temp_input_path)
        video_resized = video.resize(height=720)
        video_resized.write_videofile(temp_output_path, bitrate=bitrate, codec='libx264', audio_codec='aac')
        os.remove(temp_input_path)
        
        return temp_output_path
    except Exception as e:
        raise ValueError(f"Ocurrió un error al procesar el video: {str(e)}")

# Función para convertir imágenes a diferentes formatos
def convert_image(input_stream, output_format):
    valid_formats = ['jpeg', 'png', 'bmp', 'gif', 'tiff']

    if output_format.lower() not in valid_formats:
        raise ValueError("Formato de imagen no válido. Los formatos válidos son: " + ", ".join(valid_formats))
    
    output = BytesIO()

    try:
        with Image.open(input_stream) as img:
            # Convertir la imagen a RGB si el formato de salida no soporta transparencia
            if output_format.lower() in ['jpeg', 'bmp', 'gif', 'tiff'] and img.mode not in ['RGB', 'L']:
                img = img.convert('RGB')
            elif output_format.lower() == 'png' and img.mode == 'RGBA':
                img = img.convert('RGBA')
            
            img.save(output, format=output_format.upper())
    except UnidentifiedImageError:
        raise ValueError("El archivo no es una imagen válida.")
    except Exception as e:
        raise ValueError(f"Ocurrió un error al convertir la imagen: {str(e)}")

    output.seek(0)
    return output

# Función para convertir PDF a diferentes formatos
def convert_pdf(input_stream, output_format):
    def convert_pdf_to_docx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.docx') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            cv = PDF2DOCXConverter(temp_input_path.name)
            cv.convert(temp_output_path.name, start=0, end=None)
            cv.close()

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_pdf_to_pptx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.pptx') as temp_output_path:
 
            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            with pdfplumber.open(temp_input_path.name) as pdf:
                presentation = Presentation()

                for page in pdf.pages:
                    text = page.extract_text()
                    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
                    textbox = slide.shapes.add_textbox(left=Inches(1), top=Inches(1), width=Inches(8), height=Inches(1))
                    textbox.text = text

                presentation.save(temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_pdf_to_xlsx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.xlsx') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            # Extract data from PDF
            with pdfplumber.open(temp_input_path.name) as pdf:
                workbook = openpyxl.Workbook()
                sheet = workbook.active

                for page in pdf.pages:
                    text = page.extract_text()
                    for i, line in enumerate(text.split('\n'), start=1):
                        sheet[f'A{i}'] = line

                workbook.save(temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_pdf_to_jpeg(input_stream): 
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_input_path:
            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            images = convert_from_path(temp_input_path.name)
            images[0].save(output, format='JPEG')

        output.seek(0)
        return output

    def convert_pdf_to_png(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_input_path:
            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            images = convert_from_path(temp_input_path.name)
            images[0].save(output, format='PNG')

        output.seek(0)
        return output

    if output_format == 'docx':
        return convert_pdf_to_docx(input_stream)
    elif output_format == 'pptx':
        return convert_pdf_to_pptx(input_stream)
    elif output_format == 'xlsx':
        return convert_pdf_to_xlsx(input_stream)
    elif output_format == 'jpeg':
        return convert_pdf_to_jpeg(input_stream)
    elif output_format == 'png':
        return convert_pdf_to_png(input_stream)
    else:
        raise ValueError("Formato de salida no soportado")

# Función para convertir DOCX a diferentes formatos
def convert_docx(input_stream, output_format):
    def convert_docx_to_pdf(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.docx') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            docx2pdf_convert(temp_input_path.name, temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_docx_to_pptx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.docx') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.pptx') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            doc = Document(temp_input_path.name)
            presentation = Presentation()

            for paragraph in doc.paragraphs:
                slide_layout = presentation.slide_layouts[5]
                slide = presentation.slides.add_slide(slide_layout)
                textbox = slide.shapes.add_textbox(left=Inches(1), top=Inches(1), width=Inches(8), height=Inches(1))
                textbox.text = paragraph.text

            presentation.save(temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_docx_to_xlsx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.docx') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.xlsx') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            doc = Document(temp_input_path.name)
            workbook = Workbook()
            sheet = workbook.active

            for i, paragraph in enumerate(doc.paragraphs, start=1):
                sheet[f'A{i}'] = paragraph.text

            workbook.save(temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    if output_format == 'pdf':
        return convert_docx_to_pdf(input_stream)
    elif output_format == 'pptx':
        return convert_docx_to_pptx(input_stream)
    elif output_format == 'xlsx':
        return convert_docx_to_xlsx(input_stream)
    else:
        raise ValueError("Formato de salida no soportado")

# Función para convertir XLSX a diferentes formatos
def convert_xlsx(input_stream, output_format):
    def convert_xlsx_to_pdf(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.xlsx') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.html') as temp_html_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.pdf') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            df = pd.read_excel(temp_input_path.name)
            df.to_html(temp_html_path.name)

            pdfkit.from_file(temp_html_path.name, temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_xlsx_to_pptx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.xlsx') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.pptx') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            wb = load_workbook(temp_input_path.name)
            sheet = wb.active
            presentation = Presentation()

            for row in sheet.iter_rows(values_only=True):
                slide_layout = presentation.slide_layouts[5]
                slide = presentation.slides.add_slide(slide_layout)
                textbox = slide.shapes.add_textbox(left=Inches(1), top=Inches(1), width=Inches(8), height=Inches(1))
                textbox.text = "\t".join(map(str, row))

            presentation.save(temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    def convert_xlsx_to_docx(input_stream):
        output = BytesIO()
        with tempfile.NamedTemporaryFile(delete=True, suffix='.xlsx') as temp_input_path, \
             tempfile.NamedTemporaryFile(delete=True, suffix='.docx') as temp_output_path:

            temp_input_path.write(input_stream.read())
            temp_input_path.flush()

            wb = load_workbook(temp_input_path.name)
            sheet = wb.active
            doc = Document()

            for row in sheet.iter_rows(values_only=True):
                doc.add_paragraph("\t".join(map(str, row)))

            doc.save(temp_output_path.name)

            with open(temp_output_path.name, 'rb') as f:
                output.write(f.read())

        output.seek(0)
        return output

    if output_format == 'pdf':
        return convert_xlsx_to_pdf(input_stream)
    elif output_format == 'pptx':
        return convert_xlsx_to_pptx(input_stream)
    elif output_format == 'docx':
        return convert_xlsx_to_docx(input_stream)
    else:
        raise ValueError("Formato de salida no soportado")

# Función principal de conversión
def convert(input_stream, input_type, output_format):
    if input_type == 'pdf':
        return convert_pdf(input_stream, output_format)
    elif input_type == 'docx':
        return convert_docx(input_stream, output_format)
    elif input_type == 'xlsx':
        return convert_xlsx(input_stream, output_format)
    else:
        raise ValueError("Tipo de entrada no soportado")

# Example usage:
# with open('sample.pdf', 'rb') as f:
#     result = convert(f, 'pdf', 'docx')
#     with open('result.docx', 'wb') as out_f:
#         out_f.write(result.read())

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            abort(400, description="No se ha enviado ningún archivo.")
        
        action_type = request.form.get('action_type')
        file_type = request.form.get('file_type', '')
        quality = int(request.form.get('quality', 85))
        quality_kb = int(request.form.get('quality_kb', 100))
        bitrate = request.form.get('bitrate', '500k')
        output_format = request.form.get('output_format', '')

        # Leer el archivo en un flujo de bytes
        file_stream = BytesIO(file.read())

        # Obtener el nombre y la extensión del archivo
        filename = secure_filename(file.filename)
        
        detected_file_type = detect_file_type(file_stream)

        if file_type != detected_file_type:
            abort(400, description=f"Tipo de archivo detectado ({detected_file_type}) no coincide con el tipo proporcionado ({file_type}).")
        
        file_stream.seek(0)  # Reiniciar el flujo del archivo después de la lectura

        try:
            if action_type == 'compress':
                return handle_compression(file_stream, file_type, output_format, quality, quality_kb, bitrate)
            elif action_type == 'convert':
                return handle_conversion(file_stream, file_type, output_format, filename)
            else:
                raise ValueError("Acción no soportada.")
        except ValueError as e:
            abort(400, description=str(e))
        except Exception as e:
            import traceback
            traceback.print_exc()  # Imprime la traza del error para depuración
            abort(500, description=f"Error interno del servidor: {str(e)}")

    return render_template('index.html')


def detect_file_type(file_stream):
    import magic
    mime = magic.Magic()
    file_stream.seek(0)
    mime_type = mime.from_buffer(file_stream.read(2048))
    file_stream.seek(0)
    print(f"Mime type detectado: {mime_type}")

    if 'image' in mime_type:
        return 'imagen'
    elif 'pdf' in mime_type:
        return 'pdf'
    elif 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' in mime_type:
        return 'docx'
    elif 'application/vnd.openxmlformats-officedocument.presentationml.presentation' in mime_type:
        return 'pptx'
    elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in mime_type:
        return 'xlsx'
    elif 'video' in mime_type:
        return 'video'
    else:
        return 'unknown'


def handle_compression(file_stream, file_type, output_format, quality, quality_kb, bitrate):
    try:
        if file_type == 'imagen':
            if output_format.lower() not in ['jpeg', 'png', 'bmp', 'gif', 'tiff']:
                raise ValueError("Formato de salida no válido para imágenes. Los formatos válidos son: jpeg, png, bmp, gif, tiff.")
            
            output_file = compress_image(file_stream, quality)
            mime_type = {
                'jpeg': 'image/jpeg',
                'png': 'image/png',
                'bmp': 'image/bmp',
                'gif': 'image/gif',
                'tiff': 'image/tiff'
            }.get(output_format.lower(), 'application/octet-stream')
            filename = f'compressed_file.{output_format.lower()}'
        
        elif file_type == 'pdf':
            output_file = compress_pdf(file_stream, quality_kb)
            mime_type = 'application/pdf'
            filename = 'compressed_file.pdf'

        elif file_type == 'docx':
            output_file = compress_docx(file_stream, quality_kb)
            mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            filename = 'compressed_file.docx'

        elif file_type == 'pptx':
            output_file = compress_pptx(file_stream, quality_kb)
            mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            filename = 'compressed_file.pptx'

        elif file_type == 'xlsx':
            output_file = compress_xlsx(file_stream, quality_kb)
            mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            filename = 'compressed_file.xlsx'
        
        elif file_type == 'video':
            output_file_path = compress_video(file_stream, bitrate)
            mime_type = 'video/mp4'
            filename = 'compressed_video.mp4'
            return send_file(output_file_path, as_attachment=True, download_name=filename, mimetype=mime_type)

        else:
            raise ValueError("Tipo de archivo no soportado para compresión.")

        return send_file(output_file, as_attachment=True, download_name=filename, mimetype=mime_type)

    except ValueError as e:
        abort(400, description=str(e))
    except FileNotFoundError as e:
        abort(404, description=str(e))
    except Exception as e:
        import traceback
        traceback.print_exc()  # Imprime la traza del error para depuración
        abort(500, description=f"Error interno del servidor: {str(e)}")

# Función de detección del tipo de archivo
def detect_file_type(file_stream):
    file_stream.seek(0)
    mime = magic.Magic(mime=True)
    mime_type = mime.from_buffer(file_stream.read(2048))
    file_stream.seek(0)  # Resetear el flujo de archivo para futuras lecturas

    print(f"Mime type detectado: {mime_type}")

    if 'image' in mime_type:
        return 'imagen'
    elif 'pdf' in mime_type:
        return 'pdf'
    elif 'vnd.openxmlformats-officedocument.wordprocessingml.document' in mime_type:
        return 'docx'
    elif 'vnd.openxmlformats-officedocument.presentationml.presentation' in mime_type:
        return 'pptx'
    elif 'vnd.openxmlformats-officedocument.spreadsheetml.sheet' in mime_type:
        return 'xlsx'
    elif 'video' in mime_type:
        return 'video'
    else:
        return 'desconocido'

# Función para manejar la conversión de archivos
def handle_conversion(file_stream, file_type, output_format, filename=None):
    output_format = output_format.lower()

    # Si no se proporciona el nombre del archivo, lanza un error
    if not filename:
        abort(400, description="Nombre de archivo no proporcionado.")
    
    # Limpiar el nombre del archivo para evitar problemas de seguridad
    filename = secure_filename(filename)

    # Detectar el tipo de archivo usando el método mejorado
    detected_file_type = detect_file_type(file_stream)

    # Verificación del tipo de archivo detectado y proporcionado
    if file_type != detected_file_type:
        abort(400, description=f"Tipo de archivo no coincide. Detectado: {detected_file_type}, Proporcionado: {file_type}.")

    # Definir los formatos válidos para cada tipo de archivo
    valid_formats = {
        'imagen': ['jpeg', 'png', 'bmp', 'gif', 'tiff'],
        'pdf': ['docx', 'pptx', 'xlsx', 'jpeg', 'png'],
        'docx': ['pdf', 'pptx', 'xlsx'],
        'pptx': ['pdf', 'docx', 'xlsx'],
        'xlsx': ['pdf', 'docx', 'pptx'],
        'video': ['mp4', 'mov', 'wmv', 'avi', 'flv', 'webm']
    }

    # Verificar si el formato de salida es válido para el tipo de archivo
    if output_format not in valid_formats.get(file_type, []):
        abort(400, description=f"Formato de salida no válido para {file_type}. Los formatos válidos son: {', '.join(valid_formats[file_type])}.")

    try:
        # Procesar la conversión según el tipo de archivo
        if file_type == 'imagen':
            output_file = convert_image(file_stream, output_format)
            mime_type = {
                'jpeg': 'image/jpeg',
                'png': 'image/png',
                'bmp': 'image/bmp',
                'gif': 'image/gif',
                'tiff': 'image/tiff'
            }[output_format]
            output_filename = f'converted_image.{output_format}'

        elif file_type == 'pdf':
            output_file = convert_pdf(file_stream, output_format)
            mime_type = {
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'jpeg': 'image/jpeg',
                'png': 'image/png'
            }[output_format]
            output_filename = f'converted_file.{output_format}'

        elif file_type == 'docx':
            output_file = convert_docx(file_stream, output_format)
            mime_type = {
                'pdf': 'application/pdf',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }[output_format]
            output_filename = f'converted_file.{output_format}'

        elif file_type == 'pptx':
            output_file = convert_pptx(file_stream, output_format)
            mime_type = {
                'pdf': 'application/pdf',
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }[output_format]
            output_filename = f'converted_file.{output_format}'

        elif file_type == 'xlsx':
            output_file = convert_xlsx(file_stream, output_format)
            mime_type = {
                'pdf': 'application/pdf',
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            }[output_format]
            output_filename = f'converted_file.{output_format}'

        elif file_type == 'video':
            output_file = convert_video(file_stream, output_format)
            mime_type = {
                'mp4': 'video/mp4',
                'mov': 'video/quicktime',
                'wmv': 'video/x-ms-wmv',
                'avi': 'video/x-msvideo',
                'flv': 'video/x-flv',
                'webm': 'video/webm'
            }[output_format]
            output_filename = f'converted_video.{output_format}'

        else:
            abort(400, description="Tipo de archivo no soportado para conversión.")

        return send_file(output_file, as_attachment=True, download_name=output_filename, mimetype=mime_type)

    except ValueError as e:
        abort(400, description=str(e))
    except Exception as e:
        import traceback
        traceback.print_exc()  # Imprime la traza del error para depuración
        abort(500, description=f"Error interno del servidor: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True)
