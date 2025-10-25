"""
Utilidades optimizadas para la generación de certificados usando plantillas Word
"""
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm, Pt
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import qrcode
import os
import json
import datetime
import tempfile
from django.conf import settings
from io import BytesIO
import uuid
from .models import CertificadoGenerado

def generar_certificado_desde_plantilla(datos, qr_path, id_certificado):
    """
    Genera un certificado usando la plantilla Word existente
    """
    # Ruta a la plantilla Word en la carpeta plantillas_word
    plantilla_path = os.path.join(settings.BASE_DIR, 'plantillas_word', 'plantilla_certificado.docx')
    
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"No se encontró la plantilla de certificado Word en {plantilla_path}")
    
    # Crear un archivo temporal para el resultado
    temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
    temp_docx.close()
    
    try:
        # Cargar la plantilla
        doc = DocxTemplate(plantilla_path)
        
        # Crear un RichText para el nombre con Times New Roman
        nombre_rt = RichText()
        nombre_rt.add(datos['nombre'], font='Times New Roman', size=56, bold=True, italic=True)
        
        # Preparar el código QR como imagen inline
        qr_image = InlineImage(doc, qr_path, width=Mm(30), height=Mm(30))
        
        # Contexto para la plantilla
        context = {
            # Usamos RichText para conservar estilo si la plantilla lo admite
            'nombre': nombre_rt,
            'carrera': datos['carrera'],
            'qr_code': qr_image,
            'id_certificado': id_certificado,
            'fecha': datetime.datetime.now().strftime("%d de %B de %Y")
        }
        
        # Renderizar la plantilla
        doc.render(context)
        doc.save(temp_docx.name)
        
        # Convertir a PDF
        pdf_content = convertir_a_pdf(temp_docx.name)
        return pdf_content
        
    finally:
        # Limpiar archivo temporal
        if os.path.exists(temp_docx.name):
            os.unlink(temp_docx.name)

def convertir_a_pdf(docx_path):
    """
    Convierte el documento Word a PDF
    """
    try:
        # Intentar usar docx2pdf si está disponible
        from docx2pdf import convert
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_pdf.close()
        
        try:
            convert(docx_path, temp_pdf.name)
            with open(temp_pdf.name, 'rb') as f:
                pdf_content = f.read()
            return pdf_content
        finally:
            if os.path.exists(temp_pdf.name):
                os.unlink(temp_pdf.name)
                
    except ImportError:
        # Si docx2pdf no está disponible, usar generación PDF directa
        return generar_pdf_directo(docx_path)

def generar_pdf_directo(docx_path):
    """
    Fallback mínimo: extrae el texto del DOCX y lo convierte a PDF.
    """
    buffer = BytesIO()
    pdf_doc = SimpleDocTemplate(buffer, pagesize=A4,
                               rightMargin=72, leftMargin=72,
                               topMargin=72, bottomMargin=18)

    # Estilos
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']

    story = []

    try:
        document = Document(docx_path)
        for paragraph in document.paragraphs:
            text = paragraph.text.strip()
            if text:
                story.append(Paragraph(text, normal_style))
                story.append(Spacer(1, 8))
    except Exception:
        # Si por alguna razón no podemos leer el DOCX, generar un PDF vacío
        story.append(Paragraph("", normal_style))

    pdf_doc.build(story)
    contenido = buffer.getvalue()
    buffer.close()
    return contenido

def generar_qr_optimizado(dni, nombre, carrera, codigo):
    """
    Genera un código QR optimizado con información del certificado
    """
    # Generar ID único
    id_certificado = str(uuid.uuid4())
    fecha_generacion = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Construir URL de verificación usando BASE_URL de settings si está definida
    base_url = getattr(settings, 'BASE_URL', 'http://localhost:8000')
    url_verificacion = f"{base_url}/verificar/{id_certificado}/"
    
    # Datos para el QR
    datos_qr = {
        'id_certificado': id_certificado,
        'dni': dni,
        'nombre': nombre,
        'carrera': carrera,
        'fecha_generacion': fecha_generacion,
        'url': url_verificacion
    }
    
    # Crear código QR
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=20,
        border=4,
    )
    qr.add_data(url_verificacion)
    qr.make(fit=True)
    
    # Generar imagen
    qr_image = qr.make_image(fill_color="black", back_color="white")
    
    # Guardar QR
    qr_dir = os.path.join(settings.MEDIA_ROOT, 'qr')
    os.makedirs(qr_dir, exist_ok=True)
    qr_path = os.path.join(qr_dir, f'qr_{id_certificado}.png')
    qr_image.save(qr_path)
    
    # Guardar en base de datos
    certificado = CertificadoGenerado(
        id_certificado=id_certificado,
        codigo=codigo,
        dni=dni,
        nombre=nombre,
        carrera=carrera,
        ruta_qr=qr_path,
        url_verificacion=url_verificacion
    )
    certificado.save()
    
    return qr_path, id_certificado, url_verificacion

def crear_certificado_completo(datos, formato='pdf'):
    """
    Crear certificado y devolverlo en memoria listo para descargar
    """
    try:
        qr_path, id_certificado, url_verificacion = generar_qr_optimizado(
            datos['dni'],
            datos['nombre'],
            datos['carrera'],
            datos['codigo']
        )

        try:
            contenido = generar_certificado_desde_plantilla(
                datos,
                qr_path,
                id_certificado
            )
            mime_type = 'application/pdf'
            extension = 'pdf'

            # Guardar ruta PDF en base de datos (opcional)
            certificado = CertificadoGenerado.objects.get(id_certificado=id_certificado)
            certificado.ruta_pdf = f'/certificados/certificado_{id_certificado}.{extension}'  # solo referencia
            certificado.save()

            return {
                'contenido': contenido,
                'mime_type': mime_type,
                'nombre_archivo': f'certificado_{datos["nombre"].replace(" ", "_")}.{extension}',
                'id_certificado': id_certificado
            }

        finally:
            if os.path.exists(qr_path):
                os.unlink(qr_path)

    except Exception as e:
        raise Exception(f"Error al crear certificado: {str(e)}")
