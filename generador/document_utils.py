"""
Utilidades optimizadas para la generación de certificados usando plantillas Word
"""
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm, Pt
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT
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
    # Ruta a la plantilla Word
    plantilla_path = os.path.join(settings.MEDIA_ROOT, 'plantillas', 'plantilla_certificado.docx')
    
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError("No se encontró la plantilla de certificado Word.")
    
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
            'nombre': datos['nombre'],  # Usamos el nombre directamente como está en la plantilla
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
    Genera un PDF directamente usando ReportLab como fallback
    """
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, 
                          rightMargin=72, leftMargin=72,
                          topMargin=72, bottomMargin=18)
    
    # Estilos personalizados
    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle(
        'TituloPersonalizado',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Times-Bold'
    )
    
    nombre_style = ParagraphStyle(
        'NombrePersonalizado',
        parent=styles['Normal'],
        fontSize=18,
        spaceAfter=20,
        alignment=TA_CENTER,
        fontName='Times-BoldItalic'
    )
    
    texto_style = ParagraphStyle(
        'TextoPersonalizado',
        parent=styles['Normal'],
        fontSize=12,
        spaceAfter=15,
        alignment=TA_LEFT,
        fontName='Times-Roman'
    )
    
    # Contenido del certificado
    story = []
    
    # Título
    story.append(Paragraph("CERTIFICADO DE ESTUDIOS", titulo_style))
    story.append(Spacer(1, 20))
    
    # Texto introductorio
    story.append(Paragraph(
        "Por medio del presente certificado, se hace constar que el estudiante:",
        texto_style
    ))
    story.append(Spacer(1, 10))
    
    # Nombre del estudiante
    story.append(Paragraph(nombre, nombre_style))
    story.append(Spacer(1, 20))
    
    # Información adicional
    story.append(Paragraph(
        f"Ha completado satisfactoriamente sus estudios en la carrera de "
        f"<b>{carrera}</b> en la Universidad Peruana Los Andes (UPLA).",
        texto_style
    ))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph(
        "Este certificado es válido y puede ser verificado mediante el código QR "
        "adjunto o visitando nuestro sistema de verificación.",
        texto_style
    ))
    story.append(Spacer(1, 10))
    
    # Información de verificación
    story.append(Paragraph(
        f"ID de Certificado: {id_certificado}<br/>"
        f"Fecha de emisión: {datetime.datetime.now().strftime('%d de %B de %Y')}",
        texto_style
    ))
    story.append(Spacer(1, 30))
    
    # Código QR
    if os.path.exists(qr_path):
        qr_image = Image(qr_path, width=2*inch, height=2*inch)
        story.append(qr_image)
    
    # Generar PDF
    doc.build(story)
    return buffer.getvalue()

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
    Función principal para crear certificados usando la plantilla Word
    """
    try:
        # 1. Generar QR
        qr_path, id_certificado, url_verificacion = generar_qr_optimizado(
            datos['dni'],
            datos['nombre'],
            datos['carrera'],
            datos['codigo']
        )
        
        try:
            # 2. Generar certificado usando la plantilla
            contenido = generar_certificado_desde_plantilla(
                datos,
                qr_path,
                id_certificado
            )
            mime_type = 'application/pdf'
            extension = 'pdf'
            
            # 3. Guardar el archivo en media
            media_dir = os.path.join(settings.MEDIA_ROOT, 'certificados')
            os.makedirs(media_dir, exist_ok=True)
            archivo_path = os.path.join(media_dir, f'certificado_{id_certificado}.{extension}')
            
            with open(archivo_path, 'wb') as f:
                f.write(contenido)
            
            # 4. Actualizar registro en la base de datos
            certificado = CertificadoGenerado.objects.get(id_certificado=id_certificado)
            certificado.ruta_pdf = f'/media/certificados/certificado_{id_certificado}.{extension}'
            certificado.save()
            
            return {
                'contenido': contenido,
                'mime_type': mime_type,
                'nombre_archivo': f'certificado_{datos["nombre"].replace(" ", "_")}.{extension}',
                'id_certificado': id_certificado
            }
            
        finally:
            # Limpiar QR temporal
            if os.path.exists(qr_path):
                os.unlink(qr_path)
                
    except Exception as e:
        raise Exception(f"Error al crear certificado: {str(e)}")