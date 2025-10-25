from django.shortcuts import render, redirect
from django.http import FileResponse, HttpResponse
from django.conf import settings
import os
import tempfile
import pandas as pd
import datetime
from .document_utils import crear_certificado_completo, generar_qr_optimizado


def procesar_plantilla_word_y_generar_pdf(plantilla_path, datos, qr_path, id_certificado):
    """
    Procesa la plantilla de Word usando docxtpl y genera PDF usando reportlab
    Mantiene EXACTAMENTE el diseño de la plantilla original
    """
    try:
        # 1. Procesar la plantilla de Word usando docxtpl (como antes)
        doc = DocxTemplate(plantilla_path)
        
        # Crear un RichText para el nombre con Times New Roman (como antes)
        nombre_rt = RichText()
        nombre_rt.add(datos['nombre'], font='Times New Roman', size=56, bold=True, italic=True)
        
        # Preparar el contexto con el nombre estilizado (como antes)
        context = {
            'nombre': nombre_rt,
            'qr_code': InlineImage(doc, qr_path, width=Mm(30), height=Mm(30)),
            'id_certificado': id_certificado
        }
        
        # Renderizar la plantilla (como antes)
        doc.render(context)
        
        # Guardar temporalmente el documento procesado
        temp_docx = os.path.join(tempfile.gettempdir(), f'certificado_temp_{id_certificado}.docx')
        doc.save(temp_docx)
        
        # 2. Convertir el documento Word procesado a PDF usando reportlab
        # Esto mantiene el diseño exacto de tu plantilla
        pdf_content = convertir_docx_a_pdf_con_plantilla(temp_docx, datos, qr_path)
        
        # Limpiar archivo temporal
        os.remove(temp_docx)
        
        return pdf_content
        
    except Exception as e:
        print(f"Error al procesar plantilla Word: {e}")
        # Fallback: usar la función original si hay error
        return generar_certificado_pdf_multiplataforma(datos['nombre'], datos['carrera'], id_certificado, qr_path)


def convertir_docx_a_pdf_con_plantilla(temp_docx_path, datos, qr_path):
    """
    Convierte el documento Word procesado a PDF manteniendo el diseño exacto de la plantilla
    Usa una alternativa multiplataforma que preserve el formato
    """
    try:
        # Opción 1: Usar python-docx2pdf si está disponible (requiere LibreOffice)
        try:
            from docx2pdf import convert
            import tempfile
            
            # Crear archivo PDF temporal
            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
            temp_pdf.close()
            
            # Convertir usando docx2pdf (requiere LibreOffice en Linux/macOS)
            convert(temp_docx_path, temp_pdf.name)
            
            # Leer el contenido del PDF
            with open(temp_pdf.name, 'rb') as f:
                pdf_content = f.read()
            
            # Limpiar archivo temporal
            os.remove(temp_pdf.name)
            
            return pdf_content
            
        except ImportError:
            print("docx2pdf no disponible, usando alternativa...")
            pass
        except Exception as e:
            print(f"Error con docx2pdf: {e}, usando alternativa...")
            pass
        
        # Opción 2: Usar reportlab para recrear el diseño de la plantilla
        # Esto mantiene el contenido pero con un diseño similar
        return generar_pdf_basado_en_plantilla(temp_docx_path, qr_path)
        
    except Exception as e:
        print(f"Error al convertir DOCX a PDF: {e}")
        # Fallback: generar PDF básico con los datos correctos
        return generar_certificado_pdf_multiplataforma(datos['nombre'], datos['carrera'], datos.get('id_certificado', 'ID'), qr_path)


def generar_pdf_basado_en_plantilla(docx_path, qr_path):
    """
    Genera PDF basado en el contenido de la plantilla Word procesada
    Extrae el contenido exacto incluyendo el nombre insertado
    """
    from docx import Document
    
    # Leer el documento Word procesado (ya contiene el nombre insertado)
    doc = Document(docx_path)
    
    buffer = BytesIO()
    pdf_doc = SimpleDocTemplate(buffer, pagesize=A4, 
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
    
    # Crear estilos
    styles = getSampleStyleSheet()
    
    # Estilo para el título principal
    titulo_style = ParagraphStyle(
        'TituloPersonalizado',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Times-Bold'
    )
    
    # Estilo para el nombre del estudiante (Times New Roman, bold, italic)
    nombre_style = ParagraphStyle(
        'NombrePersonalizado',
        parent=styles['Normal'],
        fontSize=18,
        spaceAfter=20,
        alignment=TA_CENTER,
        fontName='Times-BoldItalic'
    )
    
    # Estilo para el texto del certificado
    texto_style = ParagraphStyle(
        'TextoPersonalizado',
        parent=styles['Normal'],
        fontSize=12,
        spaceAfter=15,
        alignment=TA_LEFT,
        fontName='Times-Roman'
    )
    
    story = []

    # Passthrough: insertar todos los párrafos tal cual (sin textos adicionales)
    normal_style = styles['Normal']
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            story.append(Paragraph(text, normal_style))
            story.append(Spacer(1, 6))

    # Agregar código QR si existe
    if os.path.exists(qr_path):
        try:
            from reportlab.lib.units import inch
            from reportlab.platypus import Image as RLImage
            qr_image = RLImage(qr_path, width=2*inch, height=2*inch)
            story.append(qr_image)
        except Exception:
            # Si no es posible agregar la imagen, se omite
            pass

    # Construir el PDF
    pdf_doc.build(story)

    # Obtener el contenido del buffer
    pdf_content = buffer.getvalue()
    buffer.close()

    return pdf_content


def generar_certificado_pdf_multiplataforma(nombre, carrera, id_certificado, qr_path):
    """
    Genera un certificado PDF usando reportlab (multiplataforma)
    Reemplaza la conversión de Word a PDF que requiere pythoncom
    """
    buffer = BytesIO()

    # Crear el documento PDF
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                          rightMargin=72, leftMargin=72,
                          topMargin=72, bottomMargin=18)

    # Estilos mínimos
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']

    # Contenido mínimo: sólo el nombre y el QR
    story = []
    try:
        from reportlab.lib.enums import TA_CENTER
        nombre_style = ParagraphStyle('NombrePersonalizado', parent=styles['Normal'], fontSize=18, alignment=TA_CENTER)
    except Exception:
        nombre_style = normal_style

    # Nombre del estudiante (destacado)
    story.append(Paragraph(nombre, nombre_style))
    story.append(Spacer(1, 12))

    # Agregar código QR si existe
    if os.path.exists(qr_path):
        try:
            from reportlab.lib.units import inch
            from reportlab.platypus import Image as RLImage
            qr_image = RLImage(qr_path, width=2*inch, height=2*inch)
            story.append(qr_image)
        except Exception:
            pass

    # Construir el PDF
    doc.build(story)

    # Obtener el contenido del buffer
    pdf_content = buffer.getvalue()
    buffer.close()

    return pdf_content


def generar_qr(dni, nombre, carrera, codigo):
    """
    Función legada para mantener compatibilidad.
    Usa la nueva implementación optimizada.
    """
    return generar_qr_optimizado(dni, nombre, carrera, codigo)

def validar_usuario(dni, codigo=None, solo_dni=False):
    # Ruta al archivo Excel
    excel_path = os.path.join(settings.MEDIA_ROOT, 'plantillas', 'BD_CERTIFICADOS.xlsx')
    
    if not os.path.exists(excel_path):
        return False, None
    
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(excel_path)
        
        # Convertir DNI y CODIGO a string para comparación
        df['DNI'] = df['DNI'].astype(str)
        df['CODIGO'] = df['CODIGO'].astype(str)
        
        # Buscar el usuario en el DataFrame
        if solo_dni:
            # Si solo_dni es True, buscar solo por DNI
            resultado = df[df['DNI'] == str(dni)]
        else:
            # Si no, buscar por DNI y CODIGO
            resultado = df[(df['DNI'] == str(dni)) & (df['CODIGO'] == str(codigo))]
        
        if not resultado.empty:
            datos = {
                'dni': resultado['DNI'].values[0],
                'nombre': resultado['NOMBRES'].values[0],
                'carrera': resultado['CARRERA'].values[0],
                'codigo': resultado['CODIGO'].values[0]
            }
            return True, datos
        else:
            return False, None
    except Exception as e:
        print(f"Error al validar usuario: {e}")
        return False, None

def index(request):
    mensaje_error_login = None
    
    # Si el usuario ya está autenticado, mostrar directamente la página de confirmación
    if request.session.get('autenticado'):
        if request.session.get('es_admin'):
            return redirect('opciones_admin')
        else:
            dni = request.session.get('dni_validado')
            valido, datos = validar_usuario(dni, None)
            if valido:
                return render(request, 'generador/confirmacion.html', {
                    'nombre': datos['nombre'],
                    'carrera': datos['carrera']
                })
            else:
                request.session.flush()
    
    if request.method == 'POST':
        form_type = request.POST.get('form_type')
        
        if form_type == 'login':
            username = request.POST.get('username')
            password = request.POST.get('password')
            
            # Verificar si es el administrador
            if username == 'Upla_123' and password == 'Upla321':
                request.session['autenticado'] = True
                request.session['es_admin'] = True
                return redirect('admin_panel')
            else:
                # Verificar usuario normal
                valido, datos = validar_usuario(username, password)
                if valido:
                    request.session['autenticado'] = True
                    request.session['es_admin'] = False
                    request.session['dni_validado'] = datos['dni']
                    request.session.save()  # Forzar el guardado de la sesión
                    return render(request, 'generador/confirmacion.html', {
                        'nombre': datos['nombre'],
                        'carrera': datos['carrera']
                    })
                else:
                    mensaje_error_login = "Usuario o contraseña incorrectos"
                
        elif form_type == 'logout':
            # Cerrar sesión
            from django.contrib.auth import logout
            logout(request)  # Método oficial de Django para cerrar sesión
            
            # Asegurar que se eliminen todas las cookies de sesión
            response = redirect('index')
            response.delete_cookie('sessionid')
            response.delete_cookie('csrftoken')
            
            # Establecer max-age y expires para forzar la eliminación
            response.set_cookie('sessionid', '', max_age=0, expires='Thu, 01 Jan 1970 00:00:00 GMT')
            
            return response
    
    return render(request, 'generador/index.html', {
        'mensaje_error_login': mensaje_error_login
    })

def opciones_admin(request):
    if not request.session.get('autenticado') or not request.session.get('es_admin'):
        return redirect('index')
    
    mensaje_error = None
    mensaje_exito = None
    
    if request.method == 'POST':
        if 'excel_file' not in request.FILES:
            mensaje_error = 'Por favor, seleccione un archivo Excel.'
        else:
            try:
                # Guardar el archivo Excel
                excel_file = request.FILES['excel_file']
                excel_path = os.path.join(settings.MEDIA_ROOT, 'plantillas', 'BD_CERTIFICADOS.xlsx')
                
                with open(excel_path, 'wb+') as destination:
                    for chunk in excel_file.chunks():
                        destination.write(chunk)
                
                # Cargar los datos del Excel a la base de datos
                from .models import CertificadoGenerado
                df = pd.read_excel(excel_path)
                
                # Convertir columnas a string para evitar problemas
                df['DNI'] = df['DNI'].astype(str)
                df['CODIGO'] = df['CODIGO'].astype(str)
                
                # Importar datos a la base de datos
                registros_importados = 0
                for _, row in df.iterrows():
                    # Verificar si ya existe un registro con el mismo DNI y CODIGO
                    if not CertificadoGenerado.objects.filter(dni=row['DNI'], codigo=row['CODIGO']).exists():
                        # Generar un ID único para cada registro
                        id_certificado = str(uuid.uuid4())
                        # Crear URL de verificación
                        url_verificacion = f"http://localhost:8000/verificar/{id_certificado}/"
                        
                        # Crear registro en la base de datos
                        CertificadoGenerado.objects.create(
                            id_certificado=id_certificado,
                            codigo=row['CODIGO'],
                            dni=row['DNI'],
                            nombre=row['NOMBRES'],
                            carrera=row['CARRERA'],
                            url_verificacion=url_verificacion
                        )
                        registros_importados += 1
                
                mensaje_exito = f'Archivo cargado exitosamente. {registros_importados} registros importados a la base de datos.'
            except Exception as e:
                mensaje_error = f'Error al procesar el archivo: {str(e)}'
    
    return render(request, 'generador/opciones_admin.html', {
        'mensaje_error': mensaje_error,
        'mensaje_exito': mensaje_exito
    })

def listar_certificados(request):
    if not request.session.get('autenticado') or not request.session.get('es_admin'):
        return redirect('index')
    
    from .models import CertificadoGenerado
    from django.core.paginator import Paginator
    
    # Obtener todos los certificados ordenados por fecha de generación (más reciente primero)
    certificados = CertificadoGenerado.objects.all().order_by('-fecha_generacion')
    
    # Configurar la paginación (10 por página)
    paginator = Paginator(certificados, 10)
    page_number = request.GET.get('page', 1)
    page_obj = paginator.get_page(page_number)
    
    return render(request, 'generador/listar_certificados.html', {
        'page_obj': page_obj,
    })

def admin_panel(request):
    # Redirigir a opciones_admin para evitar conflictos
    return redirect('opciones_admin')

def descargar_plantilla(request):
    """
    Vista optimizada para descargar certificados
    """
    if not request.session.get('autenticado'):
        return redirect('index')
    
    dni = request.session.get('dni_validado')
    if not dni:
        return redirect('index')
    
    # Validar usuario
    valido, datos = validar_usuario(dni, solo_dni=True)
    if not valido or not datos:
        return redirect('index')
    
    try:
        # Usar la nueva función optimizada para crear el certificado
        resultado = crear_certificado_completo(datos, formato='pdf')
        
        # Preparar respuesta
        response = HttpResponse(
            resultado['contenido'],
            content_type=resultado['mime_type']
        )
        response['Content-Disposition'] = f'attachment; filename="{resultado["nombre_archivo"]}"'
        
        return response
            
    except Exception as e:
        return render(request, 'generador/error.html', {
            'error': f'Error al generar el certificado: {str(e)}'
        })
            

def verificar_certificado(request, id_certificado):
    """
    Vista para verificar la autenticidad de un certificado mediante su ID único.
    Esta vista se accede al escanear el código QR.
    """
    from .models import CertificadoGenerado
    
    # Buscar el certificado en la base de datos
    certificado = CertificadoGenerado.objects.filter(id_certificado=id_certificado).first()
    
    if certificado:
        # Si el certificado existe, mostrar la información
        context = {
            'certificado': certificado,
            'valido': True,
            'mensaje': 'Certificado válido',
        }
    else:
        # Si el certificado no existe, mostrar mensaje de error
        context = {
            'valido': False,
            'mensaje': 'El certificado no es válido o no existe en nuestros registros.',
        }
    
    return render(request, 'generador/verificar.html', context)

def generar_lote(request):
    """
    Vista optimizada para generar lotes de certificados
    """
    if request.method != 'POST':
        return redirect('index')
    
    if 'excel_file' not in request.FILES:
        return render(request, 'generador/admin.html', {
            'error': 'Por favor, seleccione un archivo Excel.'
        })
    
    excel_file = request.FILES['excel_file']
    cantidad = int(request.POST.get('cantidad', 0))
    
    if cantidad <= 0:
        return render(request, 'generador/admin.html', {
            'error': 'Por favor, ingrese una cantidad válida.'
        })
    
    try:
        # Leer el archivo Excel
        df = pd.read_excel(excel_file)
        df = df.head(cantidad)
        
        # Lista para almacenar los archivos temporales
        temp_files = []
        
        # Generar certificados
        for _, row in df.iterrows():
            datos = {
                'dni': str(row['DNI']),
                'nombre': row['NOMBRES'],
                'carrera': row['CARRERA'],
                'codigo': str(row['CODIGO'])
            }
            
            # Usar la función optimizada para crear el certificado
            try:
                resultado = crear_certificado_completo(datos, formato='pdf')
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                temp_file.write(resultado['contenido'])
                temp_file.close()
                temp_files.append((temp_file.name, datos['nombre'], 'pdf'))
            except Exception as e:
                print(f"Error al generar certificado para {datos['nombre']}: {e}")
                continue
        
        # Crear ZIP en memoria
        from io import BytesIO
        import zipfile
        
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for temp_file, nombre, extension in temp_files:
                try:
                    with open(temp_file, 'rb') as f:
                        zip_file.writestr(f'certificado_{nombre}.{extension}', f.read())
                finally:
                    # Limpiar archivos temporales
                    if os.path.exists(temp_file):
                        os.unlink(temp_file)
        
        # Preparar respuesta
        zip_content = zip_buffer.getvalue()
        zip_buffer.close()
        
        response = HttpResponse(zip_content, content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename="certificados_lote.zip"'
        response['Content-Length'] = len(zip_content)
        response['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response['Pragma'] = 'no-cache'
        response['Expires'] = '0'
        
        return response
        
    except Exception as e:
        return render(request, 'generador/admin.html', {
            'error': f'Error al procesar el archivo: {str(e)}'
        })
