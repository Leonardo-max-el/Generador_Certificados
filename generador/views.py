from django.shortcuts import render, redirect
from django.http import FileResponse, HttpResponse
from django.conf import settings
import os
import tempfile
import pandas as pd
import qrcode
from docxtpl import DocxTemplate, InlineImage, RichText
import uuid
import json
import datetime,time
import pandas as pd
from docx import Document
from django.conf import settings
from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from docx.shared import Mm, Pt, RGBColor

from docx2pdf import convert
import pythoncom


def generar_qr(dni, nombre, carrera, codigo):
    # Generar ID único para el certificado
    id_certificado = str(uuid.uuid4())
    fecha_generacion = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Apuntar el QR directamente al PDF público del certificado
    url_verificacion = f"http://10.86.231.63:8000/media/certificados/certificado_{id_certificado}.pdf"
    
    # Crear el contenido del QR (ahora con la URL)
    datos_qr = {
        'id_certificado': id_certificado,
        'dni': dni,
        'nombre': nombre,
        'carrera': carrera,
        'fecha_generacion': fecha_generacion,
        'url': url_verificacion
    }
    
    # Convertir a JSON
    contenido_qr = json.dumps(datos_qr)
    
    # Crear el código QR que apunta directamente a la URL
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=20,
        border=4,
    )
    # Usar directamente la URL para el QR
    qr.add_data(url_verificacion)
    qr.make(fit=True)

    # Crear la imagen del QR en negro
    qr_image = qr.make_image(fill_color="black", back_color="white")
    
    # Asegurar que existe el directorio para QR
    qr_dir = os.path.join(settings.MEDIA_ROOT, 'qr')
    os.makedirs(qr_dir, exist_ok=True)
    
    # Guardar el QR
    qr_path = os.path.join(qr_dir, f'qr_{id_certificado}.png')
    qr_image.save(qr_path)
    
    # Guardar en la base de datos
    from .models import CertificadoGenerado
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
    print("Iniciando proceso de descarga...")
    print("Estado de la sesión:", request.session.items())
    
    # Verificar si el usuario está autenticado
    if not request.session.get('autenticado'):
        print("Usuario no autenticado")
        print("Contenido de la sesión:", dict(request.session))
        return redirect('index')
    
    # Obtener el DNI de la sesión
    dni = request.session.get('dni_validado')
    if not dni:
        print("DNI no encontrado en la sesión")
        print("Contenido de la sesión:", dict(request.session))
        return redirect('index')
    
    print(f"DNI encontrado en la sesión: {dni}")
    
    # Validar el DNI y obtener los datos (solo validación por DNI ya que está autenticado)
    valido, datos = validar_usuario(dni, solo_dni=True)
    if not valido or not datos:
        print("DNI no válido o datos no encontrados")
        print(f"Resultado de validación: valido={valido}, datos={datos}")
        return redirect('index')
    
    print(f"Datos del usuario validados: {datos}")
    
    qr_path = None
    temp_docx = None
    temp_pdf = None
    
    try:
        print("Generando código QR...")
        # 1. Generar el código QR
        qr_path, id_certificado, url_verificacion = generar_qr(datos['dni'], datos['nombre'], datos['carrera'], datos['codigo'])
        
        # Ruta a la plantilla Word
        plantilla_path = os.path.join(settings.MEDIA_ROOT, 'plantillas', 'plantilla_certificado.docx')
        
        if not os.path.exists(plantilla_path):
            print("Plantilla no encontrada")
            return render(request, 'generador/error.html', {
                'error': 'No se encontró la plantilla de certificado Word.'
            })
        
        print("Creando archivos temporales...")
        # 2. Crear archivo temporal para Word y PDF
        temp_docx = os.path.join(tempfile.gettempdir(), f'certificado_{datos["dni"]}.docx')
        temp_pdf = os.path.join(tempfile.gettempdir(), f'certificado_{datos["dni"]}.pdf')
        
        # 3. Generar el certificado Word
        print("Generando certificado Word...")
        doc = DocxTemplate(plantilla_path)
        
        # Crear un RichText para el nombre con Times New Roman
        nombre_rt = RichText()
        nombre_rt.add(datos['nombre'], font='Times New Roman', size=56, bold=True, italic=True)
        
        # Preparar el contexto con el nombre estilizado
        context = {
            'nombre': nombre_rt,
            'qr_code': InlineImage(doc, qr_path, width=Mm(30), height=Mm(30)),
            'id_certificado': id_certificado
        }
        
        print("Contexto preparado:", context)
        
        # Renderizar la plantilla
        doc.render(context)
        doc.save(temp_docx)
        
        print("Certificado Word generado, convirtiendo a PDF...")
        # 4. Convertir a PDF
        pythoncom.CoInitialize()
        convert(temp_docx, temp_pdf)
        pythoncom.CoUninitialize()
        
        print("PDF generado, guardando en la base de datos...")
        
        # Guardar la ruta del PDF en la base de datos
        from .models import CertificadoGenerado
        certificado = CertificadoGenerado.objects.filter(id_certificado=id_certificado).first()
        if certificado:
            # Crear una copia permanente del PDF en el directorio de medios
            pdf_dir = os.path.join(settings.MEDIA_ROOT, 'certificados')
            os.makedirs(pdf_dir, exist_ok=True)
            pdf_path = os.path.join(pdf_dir, f'certificado_{id_certificado}.pdf')
            
            # Copiar el PDF temporal al directorio de medios
            with open(temp_pdf, 'rb') as src, open(pdf_path, 'wb') as dst:
                dst.write(src.read())
            
            # Actualizar el registro en la base de datos (usar URL servible)
            pdf_url = f"http://10.86.231.63:8000/media/certificados/certificado_{id_certificado}.pdf"
            certificado.ruta_pdf = pdf_url
            certificado.save()
        
        print("PDF guardado (URL) en la base de datos, preparando respuesta...")
        # 5. Preparar la respuesta con el PDF
        with open(temp_pdf, 'rb') as pdf_file:
            # Leer el contenido del PDF en memoria
            pdf_content = pdf_file.read()
            
            # Crear la respuesta HTTP con el contenido del PDF
            response = HttpResponse(pdf_content, content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="certificado_{datos["nombre"].replace(" ", "_")}.pdf"'
            
            print("Respuesta preparada, enviando archivo...")
            return response
            
    except Exception as e:
        print(f"Error durante la generación del certificado: {e}")
        return render(request, 'generador/error.html', {
            'error': f'Error al generar el certificado: {str(e)}'
        })
        
    finally:
        print("Limpiando archivos temporales...")
        # 6. Limpiar archivos temporales
        try:
            if qr_path and os.path.exists(qr_path):
                os.remove(qr_path)
                print(f"QR eliminado: {qr_path}")
        except Exception as e:
            print(f"Error al eliminar QR: {e}")
            
        try:
            if temp_docx and os.path.exists(temp_docx):
                os.remove(temp_docx)
                print(f"DOCX temporal eliminado: {temp_docx}")
        except Exception as e:
            print(f"Error al eliminar DOCX temporal: {e}")
            
        try:
            if temp_pdf and os.path.exists(temp_pdf):
                os.remove(temp_pdf)
                print(f"PDF temporal eliminado: {temp_pdf}")
        except Exception as e:
            print(f"Error al eliminar PDF temporal: {e}")

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
    if request.method == 'POST':
        # Verificar si se subió un archivo
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
            
            # Limitar a la cantidad especificada
            df = df.head(cantidad)
            
            # Lista para almacenar los archivos temporales
            temp_files = []
            
            # Importar InlineImage y Mm
            from docxtpl import InlineImage
            from docx.shared import Mm
            
            # Generar certificados para cada registro
            for _, row in df.iterrows():
                datos = {
                    'dni': str(row['DNI']),
                    'nombre': row['NOMBRES'],
                    'carrera': row['CARRERA']
                }
                
                # Generar QR
                qr_path, id_certificado, url_verificacion = generar_qr(datos['dni'], datos['nombre'], datos['carrera'], str(row['CODIGO']))
                
                # Crear archivos temporales para Word y PDF
                temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
                temp_docx.close()
                temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
                temp_pdf.close()
                
                # Generar el certificado
                doc = DocxTemplate(os.path.join(settings.MEDIA_ROOT, 'plantillas', 'plantilla_certificado.docx'))
                
                # Crear la imagen inline con un tamaño de 30mm x 30mm
                qr_image = InlineImage(doc, qr_path, width=Mm(30))
                
                # Preparar el contexto con la imagen QR
                context = {
                    'nombre': datos['nombre'],
                    'carrera': datos['carrera'],
                    'id_certificado': id_certificado,
                    'qr_code': qr_image
                }
                
                # Renderizar el documento
                doc.render(context)
                
                # Guardar el documento Word
                doc.save(temp_docx.name)
                
                # Inicializar COM y convertir a PDF
                pythoncom.CoInitialize()
                try:
                    convert(temp_docx.name, temp_pdf.name)
                except Exception as e:
                    print(f"Error al convertir a PDF: {str(e)}")
                    # Limpiar archivos temporales antes de salir
                    os.unlink(temp_docx.name)
                    os.unlink(qr_path)
                    return render(request, 'generador/admin.html', {
                        'error': f'Error al generar el PDF para {datos["nombre"]}. Por favor, intente nuevamente.'
                    })
                finally:
                    pythoncom.CoUninitialize()
                
                # Verificar que el archivo PDF se haya creado y esperar si es necesario
                import time
                max_attempts = 10
                attempt = 0
                while attempt < max_attempts:
                    if os.path.exists(temp_pdf.name) and os.path.getsize(temp_pdf.name) > 0:
                        break
                    time.sleep(1)  # Esperar 1 segundo antes de verificar nuevamente
                    attempt += 1
                
                if attempt >= max_attempts:
                    # Limpiar archivos temporales antes de salir
                    os.unlink(temp_docx.name)
                    os.unlink(qr_path)
                    for temp_file, _ in temp_files:
                        if os.path.exists(temp_file):
                            os.unlink(temp_file)
                    return render(request, 'generador/admin.html', {
                        'error': f'Error al generar el PDF para {datos["nombre"]}. Por favor, intente nuevamente.'
                    })
                
                # Agregar el PDF y Word a la lista de archivos temporales
                temp_files.append((temp_pdf.name, datos['nombre'], 'pdf'))
                temp_files.append((temp_docx.name, datos['nombre'], 'docx'))
                
                # Eliminar el archivo QR temporal
                os.unlink(qr_path)
            
            # Crear un archivo ZIP con todos los documentos
            zip_path = os.path.join(settings.MEDIA_ROOT, 'certificados_lote.zip')
            import zipfile
            with zipfile.ZipFile(zip_path, 'w') as zip_file:
                for temp_file, nombre, extension in temp_files:
                    zip_file.write(temp_file, f'certificado_{nombre}.{extension}')
                    os.unlink(temp_file)  # Eliminar el archivo temporal
            
            try:
                # Crear un archivo ZIP en memoria
                from io import BytesIO
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                    for temp_file, nombre, extension in temp_files:
                        # Leer el contenido del archivo temporal
                        with open(temp_file, 'rb') as f:
                            zip_file.writestr(f'certificado_{nombre}.{extension}', f.read())
                        # Eliminar el archivo temporal después de agregarlo al ZIP
                        os.unlink(temp_file)
                
                # Obtener el contenido del ZIP en memoria
                zip_content = zip_buffer.getvalue()
                zip_buffer.close()
                
                # Crear la respuesta con el contenido en memoria
                response = HttpResponse(zip_content, content_type='application/zip')
                response['Content-Disposition'] = 'attachment; filename="certificados_lote.zip"'
                response['Content-Length'] = len(zip_content)
                # Agregar headers para evitar el caché
                response['Cache-Control'] = 'no-cache, no-store, must-revalidate'
                response['Pragma'] = 'no-cache'
                response['Expires'] = '0'
                
                return response
                
            except Exception as e:
                # Si ocurre algún error, intentar limpiar el archivo ZIP
                try:
                    if os.path.exists(zip_path):
                        os.unlink(zip_path)
                except:
                    pass
                
                return render(request, 'generador/admin.html', {
                    'error': f'Error al descargar el archivo ZIP: {str(e)}'
                })
            
        except Exception as e:
            return render(request, 'generador/admin.html', {
                'error': f'Error al procesar el archivo: {str(e)}'
            })
    
    return redirect('index')
