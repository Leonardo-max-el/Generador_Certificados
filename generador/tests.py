from django.test import TestCase, Client
from django.core.files.uploadedfile import SimpleUploadedFile
from django.conf import settings
import os
import tempfile
import pandas as pd
from .models import CertificadoGenerado
from .document_utils import (
    generar_certificado_pdf,
    generar_qr_optimizado,
    crear_certificado_completo
)

class DocumentUtilsTests(TestCase):
    def setUp(self):
        self.datos_prueba = {
            'dni': '12345678',
            'nombre': 'Usuario Prueba',
            'carrera': 'Carrera Prueba',
            'codigo': 'COD123'
        }
        
    def test_generar_qr_optimizado(self):
        """
        Prueba la generación optimizada de códigos QR
        """
        qr_path, id_certificado, url_verificacion = generar_qr_optimizado(
            self.datos_prueba['dni'],
            self.datos_prueba['nombre'],
            self.datos_prueba['carrera'],
            self.datos_prueba['codigo']
        )
        
        self.assertTrue(os.path.exists(qr_path))
        self.assertIsNotNone(id_certificado)
        self.assertTrue(url_verificacion.startswith(settings.BASE_URL))
        
        # Limpiar
        if os.path.exists(qr_path):
            os.unlink(qr_path)
    
    def test_generar_certificado_pdf(self):
        """
        Prueba la generación de certificados PDF
        """
        # Generar QR temporal para la prueba
        qr_path, id_certificado, _ = generar_qr_optimizado(
            self.datos_prueba['dni'],
            self.datos_prueba['nombre'],
            self.datos_prueba['carrera'],
            self.datos_prueba['codigo']
        )
        
        try:
            # Generar PDF
            pdf_content = generar_certificado_pdf(
                self.datos_prueba['nombre'],
                self.datos_prueba['carrera'],
                id_certificado,
                qr_path
            )
            
            self.assertIsNotNone(pdf_content)
            self.assertGreater(len(pdf_content), 0)
            
            # Verificar que es un PDF válido
            self.assertTrue(pdf_content.startswith(b'%PDF-'))
            
        finally:
            # Limpiar
            if os.path.exists(qr_path):
                os.unlink(qr_path)
    
    def test_crear_certificado_completo(self):
        """
        Prueba la creación completa de un certificado
        """
        resultado = crear_certificado_completo(self.datos_prueba, formato='pdf')
        
        self.assertIsNotNone(resultado)
        self.assertIn('contenido', resultado)
        self.assertIn('mime_type', resultado)
        self.assertIn('nombre_archivo', resultado)
        self.assertIn('id_certificado', resultado)
        
        # Verificar que se creó el registro en la base de datos
        certificado = CertificadoGenerado.objects.get(id_certificado=resultado['id_certificado'])
        self.assertEqual(certificado.nombre, self.datos_prueba['nombre'])
        self.assertEqual(certificado.dni, self.datos_prueba['dni'])

class ViewsTests(TestCase):
    def setUp(self):
        self.client = Client()
        self.datos_prueba = {
            'dni': '12345678',
            'nombre': 'Usuario Prueba',
            'carrera': 'Carrera Prueba',
            'codigo': 'COD123'
        }
        
        # Crear archivo Excel de prueba
        self.df = pd.DataFrame([self.datos_prueba])
        self.excel_file = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
        self.df.to_excel(self.excel_file.name, index=False)
    
    def tearDown(self):
        # Limpiar archivos temporales
        if os.path.exists(self.excel_file.name):
            os.unlink(self.excel_file.name)
    
    def test_descargar_plantilla(self):
        """
        Prueba la descarga de un certificado individual
        """
        # Simular sesión autenticada
        session = self.client.session
        session['autenticado'] = True
        session['dni_validado'] = self.datos_prueba['dni']
        session.save()
        
        # Preparar datos en la base de datos
        with open(self.excel_file.name, 'rb') as excel:
            self.client.post('/opciones_admin/', {
                'excel_file': SimpleUploadedFile(
                    'test.xlsx',
                    excel.read(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            })
        
        # Probar descarga
        response = self.client.get('/descargar_plantilla/')
        
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'], 'application/pdf')
    
    def test_generar_lote(self):
        """
        Prueba la generación de certificados en lote
        """
        with open(self.excel_file.name, 'rb') as excel:
            response = self.client.post('/generar_lote/', {
                'excel_file': SimpleUploadedFile(
                    'test.xlsx',
                    excel.read(),
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                ),
                'cantidad': 1
            })
        
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'], 'application/zip')
        self.assertIn('certificados_lote.zip', response['Content-Disposition'])
