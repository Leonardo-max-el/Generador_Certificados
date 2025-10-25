from django.db import models

class Certificado(models.Model):
    nombre = models.CharField(max_length=200)
    fecha_creacion = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return self.nombre

class CertificadoGenerado(models.Model):
    id_certificado = models.CharField(max_length=100, unique=True)
    codigo = models.CharField(max_length=50)
    dni = models.CharField(max_length=20)
    nombre = models.CharField(max_length=200)
    carrera = models.CharField(max_length=200)
    fecha_generacion = models.DateTimeField(auto_now_add=True)
    ruta_pdf = models.CharField(max_length=255, blank=True, null=True)
    ruta_qr = models.CharField(max_length=255, blank=True, null=True)
    url_verificacion = models.CharField(max_length=255, blank=True, null=True)
    
    def __str__(self):
        return f"{self.nombre} - {self.id_certificado}"
