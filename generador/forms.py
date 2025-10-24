from django import forms
from .models import Certificado

class CertificadoForm(forms.ModelForm):
    class Meta:
        model = Certificado
        fields = ['nombre']
        widgets = {
            'nombre': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Ingrese el nombre completo'})
        }