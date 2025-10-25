from django.urls import path
from . import views
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('', views.index, name='index'),
    path('opciones_admin/', views.opciones_admin, name='opciones_admin'),
    path('admin_panel/', views.admin_panel, name='admin_panel'),
    path('descargar_plantilla/', views.descargar_plantilla, name='descargar_plantilla'),
    path('generar_lote/', views.generar_lote, name='generar_lote'),
    path('verificar/<str:id_certificado>/', views.verificar_certificado, name='verificar_certificado'),
    path('listar_certificados/', views.listar_certificados, name='listar_certificados'),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)