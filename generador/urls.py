from django.urls import path
from . import views
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path('', views.index, name='index'),
    path('descargar/', views.descargar_plantilla, name='descargar_plantilla'),
    path('generar_lote/', views.generar_lote, name='generar_lote'),
    path('admin_panel/', views.admin_panel, name='admin_panel'),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)