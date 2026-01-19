from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('processar/', views.processar_arquivo, name='processar'),
    path('download/', views.download_resultado, name='download'),
]
