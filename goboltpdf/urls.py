from django.urls import path
from . import views

urlpatterns = [
    path('read-pdf/', views.read_pdf, name='read_pdf'),
]
