from django.conf.urls import url
from rest_framework import routers
from . import views
router = routers.DefaultRouter()

urlpatterns = [
    url(r'body_template', views.CustomView.as_view()),
    url(r'invoice_template', views.InvoiceView.as_view()),
    url(r'quotation_template', views.QuotationView.as_view()),
]
 
urlpatterns += router.urls