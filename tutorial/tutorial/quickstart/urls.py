from django.conf.urls import url
from rest_framework import routers
import views
router = routers.DefaultRouter()
router.register(r'users', views.UserViewSet)
router.register(r'groups', views.GroupViewSet)

urlpatterns = [
    url(r'header_template', views.CustomView1.as_view()),
    url(r'body_template', views.CustomView2.as_view()),
]
 
urlpatterns += router.urls