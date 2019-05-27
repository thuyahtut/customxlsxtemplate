from django.conf.urls import url,include
# from django.urls import path
from rest_framework import routers
from tutorial.quickstart import views

router = routers.DefaultRouter()
# router.register(r'users', views.UserViewSet)
# router.register(r'groups', views.GroupViewSet)
# router.register(r'exportformat', views.ExportFormatSet)

# Wire up our API using automatic URL routing.
# Additionally, we include login URLs for the browsable API.
urlpatterns = [
    url('', include(router.urls)),
    url('api-auth/', include('rest_framework.urls', namespace='rest_framework')),
    url(r'^api/', include('tutorial.quickstart.urls'))
]
