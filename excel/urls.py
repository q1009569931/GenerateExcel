from django.urls import path
from . import views
from django.views.decorators.csrf import csrf_exempt

urlpatterns = [
	path('', views.GDate.as_view(), name='index'),
	path('download/', csrf_exempt(views.Excel.as_view())),
]