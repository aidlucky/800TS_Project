"""scheduling URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/1.11/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  url(r'^$', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  url(r'^$', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.conf.urls import url, include
    2. Add a URL to urlpatterns:  url(r'^blog/', include('blog.urls'))
"""
from django.conf.urls import url
from django.contrib import admin
from .view_scheduling import *
from .view_working_time import *
from django.conf import settings  
from django.conf.urls.static import static

urlpatterns = [
    # url(r'^admin/', admin.site.urls),
    url(r'^get_template/$',get_template,name="get_template"),
    url(r'^home/$',sehceduing,name="sehceduing"), 
    url(r'^working_time/$',working_time,name="working_time"),
    url(r'^get_working_time_template/$',get_working_time_template,name="get_working_time_template"),
]+ static(settings.STATIC_URL, document_root = settings.STATIC_ROOT)
