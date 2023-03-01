from django.contrib import admin
from django.urls import include, path, re_path

from serverless.views import graphql_view

admin.autodiscover()

import serverless.views

# To add a new path, first import the app:
# import blog
#
# Then add the new path:
# path('blog/', blog.urls, name="blog")
#
# Learn more here: https://docs.djangoproject.com/en/2.1/topics/http/urls/

urlpatterns = [
    path("", serverless.views.index, name="index"),
    path("send-license-key", serverless.views.send_license_key),
    re_path(r'^activate-license-key/(?P<license_id>[\-a-z0-9]{0,50})/$', serverless.views.activate_license_key),
    path("admin/", admin.site.urls),
    path("graphql", graphql_view),
]


