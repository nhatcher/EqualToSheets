from django.contrib import admin
from django.urls import path, re_path

from serverless.views import activate_license_key, graphql_view, index, send_license_key

admin.autodiscover()


urlpatterns = [
    path("", index, name="index"),
    path("send-license-key", send_license_key),
    re_path(r"^activate-license-key/(?P<license_id>[\-a-z0-9]{0,50})/$", activate_license_key),
    path("admin/", admin.site.urls),
    path("graphql", graphql_view),
]
