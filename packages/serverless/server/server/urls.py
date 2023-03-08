from django.contrib import admin
from django.urls import path

from serverless.views import activate_license_key, graphql_view, index, send_license_key, get_updated_workbook

admin.autodiscover()


urlpatterns = [
    path("", index, name="index"),
    path("send-license-key", send_license_key),
    path("activate-license-key/<uuid:license_id>/", activate_license_key),
    path("admin/", admin.site.urls),
    path("graphql", graphql_view),
    path("get-updated-workbook/<uuid:workbook_id>/<int:revision>/", get_updated_workbook),
]
