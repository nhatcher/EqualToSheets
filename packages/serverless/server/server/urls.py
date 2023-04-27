from django.contrib import admin
from django.urls import include, path
from drf_spectacular.views import SpectacularAPIView, SpectacularSwaggerView

from serverless.views import (
    activate_license_key,
    create_workbook_from_xlsx,
    edit_workbook,
    get_updated_workbook,
    graphql_view,
    hacky_send_email_to_subscriber,
    redirect_to_gitbook_docs,
    send_license_key,
    simulate,
    unsubscribe_email,
)

admin.autodiscover()


urlpatterns = [
    path("send-license-key", send_license_key),
    path("create-workbook-from-xlsx", create_workbook_from_xlsx),
    path("activate-license-key/<uuid:license_id>", activate_license_key),
    path("admin/", admin.site.urls),
    path("graphql", graphql_view),
    path("get-updated-workbook/<uuid:workbook_id>/<int:revision>", get_updated_workbook),
    path("unsafe-just-for-beta/edit-workbook/<uuid:license_key>/<uuid:workbook_id>/", edit_workbook),
    path("unsubscribe-email", unsubscribe_email),
    # TODO: Move the simulate endpoint implementation to serverless.rest_api
    path("api/v1/workbooks/<uuid:workbook_id>/simulate", simulate),
    path("api/", include("serverless.rest_api.urls")),
    path("sso/", include("serverless.sso.urls")),
    path("hacky_send_email_to_subscriber/<str:email>", hacky_send_email_to_subscriber),
    path("schema/", SpectacularAPIView.as_view(), name="schema"),
    path(
        "docs/",
        SpectacularSwaggerView.as_view(url_name="schema"),
        name="docs",
    ),
    path("beta-readme", redirect_to_gitbook_docs),
]
