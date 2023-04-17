from django.urls import path, register_converter

from serverless.sso.converters import SSOIntegrationConverter
from serverless.sso.views import sso_callback, sso_login

register_converter(SSOIntegrationConverter, "sso_integration_converter")

urlpatterns = [
    path("<sso_integration_converter:sso>/callback", sso_callback, name="sso-callback"),
    path("<sso_integration_converter:sso>/login", sso_login, name="sso-login"),
]
