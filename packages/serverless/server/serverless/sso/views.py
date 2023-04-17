from django.conf import settings
from django.http.request import HttpRequest
from django.http.response import HttpResponse, HttpResponseBadRequest, HttpResponseForbidden, HttpResponseRedirect

from serverless.models import License
from serverless.sso.exceptions import SSOError
from serverless.sso.oauth import OAuthSSO


def sso_callback(request: HttpRequest, sso: OAuthSSO) -> HttpResponse:  # noqa: WPS212, C901
    if "error" in request.GET:
        # this can happen when the user denies access to their data
        return HttpResponseForbidden("Authorization has not been approved.")

    authorization_code = request.GET.get("code")
    if not authorization_code:
        return HttpResponseBadRequest("Missing authorization code.")

    try:
        email = sso.retrieve_user_email(authorization_code)
    except SSOError:
        return HttpResponseForbidden("Could not retrieve the user data.")

    try:
        license = License.objects.get(email=email)
    except License.DoesNotExist:
        # TODO: post-beta, support for license domains
        license = License.objects.create(email=email)

    return HttpResponseRedirect(f"{settings.SERVER}/#/license/activate/{license.id}")


def sso_login(request: HttpRequest, sso: OAuthSSO) -> HttpResponse:
    return HttpResponseRedirect(sso.login_url)
