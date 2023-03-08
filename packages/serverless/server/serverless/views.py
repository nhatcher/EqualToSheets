from typing import Any

import equalto
from django.http import HttpRequest, HttpResponse, HttpResponseBadRequest, HttpResponseNotAllowed
from graphene_django.views import GraphQLView

from server import settings
from serverless.log import info
from serverless.models import License, LicenseDomain, Workbook
from serverless.schema import schema


# Create your views here.
def index(request: HttpRequest) -> HttpResponse:
    return HttpResponse("Hello from Python! %s, Workbooks.count: %s" % (equalto.__version__, Workbook.objects.count()))


def send_license_key(request: HttpRequest, _send_email: bool = True) -> HttpResponse:
    info("send_license_key(): headers=%s" % request.headers)
    if not settings.DEBUG and request.method != "POST":
        return HttpResponseNotAllowed("405: Method not allowed.")

    email = request.POST.get("email", None) or request.GET.get("email", None)
    if email is None:
        return HttpResponseBadRequest("You must specify the 'email' field.")
    if License.objects.filter(email=email).exists():
        return HttpResponseBadRequest("License key already created for '%s'." % email)
    domain_csv = request.POST.get("domains", "") or request.GET.get("domains", "")
    # WARNING: during the beta, if a license has 0 domains, then the license key will work on all domains
    # TODO: post-beta, we'll require that a license key requires one or more domains
    domains = list(filter(lambda s: s != "", map(lambda s: s.strip(), domain_csv.split(","))))

    # create license & license domains
    license = License(email=email)
    license.save()
    for domain in domains:
        license_domain = LicenseDomain(license=license, domain=domain)
        license_domain.save()

    info("license_id=%s, license key=%s" % (license.id, license.key))
    activation_url = "http://localhost:5000/activate-license-key/%s/" % license.id
    info("activation_url=%s" % activation_url)
    # TODO:
    #   0. Create License and LicenseDomains records
    #   1. send email with newly created license for activation (click on verification link)
    #   2. display instructions ("check email")

    return HttpResponse(
        'License key created: <a href="/activate-license-key/%s">activate %s</a>' % (license.id, license.key),
    )


def activate_license_key(request: HttpRequest, license_id: str) -> HttpResponse:
    license = License.objects.get(id=license_id)
    license.email_verified = True
    license.save()

    return HttpResponse("Your license key: %s. Code samples: ..." % license.key)


def graphql_view(request: HttpRequest, *args: Any, **kwargs: Any) -> HttpResponse:
    return GraphQLView.as_view(graphiql=True, schema=schema)(request, *args, **kwargs)
