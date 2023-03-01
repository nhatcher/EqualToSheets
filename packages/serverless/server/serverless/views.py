import equalto
from django.http import (HttpResponse, HttpResponseBadRequest,
                         HttpResponseNotAllowed)
from django.shortcuts import render
from graphene_django.views import GraphQLView

from .log import error, info
from .models import License, LicenseDomain, Workbook
from .schema import schema
from .util import is_license_key_valid_for_host


# Create your views here.
def index(request):
    return HttpResponse('Hello from Python! %s, Workbooks.count: %s'%(equalto.__version__, Workbook.objects.count()))
    #return render(request, "index.html")



def send_license_key(request, _send_email = True):
    info("send_license_key(): headers=%s"%request.headers)
    if request.method != "POST" and False:
        return HttpResponseNotAllowed("405: Method not allowed.")
    
    email = request.POST.get("email", None) or request.GET.get("email", None)
    if email == None:
        return HttpResponseBadRequest("You must specify the 'email' field.")
    if License.objects.filter(email=email).exists():
        return HttpResponseBadRequest("License key already created for '%s'."%email)
    domain_csv = request.POST.get("domains", None) or request.GET.get("domains", None)
    if domain_csv == None:
        return HttpResponseBadRequest("You must specify the 'domains' field.")
    domains = list(filter( lambda s: s != "", map(lambda s: s.strip(), domain_csv.split(","))))
    if len(domains) == 0:
        return HttpResponseBadRequest("You must specify at least one domain from which you intend to access EqualTo.")
    
    # create license & license domains
    license = License(email=email)
    license.save()
    for domain in domains:
        license_domain = LicenseDomain(license=license, domain=domain)
        license_domain.save()

    info("license_id=%s, license key=%s"%(license.id, license.key))
    activation_url = "http://localhost:5000/activate-license-key/%s/"%license.id
    info("activation_url=%s"%activation_url)
    # TODO:
    #   0. Create License and LicenseDomains records
    #   1. send email with newly created license for activation (click on verification link)
    #   2. display instructions ("check email")

    return HttpResponse('License key created: <a href="/activate-license-key/%s">activate %s</a>'%(license.id, license.key))


def activate_license_key(request, license_id):

    license = License.objects.get(id=license_id)
    license.email_verified = True
    license.save()

    return HttpResponse('Your license key: %s. Code samples: ...'%license.key)



def graphql_view(request, *args, **kwargs):
    return GraphQLView.as_view(graphiql=True, schema=schema)(request, *args, **kwargs)