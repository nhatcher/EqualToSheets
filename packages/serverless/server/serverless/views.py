import tempfile
from asyncio import sleep
from typing import Any
from urllib.parse import quote

import equalto
from asgiref.sync import sync_to_async
from django.db import transaction
from django.http import (
    HttpRequest,
    HttpResponse,
    HttpResponseBadRequest,
    HttpResponseForbidden,
    HttpResponseNotAllowed,
    HttpResponseNotFound,
    JsonResponse,
)
from django.shortcuts import get_object_or_404
from equalto.exceptions import WorkbookError
from graphene_django.views import GraphQLView

from server import settings
from serverless.email import send_license_activation_email
from serverless.log import error, info
from serverless.models import License, LicenseDomain, Workbook
from serverless.schema import schema
from serverless.util import LicenseKeyError, get_license, is_license_key_valid_for_host

MAX_XLSX_FILE_SIZE = 2 * 1024 * 1024


def send_license_key(request: HttpRequest) -> HttpResponse:
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

    send_license_activation_email(license)

    return JsonResponse({"license_id": str(license.id), "license_key": str(license.key)})


def activate_license_key(request: HttpRequest, license_id: str) -> HttpResponse:
    license = get_object_or_404(License, id=license_id)

    license.email_verified = True
    license.save()

    workbook = Workbook.objects.filter(license=license).order_by("create_datetime").first()
    if workbook is None:
        # create a new workbook which can be used in the sample snippet
        workbook = Workbook.objects.create(license=license)

    return JsonResponse({"license_key": str(license.key), "workbook_id": str(workbook.id)})


# Note that you can manually trigger an upload using curl as follows:
#   $ curl -F xlsx-file=@/path/to/file.xlsx
#           -H "Authorization: Bearer <license key>"
#           http://localhost:5000/create-workbook-from-xlsx
def create_workbook_from_xlsx(request: HttpRequest) -> HttpResponse:
    info("create_workbook_from_xlsx(): headers=%s" % request.headers)
    try:
        license = get_license(request.META)
    except LicenseKeyError:
        return HttpResponseForbidden("Invalid license")
    origin = request.META.get("HTTP_ORIGIN")
    if not is_license_key_valid_for_host(license.key, origin):
        error("License key %s is not valid for %s." % (license.key, origin))
        return HttpResponseForbidden("License key is not valid")

    file = request.FILES["xlsx-file"]
    if file.size > MAX_XLSX_FILE_SIZE:
        return HttpResponseBadRequest("Excel file too large (max size %s bytes)." % (MAX_XLSX_FILE_SIZE))
    tmp = tempfile.NamedTemporaryFile()
    tmp.write(file.read())

    try:
        equalto_workbook = equalto.load(tmp.name)
    except WorkbookError as err:
        return HttpResponseBadRequest(f"Could not upload workbook.\n\nDetails:\n\n{err}\n")

    workbook = Workbook(license=license, workbook_json=equalto_workbook.json)
    workbook.save()

    query = (
        """# WARNING: you should avoid sharing the above URL. It contains
#          your license key, which grants full access to all
#          your EqualTo Sheets data.

query {
  workbook(workbookId: "%s") {
    sheets{ name }
  }
}"""
        % workbook.id
    )

    host = request.get_host()
    proto = "https://" if request.is_secure() else "http://"
    content = f"""
Congratulations! The workbook has been uploaded.

Workbook Id: {workbook.id}

Preview workbook: {proto}{host}/edit-workbook/{license.id}/{workbook.id}/

GraphQL query to list all sheets in this workbook: {proto}{host}/graphql?license={license.key}#query={quote(query)}

"""
    return HttpResponse(content, content_type="text/plain")


def graphql_view(request: HttpRequest, *args: Any, **kwargs: Any) -> HttpResponse:
    return GraphQLView.as_view(graphiql=True, schema=schema)(request, *args, **kwargs)


@transaction.non_atomic_requests
async def get_updated_workbook(request: HttpRequest, workbook_id: str, revision: int) -> HttpResponse:
    # TODO: Long poll design isn't the best but should be sufficient for the beta.
    try:
        license = await sync_to_async(lambda: get_license(request.META))()
    except LicenseKeyError:
        return HttpResponseForbidden("Invalid license")

    try:
        workbook = await Workbook.objects.aget(id=workbook_id, license=license)
    except Workbook.DoesNotExist:
        return HttpResponseNotFound("Requested workbook does not exist")

    if workbook.revision < revision:
        # the client's version is newer, that doesn't seem right
        return HttpResponseBadRequest()

    # TODO: We should check the license settings against the request HOST but we're skipping that in the beta.

    for _ in range(50):  # 50 attempts, one connection is open for up to ~10s
        workbook = await _get_updated_workbook(workbook_id, revision)
        if workbook is not None:
            return JsonResponse(workbook)
        await sleep(0.2)

    return HttpResponse(status=408)  # timeout


async def _get_updated_workbook(workbook_id: str, revision: int) -> dict[str, Any] | None:
    workbook = await Workbook.objects.filter(id=workbook_id, revision__gt=revision).order_by("revision").alast()
    if workbook is None:
        return None
    return {
        "revision": workbook.revision,
        "workbook_json": workbook.workbook_json,
    }
