import json
import tempfile
from asyncio import sleep
from typing import Any, Union
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
from equalto.exceptions import SuppressEvaluationErrors, WorkbookError
from graphene_django.views import GraphQLView

from server import settings
from serverless.email import send_license_activation_email
from serverless.log import error, info
from serverless.models import License, LicenseDomain, UnsubscribedEmail, Workbook
from serverless.schema import schema
from serverless.send_email_to_subscribers import send_license_email_to_subscriber
from serverless.types import SimulateInputType, SimulateOutputType, SimulateResultType
from serverless.util import LicenseKeyError, get_license, get_name_from_path, is_license_key_valid_for_host

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

    return HttpResponse(status=201)


def activate_license_key(request: HttpRequest, license_id: str) -> HttpResponse:
    license = get_object_or_404(License, id=license_id)

    license.email_verified = True
    license.save()

    workbook = Workbook.objects.filter(license=license).order_by("create_datetime").first()
    if workbook is None:
        # create a new workbook which can be used in the sample snippet
        workbook = Workbook.objects.create(
            license=license,
            workbook_json=equalto.load("serverless/data/Investment Growth Calculator.xlsx").json,
        )
        workbook.name = "Investment Growth Calculator"
        workbook.save()

    return JsonResponse({"license_key": str(license.key), "workbook_id": str(workbook.id)})


def hacky_send_email_to_subscriber(request: HttpRequest, email: str) -> HttpResponse:
    try:
        send_license_email_to_subscriber(email)
    except Exception as e:
        return HttpResponseBadRequest(f"Failed: {e}")
    return HttpResponse(f"Email sent for {email}")


def edit_workbook(request: HttpRequest, license_key: str, workbook_id: str) -> HttpResponse:
    workbook = Workbook.objects.filter(license__key=license_key, id=workbook_id).order_by("create_datetime").first()
    if workbook is None:
        return HttpResponseNotFound("Workbook not found")

    host = request.get_host()
    proto = "https://" if request.is_secure() else "http://"

    html = f"""<!doctype html>
<html lang="en">
    <head>
        <meta charset="utf-8"/>
        <title>EqualTo Sheets</title>
        <script type="text/javascript" src="/static/v1/equalto.js"></script>
        <style>
            html {{
                height: 100%;
            }}
            body {{
                height: 100%;
                margin: 0;
            }}
            #container {{
                height: 100%;
                display: flex;
                flex-direction: column;
                padding: 20px;
                box-sizing: border-box;
            }}
            #workbook-slot {{
                flex-grow: 1;
            }}
            .column {{
                float: left;
                width: 50%;
                height:100%;
            }}
            .row {{
                height:100%;
            }}

            /* Clear floats after the columns */
            .row:after {{
                content: "";
                display: table;
                clear: both;
            }}
        </style>
    </head>
    <body>
        <div id="container">
            <h1>WARNING: you should avoid sharing the above URL. It contains your license key, which
                allows full access to all your EqualTo Sheets data.</h1>
            <div class="row">
                <div class="column">
                    <pre>
&lt;div id="workbook-slot" style="height:100%"&gt;&lt;/div&gt;
&lt;script src="{proto}{host}/static/v1/equalto.js"&gt;&lt;/script&gt;
&lt;script&gt;
    // WARNING: do not expose your license key in client code,
    //          instead you should proxy calls to EqualTo.
    EqualToSheets.setLicenseKey(
        "{license_key}"
    );
    // Insert spreadsheet widget into the DOM
    EqualToSheets.load(
        "{workbook.id}",
        document.getElementById("workbook-slot")
    );
&lt;/script&gt;
                    </pre>
                </div>
                <div id="workbook-slot" class="column"></div>
            </div>
            <script type="text/javascript">
                EqualToSheets.setLicenseKey(
                    "{license_key}"
                );
                // Insert spreadsheet widget into the DOM
                EqualToSheets.load(
                    "{workbook.id}",
                    document.getElementById("workbook-slot")
                );
            </script>
        </div>
    </body>
</html>
"""

    return HttpResponse(html)


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

    with SuppressEvaluationErrors() as context:
        try:
            equalto_workbook = equalto.load(tmp.name)
        except WorkbookError as err:
            return HttpResponseBadRequest(f"Could not upload workbook.\n\nDetails:\n\n{err}\n")
        compatibility_errors = context.suppressed_errors(equalto_workbook)

    workbook_name = get_name_from_path(file.name)
    workbook = Workbook(license=license, workbook_json=equalto_workbook.json, name=workbook_name)
    workbook.save()

    query = (
        """# WARNING: you should avoid sharing the above URL. It contains
#          your license key, which grants full access to all
#          your EqualTo Sheets data.

query {
  workbook(workbookId: "%s") {
    name,
    sheets{ name }
  }
}"""
        % workbook.id
    )

    content = f"""
Congratulations! The workbook has been uploaded.

Workbook Id: {workbook.id}

Preview workbook: {settings.SERVER}/edit-workbook/{license.key}/{workbook.id}/

GraphQL query to list all sheets in this workbook: {settings.SERVER}/graphql?license={license.key}#query={quote(query)}

"""

    if compatibility_errors:
        error_message = "\n".join(compatibility_errors)
        error(error_message)
        content = f"{content}{error_message}\n\n"

    return HttpResponse(content, content_type="text/plain")


def unsubscribe_email(request: HttpRequest) -> HttpResponse:
    info("unsubscribe_email(): headers=%s" % request.headers)
    email = request.GET.get("email")
    if not email:
        return HttpResponseBadRequest("You must specify the email GET parameter.")
    if UnsubscribedEmail.objects.filter(email=email).count() > 0:
        return HttpResponse(f"The email address {email} has already been unsubscribed from all EqualTo mailings.")
    unsubscribed_email = UnsubscribedEmail(email=email)
    unsubscribed_email.save()
    return HttpResponse(f"The email address {email} has unsubscribed from all EqualTo mailings.")


def simulate(request: HttpRequest, workbook_id: str) -> Union[HttpResponse, JsonResponse]:
    info("simulate(): headers=%s" % request.headers)
    try:
        license = get_license(request.META)
    except LicenseKeyError:
        return HttpResponseForbidden("Invalid license")
    origin = request.META.get("HTTP_ORIGIN")
    if not is_license_key_valid_for_host(license.key, origin):
        error("License key %s is not valid for %s." % (license.key, origin))
        return HttpResponseForbidden("License key is not valid")

    def get(param: str) -> str:
        post_param = request.POST.get(param)
        return post_param if post_param else request.GET.get(param)

    workbook = get_object_or_404(Workbook, license=license, id=workbook_id)

    equalto_workbook = workbook.calc

    inputs_str = get("inputs")
    try:
        inputs: SimulateInputType = json.loads(inputs_str)
    except json.JSONDecodeError:
        return HttpResponseBadRequest(f"Invalid inputs: {inputs_str}")

    outputs_str = get("outputs")
    try:
        outputs: SimulateOutputType = json.loads(outputs_str)
    except json.JSONDecodeError:
        return HttpResponseBadRequest(f"Invalid outputs: {outputs_str}")

    with SuppressEvaluationErrors():
        for sheet_name, assignments in inputs.items():
            for cell_ref, value in assignments.items():
                equalto_workbook.sheets[sheet_name][cell_ref].value = value

    results: SimulateResultType = {}
    for sheet_name, cell_refs in outputs.items():  # noqa: WPS440
        results[sheet_name] = {}
        for cell_ref in cell_refs:  # noqa: WPS440
            results[sheet_name][cell_ref] = equalto_workbook.sheets[sheet_name][cell_ref].value

    return JsonResponse(results)


class LicenseGraphQLView(GraphQLView):
    graphiql_template = "graphiql.html"
    license_key: str | None = None

    def __init__(self, license_key: str | None) -> None:
        super().__init__(graphiql=True, schema=schema)
        self.license_key = license_key  # noqa: WPS601

    def render_graphiql(self, request: HttpRequest, **data: Any) -> HttpResponse:
        return super().render_graphiql(request, license_key=self.license_key, **data)


def graphql_view(request: HttpRequest, *args: Any, **kwargs: Any) -> HttpResponse:
    return LicenseGraphQLView.as_view(license_key=request.GET.get("license"))(request, *args, **kwargs)


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
