import html
import json
import tempfile
from asyncio import sleep
from collections import defaultdict
from typing import Any, Union
from urllib.parse import quote
from zoneinfo import ZoneInfo

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
from django.utils import timezone
from equalto.exceptions import CellReferenceError, SuppressEvaluationErrors, WorkbookError
from equalto.sheet import Sheet
from graphene_django.views import GraphQLView

from server import settings
from serverless.email import send_license_activation_email
from serverless.log import error, info
from serverless.models import License, LicenseDomain, UnsubscribedEmail, Workbook
from serverless.schema import schema
from serverless.send_email_to_subscribers import send_license_email_to_subscriber
from serverless.types import CellRangeValue, CellValue, SimulateInputType, SimulateOutputType, SimulateResultType
from serverless.util import LicenseKeyError, get_license, get_name_from_path, is_license_key_valid_for_host

MAX_XLSX_FILE_SIZE = 2 * 1024 * 1024


def send_license_key(request: HttpRequest) -> HttpResponse:
    info("send_license_key(): headers=%s" % request.headers)
    if not settings.DEBUG and request.method != "POST":
        return HttpResponseNotAllowed("405: Method not allowed.")

    email = request.POST.get("email", None) or request.GET.get("email", None)
    if email is None:
        return HttpResponseBadRequest("You must specify the 'email' field.")

    license = License.objects.filter(email=email).first()
    if license is None:
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

    if not license.email_verified:
        license.email_verified = True
        license.validated_datetime = timezone.now()
        license.save()

    # The "new" sample workbooks (for the open beta) were created with a created_datetime of 2023-04-19 12:00:00 UTC in
    # migration 0007_create_new_beta_workbook.py.We want to make sure that we return the first workbook associated with
    # a license key, on or after 2023-04-19 12:00:00 UTC, so that license keys created prior to 2023-04-19 12:00:00 get
    # the new workbook, not the "old" sample workbook.
    workbook = (
        Workbook.objects.filter(
            license=license,
            create_datetime__gte=timezone.datetime(2023, 4, 19, 12, 0, 0, tzinfo=ZoneInfo("UTC")),
        )
        .order_by("create_datetime")  # noqa: WPS348
        .first()  # noqa: WPS348
    )
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

    equalto_svg = """<svg xmlns="http://www.w3.org/2000/svg" width="35" height="36" fill="none">
  <rect width="34" height="34" x=".5" y="1" fill="url(#a)" rx="4.5"/>
  <g filter="url(#b)">
    <path fill="url(#c)" fill-rule="evenodd" d="M11.5 10.5a.5.5 0 0 0-.5.5v5.72a.5.5 0 0 0 .5.5h12.714a.5.5 0 0 0 .5-.5V11a.5.5 0 0 0-.5-.5H11.5Zm6.857 9.28a.5.5 0 0 0-.5.5V26a.5.5 0 0 0 .5.5h5.857a.5.5 0 0 0 .5-.5v-5.72a.5.5 0 0 0-.5-.5h-5.857Z" clip-rule="evenodd"/>
  </g>
  <rect width="34" height="34" x=".5" y="1" stroke="#343855" rx="4.5"/>
  <defs>
    <linearGradient id="a" x1="17.5" x2="17.5" y1=".5" y2="35.5" gradientUnits="userSpaceOnUse">
      <stop stop-color="#292C42"/>
      <stop offset="1" stop-color="#1F2236"/>
    </linearGradient>
    <linearGradient id="c" x1="17.857" x2="17.857" y1="10.5" y2="26.5" gradientUnits="userSpaceOnUse">
      <stop stop-color="#72ED79"/>
      <stop offset="1" stop-color="#53CD5A"/>
    </linearGradient>
    <filter id="b" width="15.714" height="19" x="10" y="10.5" color-interpolation-filters="sRGB" filterUnits="userSpaceOnUse">
      <feFlood flood-opacity="0" result="BackgroundImageFix"/>
      <feColorMatrix in="SourceAlpha" result="hardAlpha" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 127 0"/>
      <feOffset dy="2"/>
      <feGaussianBlur stdDeviation=".5"/>
      <feComposite in2="hardAlpha" operator="out"/>
      <feColorMatrix values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0.15 0"/>
      <feBlend in2="BackgroundImageFix" result="effect1_dropShadow_253_33415"/>
      <feBlend in="SourceGraphic" in2="effect1_dropShadow_253_33415" result="shape"/>
    </filter>
  </defs>
</svg>"""  # noqa: E501

    snippet_unescaped = f"""<div id="workbook-slot" style="height:100%;min-height:300px;"></div>
<script src="{proto}{host}/static/v1/equalto.js"></script>
<script>
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
</script>"""

    snippet_json = json.dumps(snippet_unescaped).replace("</", r"<\/")

    snippet_html = html.escape(snippet_unescaped)

    snippet_lines = len(snippet_unescaped.splitlines())
    snippet_lines_html = "".join([f"<li>{line}</li>" for line in range(1, snippet_lines + 1 + 1)])

    response_html = f"""<!doctype html>
<html lang="en">
    <head>
        <meta charset="utf-8"/>
        <title>EqualTo Sheets</title>
        <script type="text/javascript" src="/static/v1/equalto.js"></script>
        <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Inter:200,300,regular,500,600,700,800%7CFira+Mono:regular%7CFira+Code:300,regular,500,600%7CJetBrains+Mono:100,200,300,regular,500,600,700,800" media="all">
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
                box-sizing: border-box;
                position: relative;
            }}

            #workbook-slot {{
                flex-grow: 1;
            }}

            .warning-bar {{
                font-family: "Inter";
                font-style: normal;
                font-weight: 500;
                font-size: 9px;
                line-height: 11px;
                text-align: center;
                color: #21243A;
                background: #F5BB49;
                padding: 0 10px;
                padding: 5px;
            }}

            .radial-background {{
                background: radial-gradient(
                    70.43% 179.85% at 21.99% 27.53%,
                    rgba(88, 121, 240, 0.07) 0%,
                    rgba(88, 121, 240, 0) 100%
                    ),
                    linear-gradient(180deg, #292c42 0%, #1f2236 100%);
                background-blend-mode: normal, normal;
            }}

            .heading-bar {{
                display: flex;
                flex-direction: row;
                align-items: center;
                padding: 10px;
                border-bottom: 1px solid #545971;
            }}

            .heading-bar .sheets {{
                margin-left: 10px;
                font-family: 'JetBrains Mono', monospace;
                font-style: normal;
                font-weight: 400;
                font-size: 13px;
                line-height: 17px;
                color: #C6CAE3;
            }}

            .columns {{
                flex: 1;
                display: grid;
                grid-template-columns: 605px 1fr;
            }}

            .snippet {{
                display: grid;
                grid-template-columns: min-content 1fr;
                gap: 15px;
                background: rgba(255, 255, 255, 0.05);
                overflow: auto;
                cursor: pointer;
            }}

            .snippet .numbers {{
                list-style: none;
                padding: 0;
                margin: 0;
                user-select: none;
                padding: 20px 0 20px 20px;
            }}

            .snippet .numbers li {{
                font-family: 'JetBrains Mono';
                font-style: normal;
                font-weight: 300;
                font-size: 13px;
                line-height: 18px;
                color: #8B8FAD;
            }}

            .snippet pre {{
                padding: 20px 20px 20px 0;
                margin: 0;
                font-family: 'JetBrains Mono';
                font-style: normal;
                font-weight: 300;
                font-size: 13px;
                line-height: 18px;
                color: #D3D6E9;
            }}

            .footer {{
                display: flex;
                flex-direction: row;
                justify-content: center;
                align-items: center;
                padding: 5px;
                border-top: 1px solid #545971;
            }}

            .footer .made-by {{
                font-family: 'Inter';
                font-style: normal;
                font-weight: 400;
                font-size: 9px;
                line-height: 11px;
                color: #8B8FAD;
            }}

            .footer .divider {{
                height: 8px;
                margin: 0 10px;
                border-left: 1px solid #5F6989;
            }}

            .footer a:link,
            .footer a:visited,
            .footer a:hover,
            .footer a:active {{
                font-family: 'Inter';
                font-style: normal;
                font-weight: 400;
                font-size: 9px;
                line-height: 11px;
                color: #5879F0;
            }}

            .popover {{
                position: absolute;
                bottom: 20px;
                left: 20px;

                font-family: Inter, sans-serif;
                font-size: 14px;
                font-weight: 400;
                line-height: 1.43;
                color: rgb(255, 255, 255);
                background-color: rgb(50, 50, 50);
                text-align: center;
                padding: 6px 16px;
                border-radius: 4px;
                user-select: none;
                box-shadow: rgba(0, 0, 0, 0.2) 0px 3px 5px -1px,
                    rgba(0, 0, 0, 0.14) 0px 6px 10px 0px,
                    rgba(0, 0, 0, 0.12) 0px 1px 18px 0px;
                opacity: 1;
                transition: 0.2s ease-in-out opacity;
            }}

            .popover-hidden {{
                pointer-events: none;
                opacity: 0;
            }}

            .equalto-sheets-workbook {{
                border: none;
                filter: none;
            }}

        </style>
    </head>
    <body>
        <div id="container">
            <div class="warning-bar">
                Warning: You should avoid sharing the above URL. It contains your license key,
                which allows full access to all your EqualTo Sheets data.
            </div>

            <div class="heading-bar radial-background">
                {equalto_svg}
                <div class="sheets">
                    Sheets
                </div>
            </div>

            <div class="columns radial-background">
                <div class="snippet">
                    <ul class="numbers">
                        {snippet_lines_html}
                    </ul>
                    <pre>{snippet_html}</pre>
                </div>
                <div id="workbook-slot" class="column"></div>
            </div>

            <div class="footer radial-background">
                <span class="made-by">Made in Berlin and Warsaw by the EqualTo team</span>
                <span class="divider"></span>
                <a href="https://www.equalto.com/" target="_blank">
                    equalto.com
                </a>
            </div>

            <div id="copied" class="popover popover-hidden">
                Copied to clipboard.
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

                let currentTimeout = null;
                document.querySelector('.snippet').addEventListener('click', () => {{
                    let popover = document.querySelector('#copied');
                    popover.classList.remove('popover-hidden');
                    navigator.clipboard.writeText({snippet_json}).then(() => {{
                        if (currentTimeout) {{
                            clearTimeout(currentTimeout);
                            currentTimeout = null;
                        }}
                        currentTimeout = setTimeout(() => {{
                            popover.classList.add('popover-hidden');
                        }}, 2000);
                    }});
                }});
            </script>
        </div>
    </body>
</html>
"""  # noqa: E501

    return HttpResponse(response_html)


# Note that you can manually trigger an upload using curl as follows:
#   $ curl -F xlsx-file=@/path/to/file.xlsx
#           -H "Authorization: Bearer <license key>"
#           http://localhost:5000/create-workbook-from-xlsx
# TODO: delete this method post-beta. POST /api/v1/workbooks already supports uploading of XLSX files.
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
        """# WARNING: you should not share the above URL. It contains
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

Open spreadsheet: {settings.SERVER}/unsafe-just-for-beta/edit-workbook/{license.key}/{workbook.id}/

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

    workbook = get_object_or_404(Workbook, license=license, id=workbook_id)

    equalto_workbook = workbook.calc

    inputs: SimulateInputType
    outputs: SimulateOutputType
    if request.method == "GET":
        inputs_str = request.GET.get("inputs")
        try:
            inputs = json.loads(inputs_str)
        except json.JSONDecodeError:
            return HttpResponseBadRequest(f"Invalid inputs: {inputs_str}")

        outputs_str = request.GET.get("outputs")
        try:
            outputs = json.loads(outputs_str)
        except json.JSONDecodeError:
            return HttpResponseBadRequest(f"Invalid outputs: {outputs_str}")
    elif request.method == "POST":
        try:  # noqa: WPS229
            request_data = json.loads(request.body)
            inputs = request_data["inputs"]
            outputs = request_data["outputs"]
        except (ValueError, KeyError):
            return HttpResponseBadRequest(f"Invalid POST data: {request.body}")
    else:
        return HttpResponseNotAllowed("Method not allowed")

    with SuppressEvaluationErrors():
        for sheet_name, assignments in inputs.items():
            try:
                sheet = equalto_workbook.sheets[sheet_name]
            except WorkbookError as err:
                return HttpResponseBadRequest(str(err))

            for ref, value in assignments.items():
                try:
                    _set_cell_or_range_value(sheet, ref, value)
                except (ValueError, CellReferenceError) as err:
                    return HttpResponseBadRequest(str(err))

    results: SimulateResultType = defaultdict(dict)
    for sheet_name, refs in outputs.items():
        try:
            sheet = equalto_workbook.sheets[sheet_name]
        except WorkbookError as err:
            return HttpResponseBadRequest(str(err))

        try:
            for ref in refs:
                if ":" in ref:
                    results[sheet_name][ref] = [[cell.value for cell in row] for row in sheet.cell_range(ref)]
                else:
                    results[sheet_name][ref] = sheet[ref].value
        except CellReferenceError as err:
            return HttpResponseBadRequest(str(err))

    return JsonResponse(results)


def _set_cell_or_range_value(sheet: Sheet, ref: str, value: CellValue | CellRangeValue) -> None:  # noqa: WPS238
    error_message = f"{value} is not a valid value for {ref}"

    if ":" in ref:
        cell_range = sheet.cell_range(ref)
        if not isinstance(value, list) or len(value) != len(cell_range):
            raise ValueError(error_message)
        for row, row_data in zip(cell_range, value):
            if not isinstance(row_data, list) or len(row_data) != len(row):
                raise ValueError(error_message)
            for cell, cell_value in zip(row, row_data):
                if cell_value is not None and not isinstance(cell_value, (str, float, bool, int)):
                    raise ValueError(error_message)
                cell.value = cell_value
    else:
        if value is not None and not isinstance(value, (str, float, bool, int)):
            raise ValueError(error_message)
        sheet[ref].value = value


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

    return HttpResponse(status=204)  # no content


async def _get_updated_workbook(workbook_id: str, revision: int) -> dict[str, Any] | None:
    workbook = await Workbook.objects.filter(id=workbook_id, revision__gt=revision).order_by("revision").alast()
    if workbook is None:
        return None
    return {
        "revision": workbook.revision,
        "workbook_json": workbook.workbook_json,
    }
