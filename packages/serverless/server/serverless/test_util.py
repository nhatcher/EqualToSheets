from typing import Any

import equalto
from django.test import RequestFactory

from serverless.models import License, Workbook
from serverless.views import send_license_key


def create_workbook(license: License, workbook_data: dict[str, Any] | None = None) -> Workbook:
    workbook = Workbook.objects.create(license=license)

    if workbook_data is not None:
        _set_workbook_data(workbook, workbook_data)

    return workbook


def _set_workbook_data(workbook: Workbook, workbook_data: dict[str, Any]) -> None:
    wb = equalto.new()
    for _ in range(len(workbook_data) - 1):
        wb.sheets.add()
    for sheet_index, (sheet_name, sheet_data) in enumerate(workbook_data.items()):
        sheet = wb.sheets[sheet_index]
        sheet.name = sheet_name
        for cell_ref, user_input in sheet_data.items():
            sheet[cell_ref].set_user_input(str(user_input))

    workbook.workbook_json = wb.json
    workbook.save()


def create_verified_license(
    email: str = "joe@example.com",
    domains: str = "example.com,example2.com,*.example3.com",
) -> License:
    before_license_ids = list(License.objects.values_list("id", flat=True))
    request = RequestFactory().post("/send-license-key", {"email": email, "domains": domains})
    response = send_license_key(request)
    assert response.status_code == 201, response.content
    after_license_ids = License.objects.values_list("id", flat=True)
    new_license_ids = list(set(after_license_ids).difference(before_license_ids))
    assert len(new_license_ids) == 1, new_license_ids
    license = License.objects.get(id=new_license_ids[0])
    license.email_verified = True
    license.save()
    return license
