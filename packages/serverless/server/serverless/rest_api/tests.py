import json

from django.test import TestCase
from equalto.exceptions import SuppressEvaluationErrors
from rest_framework.test import APIClient

from serverless.models import License, Workbook
from serverless.test_util import create_verified_license, create_workbook


class RestAPITest(TestCase):
    license: License
    another_license: License

    @classmethod
    def setUpTestData(cls) -> None:
        cls.license = create_verified_license(domains="")
        cls.another_license = create_verified_license(email="bob@example.com", domains="")

        with SuppressEvaluationErrors():
            create_workbook(
                cls.license,
                {
                    "Calculation": {"A1": "=2*Data!A2"},
                    "Data": {
                        "A1": "$3.99",
                        "A2": 4,
                        "A3": "=UNSUPPORTED()",  # one invalid formula confirming that unsupported files can be edited
                    },
                },
                "WorkbookName",
            )

        # workbook linked to another license shouldn't be visible for self.license_client
        create_workbook(cls.another_license)

    def setUp(self) -> None:
        self.unauthorized_client = APIClient()

        self.license_client = APIClient()
        self.license_client.credentials(HTTP_AUTHORIZATION=f"Bearer {self.license.key}")

        self.another_license_client = APIClient()
        self.another_license_client.credentials(HTTP_AUTHORIZATION=f"Bearer {self.another_license.key}")

        self.workbook = Workbook.objects.get(license=self.license)
        self.another_workbook = Workbook.objects.exclude(license=self.license).get()

    def test_get_workbooks(self) -> None:
        response = self.license_client.get("/api/v1/workbooks")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "workbooks": [
                    {
                        "id": str(self.workbook.id),
                        "name": "WorkbookName",
                        "revision": 1,
                        "create_datetime": self.workbook.create_datetime.isoformat().replace("+00:00", "Z"),
                        "modify_datetime": self.workbook.modify_datetime.isoformat().replace("+00:00", "Z"),
                    },
                ],
            },
        )

    def test_get_workbooks_missing_license(self) -> None:
        response = self.unauthorized_client.get("/api/v1/workbooks")
        self.assertEqual(response.status_code, 403)

    def test_create_workbook(self) -> None:
        response = self.license_client.post("/api/v1/workbooks")
        self.assertEqual(response.status_code, 201)
        workbook_data = json.loads(response.content)
        workbook = Workbook.objects.get(id=workbook_data["id"])
        self.assertEqual(workbook.license, self.license)
        self.assertEqual(
            workbook_data,
            {
                "id": str(workbook.id),
                "name": "Book",
                "revision": 1,
                "create_datetime": workbook.create_datetime.isoformat().replace("+00:00", "Z"),
                "modify_datetime": workbook.modify_datetime.isoformat().replace("+00:00", "Z"),
            },
        )

    def test_get_workbook(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "id": str(self.workbook.id),
                "name": "WorkbookName",
                "revision": 1,
                "create_datetime": self.workbook.create_datetime.isoformat().replace("+00:00", "Z"),
                "modify_datetime": self.workbook.modify_datetime.isoformat().replace("+00:00", "Z"),
                "workbook_json": self.workbook.workbook_json,
            },
        )

    def test_get_workbook_missing_license(self) -> None:
        response = self.unauthorized_client.get(f"/api/v1/workbooks/{self.workbook.id}")
        self.assertEqual(response.status_code, 403)

    def test_get_workbook_invalid_license(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.another_workbook.id}")
        self.assertEqual(response.status_code, 404)
        self.assertEqual(json.loads(response.content), {"detail": "Workbook not found"})

    def test_export_workbook_to_xlsx(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/xlsx")
        self.assertEqual(response.status_code, 200)
        # TODO: parse the XLSX file and confirm it's as expected
        self.assertEqual(response.content[:2], b"PK")

    def test_export_workbook_to_xlsx_invalid_license(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.another_workbook.id}/xlsx")
        self.assertEqual(response.status_code, 404)
        self.assertEqual(json.loads(response.content), {"detail": "Workbook not found"})

    def test_get_sheets(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "sheets": [
                    {"id": 1, "name": "Calculation", "index": 0},
                    {"id": 2, "name": "Data", "index": 1},
                ],
            },
        )

    def test_get_sheets_missing_license(self) -> None:
        response = self.unauthorized_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets")
        self.assertEqual(response.status_code, 403)

    def test_get_sheets_invalid_license(self) -> None:
        response = self.another_license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets")
        self.assertEqual(response.status_code, 404)

    def test_create_sheet(self) -> None:
        response = self.license_client.post(
            f"/api/v1/workbooks/{self.workbook.id}/sheets",
            {"name": "Analytics"},
        )
        self.assertEqual(response.status_code, 201)
        self.assertEqual(
            json.loads(response.content),
            {"id": 3, "name": "Analytics", "index": 2},
        )

        self.assertEqual(
            json.loads(self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets").content),
            {
                "sheets": [
                    {"id": 1, "name": "Calculation", "index": 0},
                    {"id": 2, "name": "Data", "index": 1},
                    {"id": 3, "name": "Analytics", "index": 2},
                ],
            },
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_create_sheet_missing_license(self) -> None:
        response = self.unauthorized_client.post(f"/api/v1/workbooks/{self.workbook.id}/sheets")
        self.assertEqual(response.status_code, 403)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_create_sheet_invalid_license(self) -> None:
        response = self.another_license_client.post(f"/api/v1/workbooks/{self.workbook.id}/sheets")
        self.assertEqual(response.status_code, 404)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_create_sheet_default_name(self) -> None:
        response = self.license_client.post(f"/api/v1/workbooks/{self.workbook.id}/sheets")
        self.assertEqual(response.status_code, 201)
        self.assertEqual(
            json.loads(response.content),
            {"id": 3, "name": "Sheet1", "index": 2},
        )

        self.assertEqual(
            json.loads(self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets").content),
            {
                "sheets": [
                    {"id": 1, "name": "Calculation", "index": 0},
                    {"id": 2, "name": "Data", "index": 1},
                    {"id": 3, "name": "Sheet1", "index": 2},
                ],
            },
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_create_sheet_name_in_use(self) -> None:
        response = self.license_client.post(
            f"/api/v1/workbooks/{self.workbook.id}/sheets",
            {"name": "Data"},
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            json.loads(response.content),
            {"detail": "A worksheet already exists with that name"},
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_get_sheet(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {"id": 2, "name": "Data", "index": 1},
        )

    def test_get_invalid_sheet(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/3")
        self.assertEqual(response.status_code, 404)

    def test_get_sheet_missing_license(self) -> None:
        response = self.unauthorized_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2")
        self.assertEqual(response.status_code, 403)

    def test_get_sheet_invalid_license(self) -> None:
        response = self.another_license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2")
        self.assertEqual(response.status_code, 404)

    def test_rename_sheet(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/1",
            {"new_name": "New Calculation Sheet"},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {"id": 1, "name": "New Calculation Sheet", "index": 0},
        )

        self.assertEqual(
            json.loads(self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets").content),
            {
                "sheets": [
                    {"id": 1, "name": "New Calculation Sheet", "index": 0},
                    {"id": 2, "name": "Data", "index": 1},
                ],
            },
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_rename_sheet_missing_license(self) -> None:
        response = self.unauthorized_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/1",
            {"new_name": "New Calculation Sheet"},
        )
        self.assertEqual(response.status_code, 403)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_rename_sheet_invalid_license(self) -> None:
        response = self.another_license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/1",
            {"new_name": "New Calculation Sheet"},
        )
        self.assertEqual(response.status_code, 404)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_rename_invalid_sheet(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/3",
            {"new_name": "New Sheet Name"},
        )
        self.assertEqual(response.status_code, 404)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_rename_sheet_name_in_use(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/1",
            {"new_name": "Data"},
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            json.loads(response.content),
            {"detail": "Sheet already exists: 'Data'"},
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_rename_sheet_missing_name(self) -> None:
        response = self.license_client.put(f"/api/v1/workbooks/{self.workbook.id}/sheets/1")
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            json.loads(response.content),
            {"detail": "'new_name' parameter is not provided"},
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_delete_sheet(self) -> None:
        response = self.license_client.delete(f"/api/v1/workbooks/{self.workbook.id}/sheets/1")
        self.assertEqual(response.status_code, 204)

        self.assertEqual(
            json.loads(self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets").content),
            {
                "sheets": [
                    {"id": 2, "name": "Data", "index": 0},
                ],
            },
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_delete_sheet_missing_license(self) -> None:
        response = self.unauthorized_client.delete(f"/api/v1/workbooks/{self.workbook.id}/sheets/1")
        self.assertEqual(response.status_code, 403)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_delete_sheet_invalid_license(self) -> None:
        response = self.another_license_client.delete(f"/api/v1/workbooks/{self.workbook.id}/sheets/1")
        self.assertEqual(response.status_code, 404)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_delete_last_sheet(self) -> None:
        self.license_client.delete(f"/api/v1/workbooks/{self.workbook.id}/sheets/1")

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

        response = self.license_client.delete(f"/api/v1/workbooks/{self.workbook.id}/sheets/2")
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            json.loads(response.content),
            {"detail": "Cannot delete only sheet"},
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_delete_invalid_sheet(self) -> None:
        response = self.license_client.delete(f"/api/v1/workbooks/{self.workbook.id}/sheets/3")
        self.assertEqual(response.status_code, 404)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_get_cell(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/1/1")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "formatted_value": "$3.99",
                "value": 3.99,
                "format": "$#,##0.00",
                "type": "number",
                "formula": None,
            },
        )

    def test_get_cell_missing_license(self) -> None:
        response = self.unauthorized_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/1/1")
        self.assertEqual(response.status_code, 403)

    def test_get_cell_invalid_license(self) -> None:
        response = self.another_license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/1/1")
        self.assertEqual(response.status_code, 404)

    def test_get_invalid_cell(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/1/0")
        self.assertEqual(response.status_code, 404)

    def test_get_cell_with_formula(self) -> None:
        response = self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/1/cells/1/1")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "formatted_value": "8",
                "value": 8,
                "format": "general",
                "type": "number",
                "formula": "=2*Data!A2",
            },
        )

    def test_set_cell_input(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/2/1",  # Data!A2
            {"input": "$16"},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "formatted_value": "$16",
                "value": 16.0,
                "format": "$#,##0",
                "type": "number",
                "formula": None,
            },
        )

        # confirm that affected cells are also evaluated
        self.assertEqual(
            json.loads(self.license_client.get(f"/api/v1/workbooks/{self.workbook.id}/sheets/1/cells/1/1").content),
            {
                "formatted_value": "32",
                "value": 32.0,
                "format": "general",
                "type": "number",
                "formula": "=2*Data!A2",
            },
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_set_cell_value(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/2/1",
            {"value": "$16"},  # sets "$16" string, doesn't try to recognize the formatting
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "formatted_value": "$16",
                "value": "$16",
                "format": "general",
                "type": "text",
                "formula": None,
            },
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 2)

    def test_set_cell_input_xor_value(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/2/1",
            {"value": "$16", "input": 17},
        )
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            json.loads(response.content),
            {"detail": "Either 'input' or 'value' parameter needs to be provided, but not both"},
        )
        response = self.license_client.put(f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/2/1")
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            json.loads(response.content),
            {"detail": "Either 'input' or 'value' parameter needs to be provided, but not both"},
        )

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)

    def test_set_invalid_cell(self) -> None:
        response = self.license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/0/1",
            {"input": "$16"},
        )
        self.assertEqual(response.status_code, 404)

    def test_set_cell_input_missing_license(self) -> None:
        response = self.unauthorized_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/2/1",  # Data!A2
            {"input": "$16"},
        )
        self.assertEqual(response.status_code, 403)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)
        self.assertEqual(self.workbook.calc["Data!A2"].value, 4)

    def test_set_cell_input_invalid_license(self) -> None:
        response = self.another_license_client.put(
            f"/api/v1/workbooks/{self.workbook.id}/sheets/2/cells/2/1",  # Data!A2
            {"input": "$16"},
        )
        self.assertEqual(response.status_code, 404)

        self.workbook.refresh_from_db()
        self.assertEqual(self.workbook.revision, 1)
        self.assertEqual(self.workbook.calc["Data!A2"].value, 4)
