import json
from asyncio import sleep
from collections import namedtuple
from typing import Any
from unittest.mock import MagicMock, patch

import equalto
from asgiref.sync import sync_to_async
from django.core.files.uploadedfile import SimpleUploadedFile
from django.db import transaction
from django.test import AsyncClient, RequestFactory, TestCase, TransactionTestCase, override_settings
from django.utils.http import urlencode
from graphql import GraphQLError

from serverless.email import LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID
from serverless.log import info
from serverless.models import License, LicenseDomain, Workbook
from serverless.schema import MAX_WORKBOOK_INPUT_SIZE, MAX_WORKBOOK_JSON_SIZE, MAX_WORKBOOKS_PER_LICENSE, schema
from serverless.util import is_license_key_valid_for_host
from serverless.views import MAX_XLSX_FILE_SIZE, activate_license_key, create_workbook_from_xlsx, send_license_key


@transaction.atomic
def graphql_query(
    query: str,
    origin: str,
    license_key: str | None = None,
    variables: dict[str, Any] | None = None,
    suppress_errors: bool = False,
) -> dict[str, Any]:
    info("graphql_query(): query=%s" % query)
    context = namedtuple("context", ["META"])
    context.META = {"HTTP_ORIGIN": origin}
    if license_key is not None:
        context.META["HTTP_AUTHORIZATION"] = "Bearer %s" % str(license_key)

    info("graphql_query(): context.META=%s" % context.META)
    graphql_results = schema.execute(query, context_value=context, variable_values=variables)
    if not suppress_errors and graphql_results.errors:
        raise graphql_results.errors[0]
    return {"data": graphql_results.data}


def _create_workbook(license: License, workbook_data: dict[str, Any] | None = None) -> Workbook:
    data = graphql_query(
        """
        mutation {
            create_workbook: createWorkbook {
                workbook{id}
            }
        }
        """,
        "example.com",
        license.key,
    )
    workbook = Workbook.objects.get(id=data["data"]["create_workbook"]["workbook"]["id"])

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
            sheet[cell_ref].set_user_input(user_input)

    workbook.workbook_json = wb.json
    workbook.save()


def _create_verified_license(
    email: str = "joe@example.com",
    domains: str = "example.com,example2.com,*.example3.com",
) -> License:
    before_license_ids = list(License.objects.values_list("id", flat=True))
    request = RequestFactory().post("/send-license-key?%s" % urlencode({"email": email, "domains": domains}))
    send_license_key(request)
    after_license_ids = License.objects.values_list("id", flat=True)
    new_license_ids = list(set(after_license_ids).difference(before_license_ids))
    assert len(new_license_ids) == 1
    license = License.objects.get(id=new_license_ids[0])
    license.email_verified = True
    license.save()
    return license


class SimpleTest(TestCase):
    def setUp(self) -> None:
        # Every test needs access to the request factory.
        self.factory = RequestFactory()

    def test_send_license_key_invalid(self) -> None:
        # email & domains missing
        request = self.factory.post("/send-license-key")
        response = send_license_key(request)
        self.assertEqual(response.status_code, 400)

        # email missing
        request = self.factory.post("/send-license-key?domains=example.com")
        response = send_license_key(request)
        self.assertEqual(response.status_code, 400)

    @override_settings(ALLOW_ONLY_EMPLOYEE_EMAILS=False)
    @override_settings(SERVER="https://www.equalto.com/serverless")
    @patch("serverless.email._send_email")
    def test_send_license_key_without_domain(self, mock_send_email: MagicMock) -> None:
        # For the beta we're not going to require that users specify domains for their
        # license. And in the situation where no domain is specified, all domains will be allowed.
        # When we move to v1.0, we'll require that users specify the domains on which they
        # want to use their license key.

        # domain missing
        request = self.factory.post("/send-license-key?%s" % urlencode({"email": "joe@example.com"}))
        response = send_license_key(request)
        self.assertEqual(response.status_code, 200)

        license = License.objects.get()
        self.assertEqual(LicenseDomain.objects.filter(license=license).count(), 0)

        self.assertFalse(license.email_verified)

        # activation email should be dispatched
        mock_send_email.assert_called_once()
        args, _ = mock_send_email.call_args
        message = args[0].get()
        self.assertEqual(
            message,
            {
                "from": {"name": "EqualTo", "email": "no-reply@equalto.com"},
                "personalizations": [
                    {
                        "to": [{"email": "joe@example.com"}],
                        "dynamic_template_data": {
                            "emailVerificationURL": (
                                f"https://www.equalto.com/serverless/#/license/activate/{license.id}"
                            ),
                        },
                    },
                ],
                "template_id": LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID,
            },
        )

        # license email not verified, so license not activate
        self.assertFalse(is_license_key_valid_for_host(license.key, "example.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "example2.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "sub.example3.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "other.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "sub.example.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, None))

        # there shouldn't be any workbooks linked to that license at this point
        self.assertFalse(Workbook.objects.filter(license=license).exists())

        # verify email address, activating license
        request = self.factory.get("/activate-license-key/%s/" % license.id)
        response = activate_license_key(request, license.id)
        self.assertEqual(response.status_code, 200)
        license.refresh_from_db()
        self.assertTrue(license.email_verified)

        self.assertTrue(Workbook.objects.filter(license=license).exists())
        workbook = Workbook.objects.get(license=license)

        self.assertEqual(
            json.loads(response.content),
            {
                "license_key": str(license.key),
                "workbook_id": str(workbook.id),
            },
        )

        # license activated, should work on all domains since no LicenseDomain records
        # are associated with the license
        self.assertTrue(is_license_key_valid_for_host(license.key, "example.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, "example2.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, "sub.example3.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, "other.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, "sub.example.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, None))

        # domain parameter specified as the empty string
        request = self.factory.post("/send-license-key?%s" % urlencode({"email": "joe2@example.com", "domains": ""}))
        response = send_license_key(request)
        self.assertEqual(response.status_code, 200)
        license2 = License.objects.get(email="joe2@example.com")
        self.assertFalse(license2.email_verified)
        self.assertEqual(LicenseDomain.objects.filter(license=license2).count(), 0)

        # calling the endpoint again should be a noop
        request = self.factory.get("/activate-license-key/%s/" % license.id)
        response = activate_license_key(request, license.id)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {
                "license_key": str(license.key),
                "workbook_id": str(workbook.id),
            },
        )

    def test_create_workbook_from_xlsx(self) -> None:
        license = _create_verified_license()
        with open("serverless/test-data/test-upload.xlsx", "rb") as test_upload_file:
            xlsx_file = SimpleUploadedFile(
                "test-upload.xlsx",
                test_upload_file.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        request = self.factory.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_ORIGIN="http://example.com",
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        response = create_workbook_from_xlsx(request)
        self.assertEqual(response.status_code, 200)

        # confirm that that workbook has been created
        data = graphql_query(
            """
            query {
                workbooks {
                    sheets {
                        id
                        name
                    }
                }
            }""",
            "example.com",
            license.key,
        )

        self.assertEqual(len(data["data"]["workbooks"]), 1)
        self.assertCountEqual(
            data["data"]["workbooks"][0]["sheets"],
            [
                {"id": 1, "name": "Sheet1"},
                {"id": 3, "name": "Second"},
                {"id": 8, "name": "Sheet4"},
                {"id": 9, "name": "shared"},
                {"id": 7, "name": "Table"},
                {"id": 2, "name": "Sheet2"},
                {"id": 4, "name": "Created fourth"},
                {"id": 5, "name": "Hidden"},
            ],
        )

    def test_create_workbook_from_xlsx_bad_license(self) -> None:
        _create_verified_license()
        with open("serverless/test-data/test-upload.xlsx", "rb") as test_upload_file:
            xlsx_file = SimpleUploadedFile(
                "test-upload.xlsx",
                test_upload_file.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        invalid_license_key = "dc6325b0-9e39-44e9-b2ca-278e14be6bc5"
        request = self.factory.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_ORIGIN="http://example.com",
            HTTP_AUTHORIZATION="Bearer %s" % invalid_license_key,
        )
        # with self.assertRaisesMessage(LicenseKeyError, "Invalid license key"):
        response = create_workbook_from_xlsx(request)
        self.assertEqual(response.status_code, 403)
        self.assertEqual(
            response.content,
            b"Invalid license",
        )
        self.assertEqual(Workbook.objects.count(), 0)

    def test_create_workbook_from_xlsx_too_large(self) -> None:
        license = _create_verified_license()
        xlsx_file = SimpleUploadedFile(
            "test-upload.xlsx",
            b" " * (MAX_XLSX_FILE_SIZE + 1),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        request = self.factory.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_ORIGIN="http://example.com",
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )

        # with self.assertRaisesMessage(LicenseKeyError, "Invalid license key"):
        response = create_workbook_from_xlsx(request)
        self.assertEqual(response.status_code, 400)
        self.assertEqual(
            response.content,
            ("Excel file too large (max size %s bytes)." % MAX_XLSX_FILE_SIZE).encode("utf-8"),
        )
        self.assertEqual(Workbook.objects.count(), 0)

    def test_create_workbook_invalid_xlsx_file(self) -> None:
        license = _create_verified_license()
        xlsx_file = SimpleUploadedFile(
            "test-upload.xlsx",
            # This is not a valid XLSX file (!)
            b" ",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        request = self.factory.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_ORIGIN="http://example.com",
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )

        response = create_workbook_from_xlsx(request)
        self.assertEqual(response.status_code, 400)
        self.assertTrue(response.content.startswith(b"Could not upload workbook."))
        self.assertEqual(Workbook.objects.count(), 0)

    def test_send_license_key(self) -> None:
        self.assertEqual(License.objects.count(), 0)

        request = self.factory.post(
            "/send-license-key?%s"
            % urlencode(
                {
                    "email": "joe@example.com",
                    "domains": "example.com,example2.com,*.example3.com",
                },
            ),
        )
        response = send_license_key(request)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(License.objects.count(), 1)
        self.assertListEqual(
            list(License.objects.values_list("email", "email_verified")),
            [("joe@example.com", False)],
        )
        license = License.objects.get()
        self.assertCountEqual(
            LicenseDomain.objects.values_list("license", "domain"),
            [
                (license.id, "example.com"),
                (license.id, "example2.com"),
                (license.id, "*.example3.com"),
            ],
        )

        # license email not verified, so license not activate
        self.assertFalse(is_license_key_valid_for_host(license.key, "example.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "example2.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "sub.example3.com:443"))
        # these aren't valid, regardless of license activation
        self.assertFalse(is_license_key_valid_for_host(license.key, "other.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "sub.example.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, None))

        # verify email address, activating license
        request = self.factory.get("/activate-license-key/%s/" % license.id)
        response = activate_license_key(request, license.id)
        self.assertEqual(response.status_code, 200)
        license.refresh_from_db()
        self.assertTrue(license.email_verified)

        # license email verified, so license not activate
        self.assertTrue(is_license_key_valid_for_host(license.key, "example.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, "example2.com:443"))
        self.assertTrue(is_license_key_valid_for_host(license.key, "sub.example3.com:443"))
        # these aren't valid, regardless of license activation
        self.assertFalse(is_license_key_valid_for_host(license.key, "other.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, "sub.example.com:443"))
        self.assertFalse(is_license_key_valid_for_host(license.key, None))

    def test_query_workbooks(self) -> None:
        license = _create_verified_license()
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license.key),
            {"data": {"workbooks": []}},
        )

        workbook = _create_workbook(license)
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license.key),
            {"data": {"workbooks": [{"id": str(workbook.id)}]}},
        )

    def test_query_workbook(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId:"%s") {
                      id
                    }
                }"""
                % workbook.id,
                "example.com",
                license.key,
            ),
            {"data": {"workbook": {"id": str(workbook.id)}}},
        )

        # bob can't access joe's workbook
        license2 = _create_verified_license(email="bob@example.com")
        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId:"%s") {
                      id
                    }
                }"""
                % workbook.id,
                "example.com",
                license2.key,
                suppress_errors=True,
            ),
            {"data": {"workbook": None}},
        )

    def test_query_workbooks_multiple_users(self) -> None:
        license1 = _create_verified_license(email="joe@example.com")
        license2 = _create_verified_license(email="joe2@example.com")
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license1.key),
            {"data": {"workbooks": []}},
        )
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license2.key),
            {"data": {"workbooks": []}},
        )

        workbook = _create_workbook(license1)
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license1.key),
            {"data": {"workbooks": [{"id": str(workbook.id)}]}},
        )
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license2.key),
            {"data": {"workbooks": []}},
        )

    def test_create_workbook(self) -> None:
        self.assertEqual(Workbook.objects.count(), 0)
        license = _create_verified_license()
        data = graphql_query(
            """
            mutation {
                create_workbook: createWorkbook {
                    workbook{revision, id, workbookJson}
                }
            }
            """,
            "example.com",
            license.key,
        )
        self.assertCountEqual(
            data["data"]["create_workbook"]["workbook"].keys(),
            ["revision", "id", "workbookJson"],
        )
        self.assertEqual(Workbook.objects.count(), 1)
        self.assertEqual(Workbook.objects.get().license, license)

    def test_create_workbook_unlicensed_domain(self) -> None:
        self.assertEqual(Workbook.objects.count(), 0)
        license = _create_verified_license()
        data = graphql_query(
            """
            mutation {
                create_workbook: createWorkbook {
                    workbook{revision, id, workbookJson}
                }
            }
            """,
            "not-licensed.com",
            license.key,
            suppress_errors=True,
        )
        self.assertEqual(data["data"]["create_workbook"], None)
        self.assertEqual(Workbook.objects.count(), 0)

    def test_create_too_many_workbooks(self) -> None:
        license = _create_verified_license()
        self.assertEqual(Workbook.objects.count(), 0)
        for _ in range(MAX_WORKBOOKS_PER_LICENSE):
            _create_workbook(license)
        self.assertEqual(Workbook.objects.count(), MAX_WORKBOOKS_PER_LICENSE)
        with self.assertRaisesMessage(
            GraphQLError,
            "You cannot create more than %s workbooks with this license key." % MAX_WORKBOOKS_PER_LICENSE,
        ):
            graphql_query(
                """
                mutation {
                    create_workbook: createWorkbook {
                        workbook{revision, id, workbookJson}
                    }
                }
                """,
                "example.com",
                license.key,
            )
        self.assertEqual(Workbook.objects.count(), MAX_WORKBOOKS_PER_LICENSE)

    def test_set_cell_input(self) -> None:
        license = _create_verified_license(email="joe@example.com")
        workbook = _create_workbook(license)

        data = graphql_query(
            """
            mutation SetCellWorkbook($workbook_id: String!) {
                setCellInput(workbookId: $workbook_id, sheetId: 1, ref: "A1", input: "$2.50") { workbook { id } }
                output: setCellInput(workbookId: $workbook_id, sheetId: 1, row: 1, col: 2, input: "=A1*2") {
                    workbook {
                        id
                        sheet(sheetId: 1) {
                            id
                            B1: cell(ref: "B1") {
                                formattedValue
                                value {
                                    number
                                }
                                formula
                            }
                        }
                    }
                }
            }
            """,
            "example.com",
            license.key,
            {"workbook_id": str(workbook.id)},
        )
        self.assertEqual(
            data["data"]["output"],
            {
                "workbook": {
                    "id": str(workbook.id),
                    "sheet": {
                        "id": 1,
                        "B1": {"formattedValue": "$5.00", "value": {"number": 5}, "formula": "=A1*2"},
                    },
                },
            },
        )

        workbook.refresh_from_db()
        self.assertEqual(workbook.calc.sheets[0]["B1"].value, 5)

        # confirm that "other" license can't modify the workbook
        license_other = _create_verified_license(email="other@example.com")
        data = graphql_query(
            """
            mutation {
                set_cell_input: setCellInput(workbookId:"%s", sheetId:1, ref: "A1", input: "100") {
                    workbook{ id }
                }
            }
            """
            % str(workbook.id),
            "example.com",
            license_other.key,
            suppress_errors=True,
        )
        self.assertIsNone(data["data"]["set_cell_input"])

    def test_set_cell_input_too_large(self) -> None:
        license = _create_verified_license(email="joe@example.com")
        workbook = _create_workbook(license)

        # this should work, we're right up to the limit
        graphql_query(
            """
            mutation SetCellInput($workbook_id: String!) {
                setCellInput(workbookId: $workbook_id, sheetId: 1, ref: "A1", input: "%s") { workbook { revision } }
            }
            """
            % (" " * (MAX_WORKBOOK_INPUT_SIZE)),
            "example.com",
            license.key,
            {"workbook_id": str(workbook.id)},
        )
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)

        with self.assertRaisesMessage(GraphQLError, "Workbook input too large"):
            graphql_query(
                """
                mutation SetCellInput($workbook_id: String!) {
                    setCellInput(workbookId: $workbook_id, sheetId: 1, ref: "A1", input: "%s") { workbook { id } }
                }
                """
                % (" " * (MAX_WORKBOOK_INPUT_SIZE + 1)),
                "example.com",
                license.key,
                {"workbook_id": str(workbook.id)},
            )
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)

    def test_save_workbook(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)

        new_json = equalto.new().json

        data = graphql_query(
            """
            mutation SaveWorkbook($workbook_json: String!) {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: $workbook_json) {
                    revision
                }
            }
            """
            % str(workbook.id),
            "example.com",
            license.key,
            {"workbook_json": new_json},
        )
        self.assertEqual(data["data"]["save_workbook"], {"revision": 2})
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)
        self.assertEqual(workbook.workbook_json, new_json)

    def test_save_workbook_invalid_json(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license)
        old_workbook_json = workbook.workbook_json

        data = graphql_query(
            """
            mutation SaveWorkbook($workbook_json: String!) {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: $workbook_json) {
                    revision
                }
            }
            """
            % str(workbook.id),
            "example.com",
            license.key,
            {"workbook_json": "not really a workbook JSON"},
            suppress_errors=True,
        )
        self.assertEqual(data["data"], {"save_workbook": None})
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)
        self.assertEqual(workbook.workbook_json, old_workbook_json)

    def test_save_workbook_json_too_large(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)

        new_workbook_data = json.loads(equalto.new().json)
        # We store a large string in metadata.application to ensure the JSON exceeds the
        # MAX_WORKBOOK_JSON_SIZE.
        new_workbook_data["metadata"]["application"] = "Very Large Name: %s" % (" " * MAX_WORKBOOK_JSON_SIZE)
        with self.assertRaisesMessage(GraphQLError, "Workbook JSON too large"):
            graphql_query(
                """
                mutation SaveWorkbook($workbook_json: String!) {
                    save_workbook: saveWorkbook(workbookId: "%s", workbookJson: $workbook_json) {
                        revision
                    }
                }
                """
                % str(workbook.id),
                "example.com",
                license.key,
                {"workbook_json": json.dumps(new_workbook_data)},
            )
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)

    def test_save_workbook_invalid_license(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)

        license2 = _create_verified_license(email="bob@example.com")
        data = graphql_query(
            """
            mutation {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: "{ }") {
                    revision
                }
            }
            """
            % str(workbook.id),
            "example.com",
            license2.key,
            suppress_errors=True,
        )
        self.assertEqual(data["data"]["save_workbook"], None)
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)

    def test_save_workbook_unlicensed_domain(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)

        data = graphql_query(
            """
            mutation {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: "{ }") {
                    revision
                }
            }
            """
            % str(workbook.id),
            "not-licensed.com",
            license.key,
            suppress_errors=True,
        )
        self.assertEqual(data["data"]["save_workbook"], None)
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)

    def test_query_workbook_sheets(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})

        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId: "%s") {
                        sheets {
                            id
                            name
                        }
                    }
                }"""
                % workbook.id,
                "example.com",
                license.key,
            ),
            {
                "data": {
                    "workbook": {
                        "sheets": [
                            {"id": 1, "name": "Calculation"},
                            {"id": 2, "name": "Data"},
                        ],
                    },
                },
            },
        )

    def test_query_workbook_sheet(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"FirstSheet": {}, "Calculation": {}, "Data": {}, "LastSheet": {}})

        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId: "%s") {
                        sheet_2: sheet(sheetId: 2) {
                            id
                            name
                        }
                        sheet_data: sheet(name: "Data") {
                            id
                            name
                        }
                    }
                }"""
                % workbook.id,
                "example.com",
                license.key,
            ),
            {
                "data": {
                    "workbook": {
                        "sheet_2": {"id": 2, "name": "Calculation"},
                        "sheet_data": {"id": 3, "name": "Data"},
                    },
                },
            },
        )

    def test_query_cell(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Sheet": {"A1": "$2.50", "A2": "foobar", "A3": "true", "A4": "=2+2*2"}})

        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId: "%s") {
                        sheet(name: "Sheet") {
                            id
                            A1: cell(ref: "A1") {
                                formattedValue
                                value {
                                    text
                                    number
                                    boolean
                                }
                                type
                                format
                                formula
                            }
                            A2: cell(row: 2, col: 1) {
                                formattedValue
                                value {
                                    text
                                    number
                                    boolean
                                }
                                type
                                format
                                formula
                            }
                            A3: cell(col: 1, row: 3) {
                                formattedValue
                                value {
                                    text
                                    number
                                    boolean
                                }
                                type
                                format
                                formula
                            }
                            A4: cell(ref: "A4") {
                                formattedValue
                                value {
                                    text
                                    number
                                    boolean
                                }
                                type
                                format
                                formula
                            }
                            empty_cell: cell(ref: "A42") {
                                formattedValue
                            }
                        }
                    }
                }"""
                % workbook.id,
                "example.com",
                license.key,
            ),
            {
                "data": {
                    "workbook": {
                        "sheet": {
                            "id": 1,
                            "A1": {
                                "formattedValue": "$2.50",
                                "value": {"text": None, "number": 2.5, "boolean": None},
                                "type": "number",
                                "format": "$#,##0.00",
                                "formula": None,
                            },
                            "A2": {
                                "formattedValue": "foobar",
                                "value": {"text": "foobar", "number": None, "boolean": None},
                                "type": "text",
                                "format": "general",
                                "formula": None,
                            },
                            "A3": {
                                "formattedValue": "TRUE",
                                "value": {"text": None, "number": None, "boolean": True},
                                "type": "logical_value",
                                "format": "general",
                                "formula": None,
                            },
                            "A4": {
                                "formattedValue": "6",
                                "value": {"text": None, "number": 6.0, "boolean": None},
                                "type": "number",
                                "format": "general",
                                "formula": "=2+2*2",
                            },
                            "empty_cell": {"formattedValue": ""},
                        },
                    },
                },
            },
        )

        # confirm that None is returned if args list is invalid
        calls = [
            'cell(ref: "A4", col: 1)',
            'cell(ref: "A5", row: 1)',
            "cell(col: 1)",
            "cell(row: 1)",
            'cell(ref: "Sheet!A1")',
            'cell(ref: "InvalidRef")',
            "cell(row: 1, col: 0)",
        ]
        for call in calls:
            self.assertEqual(
                graphql_query(
                    """
                    query {
                        workbook(workbookId: "%s") {
                            sheet(name: "Sheet") {
                                cell: %s { formattedValue }
                            }
                        }
                    }"""
                    % (workbook.id, call),
                    "example.com",
                    license.key,
                    suppress_errors=True,
                ),
                {"data": {"workbook": {"sheet": None}}},
            )

    def test_create_sheet(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        response = graphql_query(
            """
            mutation CreateSheets($workbook_id: String!) {
                createSheet(workbookId: $workbook_id) { sheet { id } }
                output: createSheet(workbookId: $workbook_id, sheetName: "Analytics") {
                    sheet {
                        id
                        name
                    }
                    workbook {
                        sheets {
                            id
                            name
                        }
                    }
                }
            }""",
            "example.com",
            license.key,
            {"workbook_id": str(workbook.id)},
        )

        self.assertEqual(
            response["data"]["output"],
            {
                "sheet": {"id": 4, "name": "Analytics"},
                "workbook": {
                    "sheets": [
                        {"id": 1, "name": "Calculation"},
                        {"id": 2, "name": "Data"},
                        {"id": 3, "name": "Sheet1"},  # new sheet using the default name
                        {"id": 4, "name": "Analytics"},
                    ],
                },
            },
        )

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 3)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Calculation", "Data", "Sheet1", "Analytics"],
        )

    def test_create_sheet_name_in_use(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Sheet": {}})
        self.assertEqual(workbook.revision, 1)

        with self.assertRaisesMessage(GraphQLError, "A worksheet already exists with that name"):
            graphql_query(
                """
                mutation CreateSheets($workbook_id: String!) {
                    createSheet(workbookId: $workbook_id, sheetName: "Sheet") {
                        sheet {
                            id
                        }
                    }
                }""",
                "example.com",
                license.key,
                {"workbook_id": str(workbook.id)},
            )

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Sheet"],
        )

    def test_create_sheet_invalid_license(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        response = graphql_query(
            """
            mutation CreateSheets($workbook_id: String!) {
                output: createSheet(workbookId: $workbook_id, sheetName: "Analytics") {
                    sheet {
                        id
                        name
                    }
                }
            }""",
            "example.com",
            _create_verified_license("bob@example.com").key,
            {"workbook_id": str(workbook.id)},
            suppress_errors=True,
        )

        self.assertIsNone(response["data"]["output"])

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Calculation", "Data"],
        )

    def test_delete_sheet(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        response = graphql_query(
            """
            mutation DeleteSheet($workbook_id: String!) {
                output: deleteSheet(workbookId: $workbook_id, sheetId: 1) {
                    workbook {
                        sheets {
                            id
                            name
                        }
                    }
                }
            }""",
            "example.com",
            license.key,
            {"workbook_id": str(workbook.id)},
        )

        self.assertEqual(
            response["data"]["output"],
            {
                "workbook": {
                    "sheets": [
                        {"id": 2, "name": "Data"},
                    ],
                },
            },
        )

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Data"],
        )

    def test_delete_sheet_invalid_license(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        response = graphql_query(
            """
            mutation DeleteSheet($workbook_id: String!) {
                output: deleteSheet(workbookId: $workbook_id, sheetId: 1) {
                    workbook {
                        sheets {
                            id
                            name
                        }
                    }
                }
            }""",
            "example.com",
            _create_verified_license("bob@example.com").key,
            {"workbook_id": str(workbook.id)},
            suppress_errors=True,
        )

        self.assertIsNone(response["data"]["output"])

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Calculation", "Data"],
        )

    def test_rename_sheet(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        response = graphql_query(
            """
            mutation RenameSheet($workbook_id: String!) {
                output: renameSheet(workbookId: $workbook_id, sheetId: 1, newName: "Analytics") {
                    sheet {
                        id
                        name
                    }
                    workbook {
                        sheets {
                            id
                            name
                        }
                    }
                }
            }""",
            "example.com",
            license.key,
            {"workbook_id": str(workbook.id)},
        )

        self.assertEqual(
            response["data"]["output"],
            {
                "sheet": {"id": 1, "name": "Analytics"},
                "workbook": {
                    "sheets": [
                        {"id": 1, "name": "Analytics"},
                        {"id": 2, "name": "Data"},
                    ],
                },
            },
        )

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Analytics", "Data"],
        )

    def test_rename_sheet_name_in_use(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        with self.assertRaisesMessage(GraphQLError, "Sheet already exists: 'Data'"):
            graphql_query(
                """
                mutation RenameSheet($workbook_id: String!) {
                    renameSheet(workbookId: $workbook_id, sheetId: 1, newName: "Data") {
                        sheet {
                            id
                        }
                    }
                }""",
                "example.com",
                license.key,
                {"workbook_id": str(workbook.id)},
            )

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Calculation", "Data"],
        )

    def test_rename_sheet_invalid_license(self) -> None:
        license = _create_verified_license()
        workbook = _create_workbook(license, {"Calculation": {}, "Data": {}})
        self.assertEqual(workbook.revision, 1)

        response = graphql_query(
            """
            mutation RenameSheet($workbook_id: String!) {
                output: renameSheet(workbookId: $workbook_id, sheetId: 1, newName: "Analytics") {
                    workbook {
                        sheets {
                            id
                            name
                        }
                    }
                }
            }""",
            "example.com",
            _create_verified_license("bob@example.com").key,
            {"workbook_id": str(workbook.id)},
            suppress_errors=True,
        )

        self.assertIsNone(response["data"]["output"])

        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)
        self.assertEqual(
            [sheet.name for sheet in workbook.calc.sheets],
            ["Calculation", "Data"],
        )


class GetUpdatedWorkbookTests(TransactionTestCase):
    def setUp(self) -> None:
        self.license = _create_verified_license()
        self.workbook = _create_workbook(self.license, {"Sheet": {"A1": "42"}})

    async def test_get_updated_workbook(self) -> None:
        self.assertEqual(self.workbook.revision, 1)

        request = AsyncClient().get(
            f"/get-updated-workbook/{self.workbook.id}/{self.workbook.revision}",
            AUTHORIZATION=f"Bearer {self.license.key}",
        )

        # let's wait a second so that the previous connection can attempt to pull the updated workbook at least once
        # (a newer version doesn't exist at this stage)
        await sleep(1)

        # update the workbook
        await sync_to_async(
            lambda: graphql_query(
                """
                mutation SetCellWorkbook($workbook_id: String!) {
                    setCellInput(workbookId: $workbook_id, sheetName: "Sheet", ref: "A1", input: "$2.50") {
                        workbook { id }
                    }
                }
                """,
                "example.com",
                self.license.key,
                {"workbook_id": str(self.workbook.id)},
            ),
        )()

        response = await request
        self.assertEqual(response.status_code, 200)

        # confirm that the response contains the new version of the workbook
        response_data = json.loads(response.content)
        self.assertEqual(response_data["revision"], 2)

        wb = equalto.loads(response_data["workbook_json"])
        self.assertEqual(str(wb["Sheet!A1"]), "$2.50")

    def test_missing_license(self) -> None:
        response = self.client.get(f"/get-updated-workbook/{self.workbook.id}/{self.workbook.revision}")
        self.assertEqual(response.status_code, 403)
        self.assertEqual(response.content, b"Invalid license")

    def test_invalid_license_for_workbook(self) -> None:
        response = self.client.get(
            f"/get-updated-workbook/{self.workbook.id}/{self.workbook.revision}",
            HTTP_AUTHORIZATION=f"Bearer {_create_verified_license(email='bob@example.com').key}",
        )
        self.assertEqual(response.status_code, 404)
        self.assertEqual(response.content, b"Requested workbook does not exist")

    def test_invalid_revision(self) -> None:
        response = self.client.get(
            f"/get-updated-workbook/{self.workbook.id}/{self.workbook.revision + 1}",
            HTTP_AUTHORIZATION=f"Bearer {self.license.key}",
        )
        self.assertEqual(response.status_code, 400)
