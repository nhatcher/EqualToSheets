import json
import tempfile
from asyncio import sleep
from collections import namedtuple
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch
from urllib.parse import quote

import equalto
from asgiref.sync import sync_to_async
from django.core.files.uploadedfile import SimpleUploadedFile
from django.db import transaction
from django.http import HttpResponse
from django.test import AsyncClient, RequestFactory, TestCase, TransactionTestCase, override_settings
from django.utils.http import urlencode
from equalto.exceptions import SuppressEvaluationErrors
from graphql import GraphQLError

from serverless.email import LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID
from serverless.log import info
from serverless.models import License, LicenseDomain, UnsubscribedEmail, Workbook
from serverless.schema import MAX_WORKBOOK_INPUT_SIZE, MAX_WORKBOOK_JSON_SIZE, MAX_WORKBOOKS_PER_LICENSE, schema
from serverless.test_util import create_verified_license, create_workbook
from serverless.types import SimulateInputType, SimulateOutputType
from serverless.util import get_name_from_path, is_license_key_valid_for_host
from serverless.views import (
    MAX_XLSX_FILE_SIZE,
    activate_license_key,
    create_workbook_from_xlsx,
    send_license_key,
    unsubscribe_email,
)


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
        self.assertEqual(response.status_code, 201)

        license = License.objects.get()
        self.assertEqual(LicenseDomain.objects.filter(license=license).count(), 0)

        self.assertFalse(license.email_verified)
        self.assertIsNone(license.validated_datetime)
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
        self.assertIsNotNone(license.validated_datetime)

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
        self.assertEqual(response.status_code, 201)
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

        # requesting a new license with the same email address should result in the activation email being resent,
        # without creating any new licenses
        mock_send_email.reset_mock()
        license_count = License.objects.count()
        response = self.client.post("/send-license-key", {"email": "joe@example.com"})
        self.assertEqual(response.status_code, 201)
        self.assertEqual(License.objects.count(), license_count)
        mock_send_email.assert_called_once()

    def test_create_workbook_from_xlsx(self) -> None:
        license = create_verified_license()
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
                    name
                    id
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
        self.assertEqual(
            data["data"]["workbooks"][0]["name"],
            "test-upload.xlsx",
        )

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
        create_verified_license()
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
        license = create_verified_license()
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
        license = create_verified_license()
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

    def test_create_workbook_unsupported_function(self) -> None:
        license = create_verified_license(domains="")

        calc = equalto.new()
        with SuppressEvaluationErrors():
            calc["Sheet1!A1"].formula = "=NOTTODAY()"

        with tempfile.TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "file.xlsx"
            calc.save(str(path))
            xlsx_file = SimpleUploadedFile(
                "test-upload.xlsx",
                path.read_bytes(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        response = self.client.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_AUTHORIZATION=f"Bearer {license.key}",
        )
        self.assertEqual(response.status_code, 200)

        workbook = Workbook.objects.get(license=license)
        cell = workbook.calc["Sheet1!A1"]
        self.assertEqual(cell.formula, "=NOTTODAY()")
        self.assertEqual(cell.value, "#ERROR!")

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
        self.assertEqual(response.status_code, 201)
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
        license = create_verified_license()
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license.key),
            {"data": {"workbooks": []}},
        )

        workbook = create_workbook(license)
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license.key),
            {"data": {"workbooks": [{"id": str(workbook.id)}]}},
        )

    def test_query_workbook(self) -> None:
        license = create_verified_license()
        workbook = create_workbook(license)
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
        license2 = create_verified_license(email="bob@example.com")
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
        license1 = create_verified_license(email="joe@example.com")
        license2 = create_verified_license(email="joe2@example.com")
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license1.key),
            {"data": {"workbooks": []}},
        )
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license2.key),
            {"data": {"workbooks": []}},
        )

        workbook = create_workbook(license1)
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
        license = create_verified_license()
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
        license = create_verified_license()
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
        license = create_verified_license()
        self.assertEqual(Workbook.objects.count(), 0)
        Workbook.objects.bulk_create(Workbook(license=license) for _ in range(MAX_WORKBOOKS_PER_LICENSE))
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
        license = create_verified_license(email="joe@example.com")
        workbook = create_workbook(license)

        data = graphql_query(
            """
            mutation SetCellWorkbook($workbook_id: String!) {
                B2: setCellInput(workbookId: $workbook_id, sheetId: 1, ref: "B2", input: "=B2") { workbook { id } }
                A1: setCellInput(workbookId: $workbook_id, sheetId: 1, ref: "A1", input: "$2.50") { workbook { id } }
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
                            B2: cell(ref: "B2") {
                                formattedValue
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
                        "B2": {"formattedValue": "#CIRC!"},
                    },
                },
            },
        )

        workbook.refresh_from_db()
        self.assertEqual(workbook.calc.sheets[0]["B1"].value, 5)

        # confirm that "other" license can't modify the workbook
        license_other = create_verified_license(email="other@example.com")
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
        license = create_verified_license(email="joe@example.com")
        workbook = create_workbook(license)

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
        license = create_verified_license()
        workbook = create_workbook(license)
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
        license = create_verified_license()
        workbook = create_workbook(license)
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

    def test_save_workbook_unsupported_function(self) -> None:
        license = create_verified_license()
        workbook = create_workbook(license)

        calc = equalto.new()
        with SuppressEvaluationErrors():
            calc["Sheet1!A1"].formula = "=NOTTODAY()"
        new_json = calc.json

        response = graphql_query(
            """
            mutation SaveWorkbook($workbook_id: String!, $workbook_json: String!) {
                save_workbook: saveWorkbook(workbookId: $workbook_id, workbookJson: $workbook_json) {
                    revision
                }
            }
            """,
            "example.com",
            license.key,
            {"workbook_id": str(workbook.id), "workbook_json": new_json},
        )
        self.assertEqual(response["data"]["save_workbook"], {"revision": 2})
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)
        self.assertEqual(workbook.workbook_json, new_json)

    def test_save_workbook_json_too_large(self) -> None:
        license = create_verified_license()
        workbook = create_workbook(license)
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
        license = create_verified_license()
        workbook = create_workbook(license)
        self.assertEqual(workbook.revision, 1)

        license2 = create_verified_license(email="bob@example.com")
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
        license = create_verified_license()
        workbook = create_workbook(license)
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})

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
        license = create_verified_license()
        workbook = create_workbook(license, {"FirstSheet": {}, "Calculation": {}, "Data": {}, "LastSheet": {}})

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
        license = create_verified_license()
        workbook = create_workbook(license, {"Sheet": {"A1": "$2.50", "A2": "foobar", "A3": "true", "A4": "=2+2*2"}})

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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Sheet": {}})
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
            create_verified_license("bob@example.com").key,
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
            create_verified_license("bob@example.com").key,
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
        license = create_verified_license()
        workbook = create_workbook(license, {"Calculation": {}, "Data": {}})
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
            create_verified_license("bob@example.com").key,
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

    def test_simulate(self) -> None:
        # first, we upload the workbook
        license = create_verified_license(domains="")
        with open("serverless/test-data/simulate-example.xlsx", "rb") as simulate_example:
            xlsx_file = SimpleUploadedFile(
                "simulate-example.xlsx",
                simulate_example.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        request = self.factory.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        response = create_workbook_from_xlsx(request)
        self.assertEqual(response.status_code, 200)

        workbook = Workbook.objects.filter(license=license).get()
        # then, we execute the simulations

        # simulate noop
        inputs: SimulateInputType = {}
        outputs: SimulateOutputType = {}
        response = self.client.get(
            f"/api/v1/workbooks/{workbook.id}/simulate?inputs={json.dumps(inputs)}&outputs={json.dumps(outputs)}",
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(json.loads(response.content), {})

        # simulate changes where everything computes without error
        inputs = {
            "Sheet1": {
                "A1": 55,
                "A2": 30,
                "A3": 0.08,
                "A4": "xyz",
                "A5": 500,
                "A6": False,
            },
        }
        outputs = {
            "Sheet1": ["B1", "B2", "B3", "B4", "B5", "B6", "B500"],
        }
        response = self.client.get(
            f"/api/v1/workbooks/{workbook.id}/simulate?inputs={json.dumps(inputs)}&outputs={json.dumps(outputs)}",
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content.decode("utf-8")),
            {
                "Sheet1": {
                    "B1": 56.0,
                    "B2": 60,
                    "B3": 0.04,
                    "B4": "xyzpqr",
                    "B5": 500,
                    "B6": False,
                    "B500": "",
                },
            },
        )

        # confirm that ranges are supported
        inputs = {
            "Sheet1": {
                "A1:A6": [
                    [55],
                    [30],
                    [0.08],
                    ["xyz"],
                    [500],
                    [False],
                ],
            },
        }
        outputs = {
            "Sheet1": ["A1:B6", "W1:Z1"],
        }
        response = self.client.get(
            f"/api/v1/workbooks/{workbook.id}/simulate",
            {"inputs": json.dumps(inputs), "outputs": json.dumps(outputs)},
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content.decode("utf-8")),
            {
                "Sheet1": {
                    "A1:B6": [
                        [55, 56],
                        [30, 60],
                        [0.08, 0.04],
                        ["xyz", "xyzpqr"],
                        [500, 500],
                        [False, False],
                    ],
                    "W1:Z1": [
                        ["", "", "", ""],
                    ],
                },
            },
        )

        # simulate using POST
        inputs = {"Sheet1": {"A1": 99.0}}
        outputs = {
            "Sheet1": ["B1"],
        }
        response = self.client.post(
            f"/api/v1/workbooks/{workbook.id}/simulate",
            {"inputs": json.dumps(inputs), "outputs": json.dumps(outputs)},
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(json.loads(response.content), {"Sheet1": {"B1": 100.0}})

        # simulate changes where lots of errors occur
        inputs = {
            "Sheet1": {
                "A1": "abc",
                "A2": "def",
                "A3": "ghi",
                "A4": 89,
                "A5": "jkl",
                "A6": "mno",
            },
        }
        outputs = {
            "Sheet1": ["B1", "B2", "B3", "B4", "B5", "B6", "B500"],
        }
        response = self.client.get(
            f"/api/v1/workbooks/{workbook.id}/simulate",
            {"inputs": json.dumps(inputs), "outputs": json.dumps(outputs)},
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content.decode("utf-8")),
            {
                "Sheet1": {
                    "B1": "#VALUE!",
                    "B2": "#VALUE!",
                    "B3": "#VALUE!",
                    "B4": "89pqr",
                    "B5": "jkl",
                    "B6": "mno",
                    "B500": "",
                },
            },
        )

    def test_simulate_data_validation(self) -> None:
        license = create_verified_license(domains="")
        workbook = create_workbook(license)

        def simulate(inputs: SimulateInputType, outputs: SimulateOutputType) -> HttpResponse:
            return self.client.get(
                f"/api/v1/workbooks/{workbook.id}/simulate",
                {"inputs": json.dumps(inputs), "outputs": json.dumps(outputs)},
                HTTP_AUTHORIZATION="Bearer %s" % license.key,
            )

        response = simulate(inputs={"NonExistent": {"A1": 1}}, outputs={})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b'"NonExistent" sheet does not exist')

        response = simulate(inputs={"Sheet1": {"1A": 1}}, outputs={})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b'"1A" reference cannot be parsed')

        response = simulate(inputs={"Sheet1": {"A1": [[1]]}}, outputs={})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b"[[1]] is not a valid value for A1")

        response = simulate(inputs={"Sheet1": {"B1:1B": [[1]]}}, outputs={})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b'"B1:1B" reference cannot be parsed')

        response = simulate(inputs={"Sheet1": {"A1:B2": 1}}, outputs={})  # not a list of lists
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b"1 is not a valid value for A1:B2")

        response = simulate(inputs={"Sheet1": {"A1:B2": [1]}}, outputs={})  # type: ignore  # not a list of lists
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b"[1] is not a valid value for A1:B2")

        response = simulate(inputs={"Sheet1": {"A1:B2": [[1]]}}, outputs={})  # invalid number of rows
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b"[[1]] is not a valid value for A1:B2")

        response = simulate(inputs={"Sheet1": {"A1:B2": [[1], [1]]}}, outputs={})  # invalid number of columns per row
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b"[[1], [1]] is not a valid value for A1:B2")

        response = simulate(inputs={}, outputs={"NonExistent": ["A1"]})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b'"NonExistent" sheet does not exist')

        response = simulate(inputs={}, outputs={"Sheet1": ["1A"]})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b'"1A" reference cannot be parsed')

        response = simulate(inputs={}, outputs={"Sheet1": ["B1:1B"]})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.content, b'"B1:1B" reference cannot be parsed')

    def test_simulate_unsupported_function(self) -> None:
        license = create_verified_license(domains="")
        with SuppressEvaluationErrors():
            workbook = create_workbook(license, {"Sheet1": {"A1": "=NOTTODAY()"}})
        response = self.client.get(
            f"/api/v1/workbooks/{workbook.id}/simulate",
            {
                "inputs": json.dumps({"Sheet1": {"B1": 42}}),
                "outputs": json.dumps({"Sheet1": ["A1", "B1"]}),
            },
            HTTP_AUTHORIZATION=f"Bearer {license.key}",
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            json.loads(response.content),
            {"Sheet1": {"A1": "#ERROR!", "B1": 42}},
        )

    def test_simulate_invalid(self) -> None:
        # first, we upload the workbook
        license = create_verified_license(domains="")
        with open("serverless/test-data/simulate-example.xlsx", "rb") as simulate_example:
            xlsx_file = SimpleUploadedFile(
                "simulate-example.xlsx",
                simulate_example.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        request = self.factory.post(
            "/create-workbook-from-xlsx",
            {"xlsx-file": xlsx_file},
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        response = create_workbook_from_xlsx(request)
        self.assertEqual(response.status_code, 200)

        workbook = Workbook.objects.filter(license=license).get()

        # confirm that invalid workbook ids and license keys are rejected
        inputs: SimulateInputType = {}
        outputs: SimulateOutputType = {}
        invalid_workbook_id = "dc6325b0-9e39-44e9-b2ca-278e14be6bc5"
        invalid_license_key = invalid_workbook_id
        response = self.client.get(
            f"/api/v1/workbooks/{workbook.id}/simulate",
            {"inputs": json.dumps(inputs), "outputs": json.dumps(outputs)},
            HTTP_AUTHORIZATION="Bearer %s" % invalid_license_key,
        )
        self.assertEqual(response.status_code, 403)
        self.assertEqual(response.content, b"Invalid license")

        response = self.client.get(
            f"/api/v1/workbooks/{invalid_workbook_id}/simulate",
            {"inputs": json.dumps(inputs), "outputs": json.dumps(outputs)},
            HTTP_AUTHORIZATION="Bearer %s" % license.key,
        )
        self.assertEqual(response.status_code, 404)

    def test_unsubscribe_email(self) -> None:
        email = "test€+234@example.com"
        request = self.factory.get(f"/unsubscribe-email?email={quote(email)}")
        response = unsubscribe_email(request)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.content,
            f"The email address {email} has unsubscribed from all EqualTo mailings.".encode("utf-8"),
        )

        self.assertEqual(UnsubscribedEmail.objects.filter(email=email).count(), 1)
        # repeated unsubscription
        request = self.factory.get(f"/unsubscribe-email?email={quote(email)}")
        response = unsubscribe_email(request)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.content,
            f"The email address {email} has already been unsubscribed from all EqualTo mailings.".encode("utf-8"),
        )
        self.assertEqual(UnsubscribedEmail.objects.filter(email=email).count(), 1)


class GetUpdatedWorkbookTests(TransactionTestCase):
    def setUp(self) -> None:
        self.license = create_verified_license()
        self.workbook = create_workbook(self.license, {"Sheet": {"A1": "42"}})

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
            HTTP_AUTHORIZATION=f"Bearer {create_verified_license(email='bob@example.com').key}",
        )
        self.assertEqual(response.status_code, 404)
        self.assertEqual(response.content, b"Requested workbook does not exist")

    def test_invalid_revision(self) -> None:
        response = self.client.get(
            f"/get-updated-workbook/{self.workbook.id}/{self.workbook.revision + 1}",
            HTTP_AUTHORIZATION=f"Bearer {self.license.key}",
        )
        self.assertEqual(response.status_code, 400)

    def test_edit_workbook(self) -> None:
        response = self.client.get(
            f"/unsafe-just-for-beta/edit-workbook/{self.license.key}/{self.workbook.id}/",
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.content.startswith(b"<!doctype html>"))
        self.assertTrue(f'"{self.license.key}"'.encode("utf-8") in response.content)
        self.assertTrue(f'"{self.workbook.id}"'.encode("utf-8") in response.content)

    def test_edit_workbook_invalid_license_id(self) -> None:
        invalid_license_key = "dc6325b0-9e39-44e9-b2ca-278e14be6bc5"
        response = self.client.get(
            f"/unsafe-just-for-beta/edit-workbook/{invalid_license_key}/{self.workbook.id}/",
        )
        self.assertEqual(response.status_code, 404)

    def test_edit_workbook_invalid_workbook_id(self) -> None:
        invalid_workbook_id = "dc6325b0-9e39-44e9-b2ca-278e14be6bc5"
        response = self.client.get(
            f"/unsafe-just-for-beta/edit-workbook/{self.license.key}/{invalid_workbook_id}/",
        )
        self.assertEqual(response.status_code, 404)

    def test_get_name_from_path(self) -> None:
        self.assertEqual(get_name_from_path("/path/to/file.xlsx"), "file.xlsx")
        self.assertEqual(get_name_from_path(r"c:\path\to\file.xlsx"), "file.xlsx")
        self.assertEqual(get_name_from_path("/path/to€/file€.xlsx"), "file€.xlsx")
        self.assertEqual(get_name_from_path(""), "")
