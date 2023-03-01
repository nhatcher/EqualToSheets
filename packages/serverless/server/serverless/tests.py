
from collections import namedtuple

from django.test import RequestFactory, TestCase
from django.utils.http import urlencode

from .log import error, info
from .models import License, LicenseDomain, Workbook
from .schema import schema
from .util import is_license_key_valid_for_host
from .views import activate_license_key, send_license_key


def graphql_query(query, origin, license_key=None):
    info("graphql_query(): query=%s"%query)
    context = namedtuple("context", ["META"])
    context.META = {"HTTP_ORIGIN": origin}
    if license_key is not None:
        context.META["HTTP_AUTHORIZATION"] = "Bearer %s"%str(license_key)

    info("graphql_query(): context.META=%s"%context.META)
    graphql_results = schema.execute(query, context_value=context, variable_values=None)
    return {"data": graphql_results.data}


def _create_workbook(license) -> Workbook:
    data = graphql_query(
        """
        mutation {
            create_workbook: createWorkbook {
                workbook{id}
            }
        }
        """,
        "example.com",
        license.key
    )
    return Workbook.objects.get(id=data["data"]["create_workbook"]["workbook"]["id"])
class SimpleTest(TestCase):
    def setUp(self):
        # Every test needs access to the request factory.
        self.factory = RequestFactory()

    def test_send_license_key_invalid(self):
        # email & domains missing
        request = self.factory.get("/send-license-key")
        response = send_license_key(request)
        self.assertEqual(response.status_code, 400)

        # email missing
        request = self.factory.get("/send-license-key?domains=example.com")
        response = send_license_key(request)
        self.assertEqual(response.status_code, 400)

        # domain missing
        request = self.factory.get("/send-license-key?%s"%urlencode({"email": "joe@example.com"}))
        response = send_license_key(request)
        self.assertEqual(response.status_code, 400)

        # domain missing
        request = self.factory.get("/send-license-key?%s"%urlencode({"email": "joe@example.com", "domains": ""}))
        response = send_license_key(request)
        self.assertEqual(response.status_code, 400)


    def test_send_license_key(self):
        self.assertEqual(License.objects.count(), 0)

        request = self.factory.get("/send-license-key?%s"%urlencode({"email": "joe@example.com", "domains": "example.com,example2.com,*.example3.com"}))
        response = send_license_key(request, _send_email=False)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(License.objects.count(), 1)
        self.assertListEqual(
            list(License.objects.values_list("email", "email_verified")),
            [
                ("joe@example.com", False)
            ]
        )
        license = License.objects.get()
        self.assertCountEqual(
            LicenseDomain.objects.values_list("license", "domain"),
            [
                (license.id, "example.com"),
                (license.id, "example2.com"),
                (license.id, "*.example3.com"),
            ]
        )

        # license email not verified, so license not activate
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "example.com:443")
        )
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "example2.com:443")
        )
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "sub.example3.com:443")
        )
        # these aren't valid, regardless of license activation
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "other.com:443")
        )
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "sub.example.com:443")
        )

        # verify email address, activating license
        request = self.factory.get("/activate-license-key/%s/"%license.id)
        response = activate_license_key(request, license.id)
        self.assertEqual(response.status_code, 200)
        license.refresh_from_db()
        self.assertTrue(license.email_verified)

        # license email verified, so license not activate
        self.assertTrue(
            is_license_key_valid_for_host(license.key, "example.com:443")
        )
        self.assertTrue(
            is_license_key_valid_for_host(license.key, "example2.com:443")
        )
        self.assertTrue(
            is_license_key_valid_for_host(license.key, "sub.example3.com:443")
        )
        # these aren't valid, regardless of license activation
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "other.com:443")
        )
        self.assertFalse(
            is_license_key_valid_for_host(license.key, "sub.example.com:443")
        )

    def _create_verified_license(self, email="joe@example.com", domains="example.com,example2.com,*.example3.com"):
        before_license_ids = list(License.objects.values_list("id", flat=True))
        request = self.factory.get("/send-license-key?%s"%urlencode({"email": email, "domains": domains}))
        response = send_license_key(request, _send_email=False)
        after_license_ids = License.objects.values_list("id", flat=True)
        new_license_ids = list(set(after_license_ids).difference(before_license_ids))
        assert len(new_license_ids) == 1
        license = License.objects.get(id=new_license_ids[0])
        license.email_verified = True
        license.save()
        return license

    def test_query_workbooks(self):
        license = self._create_verified_license()
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license.key),
            {"data": {"workbooks":[]}}
        )

        workbook = _create_workbook(license)
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license.key),
            {"data": {"workbooks":[{"id": str(workbook.id)}]}}
        )


    def test_query_workbook(self):
        license = self._create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId:"%s") {
                      id
                    }
                }"""%workbook.id,
                "example.com",
                license.key
            ),
            {"data": {"workbook":{"id": str(workbook.id)}}}
        )

        # bob can't access joe's workbook
        license2 = self._create_verified_license(email="bob@example.com")
        self.assertEqual(
            graphql_query(
                """
                query {
                    workbook(workbookId:"%s") {
                      id
                    }
                }"""%workbook.id,
                "example.com",
                license2.key
            ),
            {"data": {"workbook":None}}
        )

    def test_query_workbooks_multiple_users(self):
        license1 = self._create_verified_license(email="joe@example.com")
        license2 = self._create_verified_license(email="joe2@example.com")
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license1.key),
            {"data": {"workbooks":[]}}
        )
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license2.key),
            {"data": {"workbooks":[]}}
        )


        workbook = _create_workbook(license1)
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license1.key),
            {"data": {"workbooks":[{"id": str(workbook.id)}]}}
        )
        self.assertEqual(
            graphql_query("query {workbooks{id}}", "example.com", license2.key),
            {"data": {"workbooks":[]}}
        )

    def test_create_workbook(self):
        self.assertEqual(Workbook.objects.count(), 0)
        license = self._create_verified_license()
        data = graphql_query(
            """
            mutation {
                create_workbook: createWorkbook {
                    workbook{revision, id, workbookJson}
                }
            }
            """,
            "example.com",
            license.key
        )
        self.assertCountEqual(
            data["data"]["create_workbook"]["workbook"].keys(),
            ["revision", "id", "workbookJson"]
        )
        self.assertEqual(Workbook.objects.count(), 1)
        self.assertEqual(Workbook.objects.get().license, license)

    def test_create_workbook_unlicensed_domain(self):
        self.assertEqual(Workbook.objects.count(), 0)
        license = self._create_verified_license()
        data = graphql_query(
            """
            mutation {
                create_workbook: createWorkbook {
                    workbook{revision, id, workbookJson}
                }
            }
            """,
            "not-licensed.com",
            license.key
        )
        self.assertEqual(
            data["data"]["create_workbook"],
            None
        )
        self.assertEqual(Workbook.objects.count(), 0)


    def test_set_cell_input(self):
        license = self._create_verified_license(email="joe@example.com")
        workbook = _create_workbook(license)

        # TODO: confirm that workbook updated and recomputed
        data = graphql_query(
            """
            mutation {
                set_cell_input: setCellInput(workbookId:"%s", sheetId:"sheet-id-1", ref: "A1", input: "100") {
                    workbook{ id }
                }
            }
            """%str(workbook.id),
            "example.com",
            license.key
        )
        self.assertEqual(
            data["data"]["set_cell_input"]["workbook"]["id"], str(workbook.id)
        )

        # confirm that "other" license can't modify the workbook
        license_other = self._create_verified_license(email="other@example.com")
        data = graphql_query(
            """
            mutation {
                set_cell_input: setCellInput(workbookId:"%s", sheetId:"sheet-id-1", ref: "A1", input: "100") {
                    workbook{ id }
                }
            }
            """%str(workbook.id),
            "example.com",
            license_other.key
        )
        self.assertIsNone(data["data"]["set_cell_input"])

    def test_save_workbook(self):
        license = self._create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)
        # TODO: we should set valid workbook JSON, invalid JSON should be rejected
        data = graphql_query(
            """
            mutation {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: "{ }") {
                    revision
                }
            }
            """%str(workbook.id),
            "example.com",
            license.key
        )
        self.assertEqual(
            data["data"]["save_workbook"],
            {"revision": 2}
        )
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 2)
        self.assertEqual(workbook.workbook_json, "{ }")

    def test_save_workbook_invalid_license(self):
        license = self._create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)
        
        license2 = self._create_verified_license(email="bob@example.com")
        data = graphql_query(
            """
            mutation {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: "{ }") {
                    revision
                }
            }
            """%str(workbook.id),
            "example.com",
            license2.key
        )
        self.assertEqual(
            data["data"]["save_workbook"],
            None
        )
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)

    def test_save_workbook_unlicensed_domain(self):
        license = self._create_verified_license()
        workbook = _create_workbook(license)
        self.assertEqual(workbook.revision, 1)

        data = graphql_query(
            """
            mutation {
                save_workbook: saveWorkbook(workbookId: "%s", workbookJson: "{ }") {
                    revision
                }
            }
            """%str(workbook.id),
            "not-licensed.com",
            license.key
        )
        self.assertEqual(
            data["data"]["save_workbook"],
            None
        )
        workbook.refresh_from_db()
        self.assertEqual(workbook.revision, 1)

