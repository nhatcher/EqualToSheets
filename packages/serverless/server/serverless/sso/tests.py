from typing import Any
from unittest.mock import MagicMock, patch

from django.test import TestCase, override_settings

from serverless.models import License


class MockResponse:
    def __init__(self, obj: Any, status_code: int = 200) -> None:
        self.json = lambda: obj
        self.status_code = status_code


@override_settings(  # noqa: S106
    SERVER="https://sheets.equalto.com",
    GITHUB_SSO_CLIENT_ID="github-client-id",
    GITHUB_SSO_CLIENT_SECRET="github-client-secret",
    GOOGLE_SSO_CLIENT_ID="google-client-id",
    GOOGLE_SSO_CLIENT_SECRET="google-client-secret",
    MICROSOFT_SSO_CLIENT_ID="microsoft-client-id",
    MICROSOFT_SSO_CLIENT_SECRET="microsoft-client-secret",
)
class TestSSO(TestCase):
    def test_github_login(self) -> None:
        response = self.client.get("/sso/github/login")
        self.assertRedirects(
            response,
            (
                "https://github.com/login/oauth/authorize"
                + "?response_type=code"
                + "&client_id=github-client-id"
                + "&redirect_uri=https%3A%2F%2Fsheets.equalto.com%2Fsso%2Fgithub%2Fcallback"
                + "&scope=user%3Aemail"
            ),
            fetch_redirect_response=False,
        )

    def test_google_login(self) -> None:
        response = self.client.get("/sso/google/login")
        self.assertRedirects(
            response,
            (
                "https://accounts.google.com/o/oauth2/auth"
                + "?response_type=code"
                + "&client_id=google-client-id"
                + "&redirect_uri=https%3A%2F%2Fsheets.equalto.com%2Fsso%2Fgoogle%2Fcallback"
                + "&scope=profile+email"
            ),
            fetch_redirect_response=False,
        )

    def test_microsoft_login(self) -> None:
        response = self.client.get("/sso/microsoft/login")
        self.assertRedirects(
            response,
            (
                "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
                + "?response_type=code"
                + "&client_id=microsoft-client-id"
                + "&redirect_uri=https%3A%2F%2Fsheets.equalto.com%2Fsso%2Fmicrosoft%2Fcallback"
                + "&scope=User.Read+openid"
            ),
            fetch_redirect_response=False,
        )

    @patch(
        "serverless.sso.oauth.session.get",
        return_value=MockResponse(
            [
                {"email": "eve@example.org", "primary": False},
                {"email": "alice@example.org", "primary": True},
            ],
        ),
    )
    @patch("serverless.sso.oauth.session.post", return_value=MockResponse({"access_token": "access-token"}))
    def test_github_callback(self, mock_post: MagicMock, mock_get: MagicMock) -> None:
        response = self.client.get("/sso/github/callback", {"code": "auth-code"})

        license = License.objects.get(email="alice@example.org")

        self.assertRedirects(
            response,
            f"https://sheets.equalto.com/#/license/activate/{license.id}",
            fetch_redirect_response=False,
        )

        mock_post.assert_called_once_with(  # POST request is used to retrieve the access token
            url="https://github.com/login/oauth/access_token",  # noqa: S105
            data={
                "grant_type": "authorization_code",
                "code": "auth-code",
                "client_id": "github-client-id",
                "client_secret": "github-client-secret",
                "redirect_uri": "https://sheets.equalto.com/sso/github/callback",
            },
            headers={
                "Content-type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
            },
        )
        mock_get.assert_called_once_with(  # GET request is used to retrieve the user's email address
            url="https://api.github.com/user/emails",
            headers={
                "Content-Type": "application/json",
                "Authorization": "Bearer access-token",
                "X-PrettyPrint": "1",
            },
        )

        # requesting the callback the second time should redirect to the same license activation page
        self.assertRedirects(
            self.client.get("/sso/github/callback", {"code": "auth-code"}),
            f"https://sheets.equalto.com/#/license/activate/{license.id}",
            fetch_redirect_response=False,
        )

    @patch("serverless.sso.oauth.session.get", return_value=MockResponse({"email": "alice@example.org"}))
    @patch("serverless.sso.oauth.session.post", return_value=MockResponse({"access_token": "access-token"}))
    def test_google_callback(self, mock_post: MagicMock, mock_get: MagicMock) -> None:
        response = self.client.get("/sso/google/callback", {"code": "auth-code"})

        license = License.objects.get(email="alice@example.org")

        self.assertRedirects(
            response,
            f"https://sheets.equalto.com/#/license/activate/{license.id}",
            fetch_redirect_response=False,
        )

        mock_post.assert_called_once_with(  # POST request is used to retrieve the access token
            url="https://accounts.google.com/o/oauth2/token",  # noqa: S105
            data={
                "grant_type": "authorization_code",
                "code": "auth-code",
                "client_id": "google-client-id",
                "client_secret": "google-client-secret",
                "redirect_uri": "https://sheets.equalto.com/sso/google/callback",
            },
            headers={
                "Content-type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
            },
        )
        mock_get.assert_called_once_with(  # GET request is used to retrieve the user's email address
            url="https://www.googleapis.com/oauth2/v1/userinfo",
            headers={
                "Content-Type": "application/json",
                "Authorization": "Bearer access-token",
                "X-PrettyPrint": "1",
            },
        )

        # requesting the callback the second time should redirect to the same license activation page
        self.assertRedirects(
            self.client.get("/sso/google/callback", {"code": "auth-code"}),
            f"https://sheets.equalto.com/#/license/activate/{license.id}",
            fetch_redirect_response=False,
        )

    @patch("serverless.sso.oauth.session.get", return_value=MockResponse({"userPrincipalName": "alice@example.org"}))
    @patch("serverless.sso.oauth.session.post", return_value=MockResponse({"access_token": "access-token"}))
    def test_microsoft_callback(self, mock_post: MagicMock, mock_get: MagicMock) -> None:
        response = self.client.get("/sso/microsoft/callback", {"code": "auth-code"})

        license = License.objects.get(email="alice@example.org")

        self.assertRedirects(
            response,
            f"https://sheets.equalto.com/#/license/activate/{license.id}",
            fetch_redirect_response=False,
        )

        mock_post.assert_called_once_with(  # POST request is used to retrieve the access token
            url="https://login.microsoftonline.com/common/oauth2/v2.0/token",  # noqa: S105
            data={
                "grant_type": "authorization_code",
                "code": "auth-code",
                "client_id": "microsoft-client-id",
                "client_secret": "microsoft-client-secret",
                "redirect_uri": "https://sheets.equalto.com/sso/microsoft/callback",
            },
            headers={
                "Content-type": "application/x-www-form-urlencoded",
                "Accept": "application/json",
            },
        )
        mock_get.assert_called_once_with(  # GET request is used to retrieve the user's email address
            url="https://graph.microsoft.com/v1.0/me",
            headers={
                "Content-Type": "application/json",
                "Authorization": "Bearer access-token",
                "X-PrettyPrint": "1",
            },
        )

        # requesting the callback the second time should redirect to the same license activation page
        self.assertRedirects(
            self.client.get("/sso/microsoft/callback", {"code": "auth-code"}),
            f"https://sheets.equalto.com/#/license/activate/{license.id}",
            fetch_redirect_response=False,
        )

    def test_invalid_login_idp(self) -> None:
        self.assertEqual(
            self.client.get("/sso/facebook/login").status_code,
            404,
        )

    def test_invalid_callback_idp(self) -> None:
        self.assertEqual(
            self.client.get("/sso/facebook/callback", {"code": "auth-code"}).status_code,
            404,
        )

    def test_callback_error(self) -> None:
        self.assertEqual(
            self.client.get("/sso/google/callback", {"error": 42}).content,
            b"Authorization has not been approved.",
        )

    def test_callback_missing_code(self) -> None:
        self.assertEqual(
            self.client.get("/sso/google/callback").content,
            b"Missing authorization code.",
        )
