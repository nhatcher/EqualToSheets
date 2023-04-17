from __future__ import annotations

import enum
from abc import ABC, abstractmethod

import requests
from django.conf import settings
from django.utils.http import urlencode
from requests.adapters import HTTPAdapter
from requests.exceptions import RequestException

from serverless.sso.exceptions import SSOError

session = requests.Session()
adapter = HTTPAdapter(max_retries=2)
session.mount("http://", adapter)
session.mount("https://", adapter)


class IdentityProvider(enum.Enum):
    github = "github"
    google = "google"
    microsoft = "microsoft"


class OAuthSSO(ABC):
    """Abstract class for OAuth SSO integrations."""

    identity_provider: IdentityProvider

    scopes: list[str]

    email_attribute: str | None = None

    authorize_url: str
    token_url: str
    email_api_url: str

    @property
    @abstractmethod
    def client_id(self) -> str:
        pass

    @property
    @abstractmethod
    def client_secret(self) -> str:
        pass

    @property
    def callback_url(self) -> str:
        return f"{settings.SERVER}/sso/{self.identity_provider.value}/callback"

    def extract_email(self, response: requests.Response) -> str:
        """Extract user email address from the API response."""
        assert self.email_attribute

        response_data = response.json()

        email = response_data.get(self.email_attribute)
        if not email:
            raise SSOError(f"Invalid response data: {response_data}")
        return email

    @property
    def login_url(self) -> str:
        """Returns URL used to initialize the authentication process."""
        query = {
            "response_type": "code",
            "client_id": self.client_id,
            "redirect_uri": self.callback_url,
            "scope": " ".join(self.scopes),
        }
        return f"{self.authorize_url}?{urlencode(query)}"

    def retrieve_user_email(self, authorization_code: str) -> str:
        """
        Retrieve the user's email address.

        This method performs external API calls.
        """
        access_token = self._retrieve_access_token(authorization_code)

        try:
            with session:
                response = session.get(
                    url=self.email_api_url,
                    headers={
                        "Content-Type": "application/json",
                        "Authorization": f"Bearer {access_token}",
                        "X-PrettyPrint": "1",
                    },
                )
        except RequestException:
            raise SSOError("Could not establish connection")

        if response.status_code != requests.codes.ok:
            raise SSOError(f"Authentication error: {response.text}")

        return self.extract_email(response)

    def _retrieve_access_token(self, authorization_code: str) -> str:
        """
        Retrieve the access token.

        This method performs an external API call.
        """
        try:
            with session:
                response = session.post(
                    url=self.token_url,
                    data={
                        "grant_type": "authorization_code",
                        "code": authorization_code,
                        "client_id": self.client_id,
                        "client_secret": self.client_secret,
                        "redirect_uri": self.callback_url,
                    },
                    headers={
                        "Content-type": "application/x-www-form-urlencoded",
                        "Accept": "application/json",
                    },
                )
        except RequestException:
            raise SSOError("Could not establish connection")

        if response.status_code != requests.codes.ok:
            raise SSOError(f"Authentication error: {response.text}")

        response_data = response.json()
        access_token = response_data.get("access_token")
        if not isinstance(access_token, str):
            raise SSOError(f"Invalid access token request response: {response_data}")
        return access_token


class GoogleSSO(OAuthSSO):
    identity_provider = IdentityProvider.google

    email_attribute = "email"

    scopes = ["profile", "email"]

    authorize_url = "https://accounts.google.com/o/oauth2/auth"
    token_url = "https://accounts.google.com/o/oauth2/token"  # noqa: S105
    email_api_url = "https://www.googleapis.com/oauth2/v1/userinfo"

    @property
    def client_id(self) -> str:
        return settings.GOOGLE_SSO_CLIENT_ID

    @property
    def client_secret(self) -> str:
        return settings.GOOGLE_SSO_CLIENT_SECRET


class MicrosoftSSO(OAuthSSO):
    identity_provider = IdentityProvider.microsoft

    email_attribute = "userPrincipalName"

    scopes = ["User.Read", "openid"]

    authorize_url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"  # noqa: S105
    email_api_url = "https://graph.microsoft.com/v1.0/me"

    @property
    def client_id(self) -> str:
        return settings.MICROSOFT_SSO_CLIENT_ID

    @property
    def client_secret(self) -> str:
        return settings.MICROSOFT_SSO_CLIENT_SECRET


class GitHubSSO(OAuthSSO):
    identity_provider = IdentityProvider.github

    scopes = ["user:email"]

    authorize_url = "https://github.com/login/oauth/authorize"
    token_url = "https://github.com/login/oauth/access_token"  # noqa: S105
    email_api_url = "https://api.github.com/user/emails"

    @property
    def client_id(self) -> str:
        return settings.GITHUB_SSO_CLIENT_ID

    @property
    def client_secret(self) -> str:
        return settings.GITHUB_SSO_CLIENT_SECRET

    def extract_email(self, response: requests.Response) -> str:
        for email in response.json():
            if email["primary"]:
                return email["email"]
        raise SSOError("Email address not found in the response")
