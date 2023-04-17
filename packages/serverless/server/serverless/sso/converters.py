from serverless.sso.oauth import GitHubSSO, GoogleSSO, IdentityProvider, MicrosoftSSO, OAuthSSO


class SSOIntegrationConverter:
    regex = ".*"

    idp_to_sso_integration = {
        IdentityProvider.github: GitHubSSO(),
        IdentityProvider.google: GoogleSSO(),
        IdentityProvider.microsoft: MicrosoftSSO(),
    }

    def to_python(self, value: str) -> OAuthSSO:
        return self.idp_to_sso_integration[IdentityProvider(value)]

    def to_url(self, sso: OAuthSSO) -> str:
        return sso.identity_provider.value
