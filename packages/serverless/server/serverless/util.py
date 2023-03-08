from serverless.log import info
from serverless.models import License, LicenseDomain


def is_license_key_valid_for_host(license_key: str, host: str) -> bool:
    # the license is valid if:
    #   1. LicenseDomain contains a record for domain, OR
    #   2. LicenseDomain contains a record for the parent of domain, with *. at the start
    if host.startswith("http://"):
        host = host[len("http://") :]
    elif host.startswith("https://"):
        host = host[len("https://") :]
    domain = host.split(":")[0]
    parent = ".".join(domain.split(".")[1:])
    parent_wild_card = "*.%s" % parent
    info("license_key=%s, host=%s" % (license_key, host))
    license = License.objects.get(key=license_key)
    info("got license for %s, host=%s, email_verified=%s" % (license.email, host, license.email_verified))
    return license.email_verified and (
        # WARNING: during the beta, if a license has 0 domains then it works on all domains
        # TODO: remove this after the beta
        LicenseDomain.objects.filter(license=license).count() == 0
        or (
            # exact mach
            LicenseDomain.objects.filter(license=license, domain=domain).exists()
            or
            # parent with wild card
            LicenseDomain.objects.filter(license=license, domain=parent_wild_card).exists()
        )
    )
