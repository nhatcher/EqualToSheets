from urllib.parse import quote

from django.conf import settings
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Email, Mail

from serverless.models import License

sendgrid_client = SendGridAPIClient(settings.SENDGRID_API_KEY)

EMAIL_TO_SUBSCRIBERS_TEMPLATE = "d-0b7571668db1404796cbf93a2774a25b"


def _send_email(message: Mail) -> None:  # pragma: no cover
    sendgrid_client.send(message)


def send_license_email_to_subscriber(email: str) -> None:
    license = License.objects.filter(email=email).first()
    if not license:
        license = License(email=email)
    # we auto-verify the license for subscribers
    license.save()
    message = Mail(
        from_email=Email(name="Diarmuid Glynn", email="diarmuid.glynn@equalto.com"),
        to_emails=email,
    )
    message.template_id = EMAIL_TO_SUBSCRIBERS_TEMPLATE
    message.dynamic_template_data = {
        "licenseKeyActivatedURL": f"{settings.SERVER}/#/license/activate/{license.id}",
        "unsubscribeURL": f"{settings.SERVER}/unsubscribe-email?email={quote(email)}",
    }
    _send_email(message)
