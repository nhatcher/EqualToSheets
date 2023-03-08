from __future__ import annotations

import re

from django.conf import settings
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Email, Mail

from serverless import log
from serverless.models import License

sendgrid_client = SendGridAPIClient(settings.SENDGRID_API_KEY)

LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID = "d-e535aeb95e8e4c598574f32e50450cbf"


def send_license_activation_email(license: License) -> None:
    assert LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID, "LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID is not set"

    message = Mail(
        from_email=Email(name="EqualTo", email="no-reply@equalto.com"),
        to_emails=license.email,
    )
    message.template_id = LICENSE_ACTIVATION_EMAIL_TEMPLATE_ID
    message.dynamic_template_data = {"emailVerificationURL": f"{settings.SERVER}/#/license/activate/{license.id}"}

    should_send_email = not settings.ALLOW_ONLY_EMPLOYEE_EMAILS or re.match(
        settings.EMPLOYEE_EMAIL_PATTERN,
        license.email,
    )

    if should_send_email:
        _send_email(message)
        log.info(f"Sent license activation email to {license.email}.")
    elif not settings.TEST:
        log.error(
            "The following dynamic template email would be sent but the config allows to send emails only "
            + f"to EqualTo employees: {message.get()}.",
        )


def _send_email(message: Mail) -> None:  # pragma: no cover
    if settings.TEST:
        return  # emails shouldn't be sent in unit tests

    assert settings.SENDGRID_API_KEY, "SENDGRID_API_KEY is not set"

    sendgrid_client.send(message)
