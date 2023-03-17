import uuid
from functools import cache
from typing import Any

import equalto.workbook
from django.db import models
from django.utils import timezone
from django.utils.functional import cached_property
from equalto.exceptions import SuppressEvaluationErrors


@cache
def get_default_workbook() -> str:
    # TODO: use a better default workbook
    with open("serverless/data/XLOOKUP.xlsx.json", encoding="utf-8") as f:
        return f.read()


# 1. User visits www.equalto.com/serverless
# 2. Enters email address to sign-up
# 3. Receives an email with the license key


# To access:
class License(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False, null=False)

    key = models.UUIDField(default=uuid.uuid4, null=False)

    email = models.CharField(max_length=256, null=False, unique=True)

    email_verified = models.BooleanField(default=False, null=False)

    create_datetime = models.DateTimeField(default=timezone.now, null=False)


class LicenseDomain(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False, null=False)

    license = models.ForeignKey(License, on_delete=models.CASCADE)

    # Either specific domain ("example.com") or wild-card subdomain ("*.example.com")
    domain = models.CharField(max_length=256, null=False)

    # Add unique constraint for (license, domain)


class Workbook(models.Model):
    id = models.UUIDField(primary_key=True, default=uuid.uuid4, editable=False)
    name = models.CharField(max_length=256, null=False, default="")
    license = models.ForeignKey(License, on_delete=models.CASCADE)
    workbook_json = models.JSONField(null=False, default=get_default_workbook)
    revision = models.IntegerField(default=1, null=False)
    create_datetime = models.DateTimeField(default=timezone.now, null=False)
    modify_datetime = models.DateTimeField(default=timezone.now, null=False)

    # workbook json version
    version = models.IntegerField(default=1, null=False)

    def set_workbook_json(self, workbook_json: dict[str, Any]) -> None:
        if self.workbook_json == workbook_json:
            return

        self.workbook_json = workbook_json
        self.revision += 1

        self.save(update_fields=["workbook_json", "revision"])

        # invalidate `calc` property cache
        self.__dict__.pop("calc", None)

    @cached_property
    def calc(self) -> equalto.workbook.Workbook:
        with SuppressEvaluationErrors():
            return equalto.loads(self.workbook_json)


class UnsubscribedEmail(models.Model):
    # WARNING!
    # This is a hack - but I need a quick way to track users who want to unsubscribe
    # from our mailings. We should move to proper system post-beta (HubSpot?).
    # Users can be appended to the UnsubscribedEmail list by clicking on a link
    # to /unsubscribe?email=...
    email = models.CharField(max_length=256, primary_key=True, null=False)
    create_datetime = models.DateTimeField(default=timezone.now, null=False)
