import uuid

import equalto
from django.db import models
from django.utils import timezone


def get_default_workbook():
    # TODO: use a better default workbook
    return open("serverless/data/XLOOKUP.xlsx.json", encoding="utf-8").read()




# 1. User visits www.equalto.com/serverless
# 2. Enters email address to sign-up
# 3. Receives an email with the license key

# To access:
# 
class License(models.Model):
    id = models.UUIDField(
          primary_key = True,
          default = uuid.uuid4,
          editable = False,
          null=False)
    
    key = models.UUIDField(
          default = uuid.uuid4,
          null=False)
    
    email = models.CharField(
        max_length=256,
        null=False,
        unique=True)
    
    email_verified = models.BooleanField(
        default=False,
        null=False)
    
    create_datetime = models.DateTimeField(default=timezone.now, null=False)

class LicenseDomain(models.Model):
    id = models.UUIDField(
          primary_key = True,
          default = uuid.uuid4,
          editable = False,
          null=False)
    
    license = models.ForeignKey(License, on_delete=models.CASCADE)
    
    # Either specific domain ("example.com") or wild-card subdomain ("*.example.com")
    domain = models.CharField(
        max_length=256,
        null=False)
    
    # Add unique constraint for (license, domain)



class Workbook(models.Model):
    id = models.UUIDField(
         primary_key = True,
         default = uuid.uuid4,
         editable = False)
    license = models.ForeignKey(License, on_delete=models.CASCADE)
    workbook_json = models.JSONField(null=False, default=get_default_workbook)
    revision = models.IntegerField(default=1, null=False)
    create_datetime = models.DateTimeField(default=timezone.now, null=False)
    modify_datetime = models.DateTimeField(default=timezone.now, null=False)




