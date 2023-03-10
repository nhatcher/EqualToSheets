# Run in manage.py

from serverless.models import License, Workbook

license = License.objects.create(email="test@equalto.com", email_verified=True, key="00000000-0000-0000-0000-000000000000")
Workbook.objects.create(license=license, id="00000000-0000-0000-0000-000000000000")

