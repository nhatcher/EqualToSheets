from django.urls import path

from serverless.rest_api.v1 import (
    CellByIndexView,
    SheetDetailView,
    SheetListView,
    WorkbookDetailView,
    WorkbookListView,
    WorkbookXLSXView,
)

urlpatterns = [
    path("v1/workbooks", WorkbookListView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>", WorkbookDetailView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/xlsx", WorkbookXLSXView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/sheets", SheetListView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/sheets/<int:sheet_id>", SheetDetailView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/sheets/<int:sheet_id>/cells/<int:row>/<int:col>", CellByIndexView.as_view()),
]
