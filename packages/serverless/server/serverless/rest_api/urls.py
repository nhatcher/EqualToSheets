from django.urls import path, register_converter

from serverless.rest_api.converters import CreateWorkbookTypeConverter
from serverless.rest_api.v1 import (
    CellByIndexView,
    SheetDetailView,
    SheetListView,
    WorkbookDetailView,
    WorkbookListView,
    WorkbookXLSXView,
)

register_converter(CreateWorkbookTypeConverter, "create_workbook_type_converter")

urlpatterns = [
    path("v1/workbooks", WorkbookListView.as_view()),
    path(
        "v1/workbooks/<create_workbook_type_converter:create_type>",
        WorkbookListView.as_view(http_method_names=["post"]),
    ),
    path("v1/workbooks/<uuid:workbook_id>", WorkbookDetailView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/xlsx", WorkbookXLSXView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/sheets", SheetListView.as_view()),
    path("v1/workbooks/<uuid:workbook_id>/sheets/<int:sheet_index>", SheetDetailView.as_view()),
    path(
        "v1/workbooks/<uuid:workbook_id>/sheets/<int:sheet_index>/cells/<int:row>/<int:col>",
        CellByIndexView.as_view(),
    ),
]
