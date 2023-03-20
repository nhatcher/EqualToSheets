from typing import Any, Callable

import equalto.cell
import equalto.exceptions
import equalto.sheet
import equalto.workbook
from django.db import transaction
from django.db.models import QuerySet
from rest_framework import serializers, status
from rest_framework.exceptions import AuthenticationFailed, NotAuthenticated, NotFound
from rest_framework.request import Request
from rest_framework.response import Response
from rest_framework.views import APIView, exception_handler

from serverless import log
from serverless.models import License, Workbook
from serverless.util import LicenseKeyError, get_license, is_license_key_valid_for_host


class ServerlessView(APIView):
    def get_exception_handler(self) -> Callable[[Exception, Any], Response]:
        """Exception handler replacing uncaught WorkbookErrors with 400 Bad Request."""

        def workbook_exception_handler(exception: Exception, context: Any) -> Response:
            if isinstance(exception, equalto.exceptions.WorkbookError):
                return Response({"detail": str(exception)}, status=status.HTTP_400_BAD_REQUEST)
            return exception_handler(exception, context)

        return workbook_exception_handler

    def _get_license(self) -> License:
        try:
            license = get_license(self.request.META)
        except LicenseKeyError:
            raise NotAuthenticated()

        origin = self.request.META.get("HTTP_ORIGIN")
        if not is_license_key_valid_for_host(license.key, origin):
            log.error(f"License key {license.key} is not valid for {origin}")
            raise AuthenticationFailed("License key is not valid")

        return license

    def _get_workbooks(self) -> QuerySet[Workbook]:
        qs = Workbook.objects.filter(license=self._get_license())
        if self.request.method in {"POST", "PUT", "DELETE"}:
            qs = qs.select_for_update()
        return qs

    def _get_workbook(self, workbook_id: str) -> Workbook:
        try:
            return self._get_workbooks().get(id=workbook_id)
        except Workbook.DoesNotExist:
            raise NotFound("Workbook not found")

    def _get_sheet(self, workbook: Workbook, sheet_id: int) -> equalto.sheet.Sheet:
        try:
            return workbook.calc.sheets.get_sheet_by_id(sheet_id)
        except equalto.exceptions.WorkbookError:
            raise NotFound("Sheet not found")


class WorkbookSerializer(serializers.Serializer):
    id = serializers.UUIDField()
    name = serializers.CharField()
    revision = serializers.IntegerField()
    create_datetime = serializers.DateTimeField()
    modify_datetime = serializers.DateTimeField()


class WorkbookDetailSerializer(WorkbookSerializer):
    workbook_json = serializers.JSONField()


class WorkbookListView(ServerlessView):
    def get(self, request: Request) -> Response:
        """Get a list of all workbooks."""
        serializer = WorkbookSerializer(self._get_workbooks().order_by("create_datetime"), many=True)
        return Response({"workbooks": serializer.data})

    @transaction.atomic
    def post(self, request: Request) -> Response:
        """Create a new blank workbook."""
        workbook = Workbook.objects.create(license=self._get_license())
        serializer = WorkbookSerializer(workbook)
        return Response(serializer.data, status=status.HTTP_201_CREATED)


class WorkbookDetailView(ServerlessView):
    def get(self, request: Request, workbook_id: Workbook) -> Response:
        """Get the workbook data."""
        serializer = WorkbookDetailSerializer(self._get_workbook(workbook_id))
        return Response(serializer.data)


class SheetListView(ServerlessView):
    def get(self, request: Request, workbook_id: str) -> Response:
        """Get the sheets in a workbook."""
        workbook = self._get_workbook(workbook_id)
        sheets = workbook.calc.sheets
        return Response({"sheets": [serialize_sheet(sheet) for sheet in sheets]})

    @transaction.atomic
    def post(self, request: Request, workbook_id: str) -> Response:
        """Create a new blank sheet in a workbook."""
        name = request.data.get("name")

        workbook = self._get_workbook(workbook_id)

        new_sheet = workbook.calc.sheets.add(name)

        workbook.set_workbook_json(workbook.calc.json)

        return Response(serialize_sheet(new_sheet), status=status.HTTP_201_CREATED)


class SheetDetailView(ServerlessView):
    def get(self, request: Request, workbook_id: str, sheet_id: int) -> Response:
        """Get the metadata of a sheet in a workbook."""
        sheet = self._get_sheet(self._get_workbook(workbook_id), sheet_id)
        return Response(serialize_sheet(sheet))

    @transaction.atomic
    def put(self, request: Request, workbook_id: str, sheet_id: int) -> Response:
        """Rename a sheet in a workbook."""
        new_name = request.data.get("new_name")
        if not new_name:
            return Response(
                {"detail": "'new_name' parameter is not provided"},
                status=status.HTTP_400_BAD_REQUEST,
            )

        workbook = self._get_workbook(workbook_id)

        sheet = self._get_sheet(workbook, sheet_id)
        sheet.name = new_name

        workbook.set_workbook_json(workbook.calc.json)

        return Response(serialize_sheet(sheet))

    @transaction.atomic
    def delete(self, request: Request, workbook_id: str, sheet_id: int) -> Response:
        """Delete a sheet in a workbook."""
        workbook = self._get_workbook(workbook_id)

        self._get_sheet(workbook, sheet_id).delete()

        workbook.set_workbook_json(workbook.calc.json)

        return Response(status=status.HTTP_204_NO_CONTENT)


class CellByIndexView(ServerlessView):
    def get(self, request: Request, workbook_id: str, sheet_id: int, row: int, col: int) -> Response:
        """Get all the cell data."""
        sheet = self._get_sheet(self._get_workbook(workbook_id), sheet_id)
        try:
            cell = sheet.cell(row, col)
        except equalto.exceptions.WorkbookError:
            raise NotFound("Cell not found")
        return Response(serialize_cell(cell))

    @transaction.atomic
    def put(self, request: Request, workbook_id: str, sheet_id: int, row: int, col: int) -> Response:
        """Update the cell data."""
        cell_input = request.data.get("input")
        cell_value = request.data.get("value")
        if (cell_input is None) == (cell_value is None):
            return Response(
                {"detail": "Either 'input' or 'value' parameter needs to be provided, but not both"},
                status=status.HTTP_400_BAD_REQUEST,
            )

        workbook = self._get_workbook(workbook_id)
        sheet = self._get_sheet(workbook, sheet_id)

        try:
            cell = sheet.cell(row, col)
        except equalto.exceptions.WorkbookError:
            raise NotFound("Cell not found")

        with equalto.exceptions.SuppressEvaluationErrors():
            # TODO: We should probably return the suppressed errors to the user post beta
            if cell_input is not None:
                cell.set_user_input(cell_input)
            else:
                assert cell_value is not None
                cell.value = cell_value

        workbook.set_workbook_json(workbook.calc.json)

        return Response(serialize_cell(cell))


def serialize_sheet(sheet: equalto.sheet.Sheet) -> dict[str, Any]:
    return {"id": sheet.sheet_id, "name": sheet.name, "index": sheet.index}


def serialize_cell(cell: equalto.cell.Cell) -> dict[str, Any]:
    return {
        "formatted_value": str(cell),
        "value": cell.value,
        "format": cell.style.format,
        "type": cell.type.name,
        "formula": cell.formula,
    }
