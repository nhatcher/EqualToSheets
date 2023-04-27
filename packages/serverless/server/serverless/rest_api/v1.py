import enum
import os.path
import tempfile
from typing import Any, Callable

import equalto.cell
import equalto.exceptions
import equalto.sheet
import equalto.workbook
from django.db import transaction
from django.db.models import QuerySet
from django.http import HttpResponse
from drf_spectacular.types import OpenApiTypes
from drf_spectacular.utils import extend_schema
from rest_framework import serializers, status
from rest_framework.exceptions import AuthenticationFailed, NotAuthenticated, NotFound
from rest_framework.request import Request
from rest_framework.response import Response
from rest_framework.views import APIView, exception_handler

from serverless import log
from serverless.models import License, Workbook
from serverless.util import LicenseKeyError, get_license, is_license_key_valid_for_host
from serverless.views import MAX_XLSX_FILE_SIZE


class BadRequestError(Exception):
    """Exception translated to 400 Bad Request response."""


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


def custom_exception_handler(exception: Exception, context: Any) -> Response:
    """Exception handler replacing uncaught WorkbookErrors and BadRequestErrors with 400 Bad Request."""
    if isinstance(exception, (BadRequestError, equalto.exceptions.WorkbookError)):
        return Response({"detail": str(exception)}, status=status.HTTP_400_BAD_REQUEST)
    return exception_handler(exception, context)


class ServerlessView(APIView):
    def get_exception_handler(self) -> Callable[[Exception, Any], Response]:
        return custom_exception_handler

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

    def _get_sheet(self, workbook: Workbook, sheet_index: int) -> equalto.sheet.Sheet:
        if sheet_index <= 0:
            raise NotFound("Sheet index must be >= 1")
        try:
            return workbook.calc.sheets[sheet_index - 1]
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


class CreateWorkbookType(enum.Enum):
    blank = "blank"
    json = "json"
    xlsx = "xlsx"

    def __str__(self) -> str:
        return self.value


class WorkbookListView(ServerlessView):
    @extend_schema(responses={(201, "application/json"): WorkbookSerializer(many=True)})
    def get(self, request: Request) -> Response:
        """Get a list of all workbooks."""
        serializer = WorkbookSerializer(self._get_workbooks().order_by("create_datetime"), many=True)
        return Response({"workbooks": serializer.data})

    @transaction.atomic
    def post(self, request: Request, create_type: CreateWorkbookType | None = None) -> Response:
        """Create a new workbook.

        POST parameters supported by this endpoint:

        * `name`: (optional) name of newly created workbook
        * `version`: (optional) version of schema used in the associated `workbook_json`
        * `workbook_json`: (optional) the JSON used to create the new workbook

        There are multiple ways to use this endpoint:

        1. Create a blank workbook, you may specify the `name` but nothing else.
        2. Create a workbook from JSON, you must specify `version` and `workbook_json`, you may specify the `name` and
           nothing else.
        3. Create a workbook from XLSX, you must specify `xlsx_file`, you may specify the `name` and nothing else.

        You can explicitly specify the input data type in the URL in order to get better error reporting:

        * /api/v1/workbooks/blank
        * /api/v1/workbooks/json
        * /api/v1/workbooks/xlsx
        """
        create_type = create_type or self._get_create_type_from_parameters()

        workbook = self._create_workbook(create_type)

        serializer = WorkbookSerializer(workbook)
        return Response(serializer.data, status=status.HTTP_201_CREATED)

    def _create_workbook(self, create_type: CreateWorkbookType) -> Workbook:
        return {
            CreateWorkbookType.blank: self._create_blank_workbook,
            CreateWorkbookType.json: self._create_workbook_from_json,
            CreateWorkbookType.xlsx: self._create_workbook_from_xlsx,
        }[create_type]()

    def _create_blank_workbook(self) -> Workbook:
        return Workbook.objects.create(license=self._get_license(), name=self.request.data.get("name", "Book"))

    def _create_workbook_from_json(self) -> Workbook:
        if "version" not in self.request.data:
            raise BadRequestError("'version' parameter is not provided")
        if "workbook_json" not in self.request.data:
            raise BadRequestError("'workbook_json' parameter is not provided")

        version = self.request.data["version"]
        workbook_json = self.request.data["workbook_json"]

        # creating a workbook from json
        if version != "1":
            # TODO: in future, when we have new JSON schemas, we'll need to auto-migrate old JSON to the new
            #       structure
            raise BadRequestError("Currently, we only support JSON using the version 1 schema.")

        # evaluate the workbook JSON, to ensure it's valid and fully computed
        with equalto.exceptions.SuppressEvaluationErrors():
            equalto_workbook = equalto.loads(workbook_json)

        return Workbook.objects.create(
            license=self._get_license(),
            workbook_json=equalto_workbook.json,
            name=self.request.data.get("name", "Book"),
        )

    def _create_workbook_from_xlsx(self) -> Workbook:
        if "xlsx_file" not in self.request.FILES:
            raise BadRequestError("'xlsx_file' is not provided")
        xlsx_file = self.request.FILES["xlsx_file"]

        if xlsx_file.size > MAX_XLSX_FILE_SIZE:
            raise BadRequestError(f"Excel file too large (max size {MAX_XLSX_FILE_SIZE} bytes).")
        tmp = tempfile.NamedTemporaryFile()
        tmp.write(xlsx_file.read())

        with equalto.exceptions.SuppressEvaluationErrors():
            try:
                equalto_workbook = equalto.load(tmp.name)
            except equalto.exceptions.WorkbookError:
                raise BadRequestError("Could not upload workbook.")

        return Workbook.objects.create(
            license=self._get_license(),
            workbook_json=equalto_workbook.json,
            name=self.request.data.get("name", "Book"),
        )

    def _get_create_type_from_parameters(self) -> CreateWorkbookType:
        # TODO: We shouldn't allow to create new workbooks without explicitly specifying `create_type` in the URL
        data = {
            "version": self.request.data.get("version"),
            "workbook_json": self.request.data.get("workbook_json"),
            "name": self.request.data.get("name"),
            "xlsx_file": self.request.FILES.get("xlsx_file"),
        }

        # define the "function signatures" for the different ways to create a workbook, so that
        # we can implement "multiple dispatch" function overloading
        signatures: dict[CreateWorkbookType, dict[str, list[str]]] = {  # noqa: WPS234
            CreateWorkbookType.json: {"required": ["version", "workbook_json"], "optional": ["name"]},
            CreateWorkbookType.xlsx: {"required": ["xlsx_file"], "optional": ["name"]},
            CreateWorkbookType.blank: {"required": [], "optional": ["name"]},
        }

        # based upon the function signatures above, determine how exactly the user wants to create the workbook
        # and record it in the create_type variable.
        create_type = None
        all_params = data.keys()
        for c, obj in signatures.items():
            required = obj["required"]
            optional = obj["optional"]
            forbidden = set(all_params).difference(required + optional)
            if all(data[r] is not None for r in required):
                create_type = c
                if any(data[r] is not None for r in forbidden):
                    should_not_be_present = [f for f in forbidden if data[f]]
                    should_not_be_present.sort()
                    raise BadRequestError(
                        f"When creating a workbook from '{create_type}' data, the following parameters "
                        + f"are forbidden: {should_not_be_present}.",
                    )
                break
        assert create_type is not None
        return create_type


class WorkbookDetailView(ServerlessView):
    @extend_schema(responses={(200, "application/json"): WorkbookDetailSerializer(many=True)})
    def get(self, request: Request, workbook_id: Workbook) -> Response:
        """Get the workbook data."""
        serializer = WorkbookDetailSerializer(self._get_workbook(workbook_id))
        return Response(serializer.data)


class WorkbookXLSXView(ServerlessView):
    @extend_schema(
        responses={
            (200, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"): OpenApiTypes.BINARY,
        },
    )
    def get(self, request: Request, workbook_id: Workbook) -> HttpResponse:
        """Export workbook as a Microsoft XLSX file."""
        workbook = self._get_workbook(workbook_id)
        with equalto.exceptions.SuppressEvaluationErrors() as _:  # noqa: WPS122
            equalto_workbook = equalto.loads(workbook.workbook_json)
        tmp_name = os.path.join(tempfile.mkdtemp(), "temp.xlsx")
        equalto_workbook.save(tmp_name)
        with open(tmp_name, "rb") as xlsx_file:
            data = xlsx_file.read()
            # TODO: set the filename to the (sanitized) name of the workbook
            return HttpResponse(
                data,
                headers={"Content-Disposition": 'attachment; filename="download.xlsx"'},
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


class SheetListView(ServerlessView):
    @extend_schema(responses={(200, "application/json"): serialize_sheet})
    def get(self, request: Request, workbook_id: str) -> Response:
        """Get the sheets in a workbook."""
        workbook = self._get_workbook(workbook_id)
        sheets = workbook.calc.sheets
        return Response({"sheets": [serialize_sheet(sheet) for sheet in sheets]})

    @transaction.atomic
    @extend_schema(responses={(201, "application/json"): serialize_sheet})
    def post(self, request: Request, workbook_id: str) -> Response:
        """Create a new blank sheet in a workbook."""
        name = request.data.get("name")

        workbook = self._get_workbook(workbook_id)

        new_sheet = workbook.calc.sheets.add(name)

        workbook.set_workbook_json(workbook.calc.json)

        return Response(serialize_sheet(new_sheet), status=status.HTTP_201_CREATED)


class SheetDetailView(ServerlessView):
    @extend_schema(responses={(200, "application/json"): serialize_sheet})
    def get(self, request: Request, workbook_id: str, sheet_index: int) -> Response:
        """Get the metadata of a sheet in a workbook."""
        sheet = self._get_sheet(self._get_workbook(workbook_id), sheet_index)
        return Response(serialize_sheet(sheet))

    @extend_schema(responses={(200, "application/json"): serialize_sheet})
    @transaction.atomic
    def put(self, request: Request, workbook_id: str, sheet_index: int) -> Response:
        """Rename a sheet in a workbook."""
        new_name = request.data.get("new_name")
        if not new_name:
            raise BadRequestError("'new_name' parameter is not provided")

        workbook = self._get_workbook(workbook_id)

        sheet = self._get_sheet(workbook, sheet_index)
        sheet.name = new_name

        workbook.set_workbook_json(workbook.calc.json)

        return Response(serialize_sheet(sheet))

    @transaction.atomic
    def delete(self, request: Request, workbook_id: str, sheet_index: int) -> Response:
        """Delete a sheet in a workbook."""
        workbook = self._get_workbook(workbook_id)

        self._get_sheet(workbook, sheet_index).delete()

        workbook.set_workbook_json(workbook.calc.json)

        return Response(status=status.HTTP_204_NO_CONTENT)


class CellByIndexView(ServerlessView):
    @extend_schema(responses={(200, "application/json"): serialize_cell})
    def get(self, request: Request, workbook_id: str, sheet_index: int, row: int, col: int) -> Response:
        """Get all the cell data."""
        sheet = self._get_sheet(self._get_workbook(workbook_id), sheet_index)
        try:
            cell = sheet.cell(row, col)
        except equalto.exceptions.WorkbookError:
            raise NotFound("Cell not found")
        return Response(serialize_cell(cell))

    @extend_schema(responses={(200, "application/json"): serialize_cell})
    @transaction.atomic
    def put(self, request: Request, workbook_id: str, sheet_index: int, row: int, col: int) -> Response:
        """Update the cell data."""
        cell_input = request.data.get("input")
        cell_value = request.data.get("value")
        if (cell_input is None) == (cell_value is None):
            raise BadRequestError("Either 'input' or 'value' parameter needs to be provided, but not both")

        workbook = self._get_workbook(workbook_id)
        sheet = self._get_sheet(workbook, sheet_index)

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
