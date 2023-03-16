from typing import Any, Callable, Iterable, Self

import equalto.cell
import equalto.exceptions
import equalto.sheet
import equalto.workbook
import graphene
from graphene_django import DjangoObjectType
from graphql import GraphQLError

from serverless import log, models
from serverless.util import get_license, is_license_key_valid_for_host

MAX_WORKBOOKS_PER_LICENSE = 1000
# 10Mb limit for the beta
MAX_WORKBOOK_JSON_SIZE = 10 * 1024 * 1024
# The max length of the input value assigned to a cell
MAX_WORKBOOK_INPUT_SIZE = 512


def _validate_license_for_workbook(info: graphene.ResolveInfo, workbook_id: str) -> None:
    license = get_license(info.context.META)
    workbook = models.Workbook.objects.get(id=workbook_id)
    if workbook.license != license:
        log.error("User %s cannot access workbook %s" % (license.email, workbook))
        raise GraphQLError("ERROR - user can't access workbook")
    origin = info.context.META.get("HTTP_ORIGIN")
    if not is_license_key_valid_for_host(license.key, origin):
        log.error("User %s cannot access workbook %s from origin %s" % (license.email, workbook, origin))
        raise GraphQLError("ERROR - user can't access workbook")


def validate_license_for_workbook_mutation(graphql_fn: Callable[..., Any]) -> Callable[..., Any]:
    def fn(
        cls: type[graphene.Mutation],
        root: Any,
        info: graphene.ResolveInfo,
        workbook_id: str,
        *args: Any,
        **kwargs: Any,
    ) -> Callable[..., Any]:
        _validate_license_for_workbook(info, workbook_id)
        return graphql_fn(cls, root, info, workbook_id, *args, **kwargs)

    return fn


def validate_license_for_workbook_query(graphql_fn: Callable[..., Any]) -> Callable[..., Any]:
    def fn(root: Any, info: graphene.ResolveInfo, workbook_id: str, *args: Any, **kwargs: Any) -> Callable[..., Any]:
        _validate_license_for_workbook(info, workbook_id)
        return graphql_fn(root, info, workbook_id, *args, **kwargs)

    return fn


CellType = graphene.Enum.from_enum(equalto.cell.CellType)


class CellValue(graphene.ObjectType):
    text = graphene.String()
    number = graphene.Float()
    boolean = graphene.Boolean()


class Cell(graphene.ObjectType):
    def __init__(self, calc_cell: equalto.cell.Cell) -> None:
        self._calc_cell = calc_cell

    formatted_value = graphene.String(required=True)
    value = graphene.Field(CellValue, required=True)
    format = graphene.String(required=True)
    type = graphene.Field(CellType, required=True)
    formula = graphene.String()

    def resolve_formatted_value(self, info: graphene.ResolveInfo) -> str:
        return str(self._calc_cell)

    def resolve_value(self, info: graphene.ResolveInfo) -> CellValue:
        value = self._calc_cell.value
        if isinstance(value, bool):
            return CellValue(boolean=value)
        elif isinstance(value, (int, float)):
            return CellValue(number=float(value))
        elif isinstance(value, str):
            return CellValue(text=value)
        else:
            raise Exception(f"Unrecognized {type(value)=}")

    def resolve_format(self, info: graphene.ResolveInfo) -> str:
        return self._calc_cell.style.format

    def resolve_type(self, info: graphene.ResolveInfo) -> equalto.cell.CellType:
        return self._calc_cell.type

    def resolve_formula(self, info: graphene.ResolveInfo) -> str | None:
        return self._calc_cell.formula


class Sheet(graphene.ObjectType):
    def __init__(self, calc_sheet: equalto.sheet.Sheet) -> None:
        self._calc_sheet = calc_sheet

    name = graphene.String()
    id = graphene.Int()

    def resolve_name(self, info: graphene.ResolveInfo) -> str:
        return self._calc_sheet.name

    def resolve_id(self, info: graphene.ResolveInfo) -> int:
        return self._calc_sheet.sheet_id

    cell = graphene.Field(Cell, required=True, ref=graphene.String(), row=graphene.Int(), col=graphene.Int())

    def resolve_cell(
        self,
        info: graphene.ResolveInfo,
        ref: str | None = None,
        row: int | None = None,
        col: int | None = None,
    ) -> Cell:
        if ref is not None and row is None and col is None:
            return Cell(calc_cell=self._calc_sheet[ref])
        elif ref is None and row is not None and col is not None:
            return Cell(calc_cell=self._calc_sheet.cell(row, col))
        else:
            log.error("ERROR - ref/row/col")
            raise GraphQLError("ERROR - ref/row/col")


class Workbook(DjangoObjectType):
    class Meta:
        model = models.Workbook
        fields = (
            "id",
            "workbook_json",
            "create_datetime",
            "modify_datetime",
            "revision",
            "name",
        )

    sheet = graphene.Field(Sheet, sheet_id=graphene.Int(), name=graphene.String())
    sheets = graphene.List(Sheet)

    def resolve_sheet(self, info: graphene.ResolveInfo, sheet_id: int | None = None, name: str | None = None) -> Sheet:
        if sheet_id is not None and name is None:
            return Sheet(calc_sheet=self.calc.sheets.get_sheet_by_id(sheet_id))
        elif sheet_id is None and name is not None:
            return Sheet(calc_sheet=self.calc.sheets[name])
        else:
            log.error("ERROR - name/id")
            raise Exception("ERROR - name/id")

    def resolve_sheets(self, info: graphene.ResolveInfo) -> Iterable[Sheet]:
        log.info("Origin: %s" % info.context.META.get("Origin"))
        yield from (Sheet(calc_sheet=sheet) for sheet in self.calc.sheets)


class Query(graphene.ObjectType):
    workbook = graphene.Field(Workbook, workbook_id=graphene.String(required=True))
    workbooks = graphene.List(Workbook)

    @validate_license_for_workbook_query
    def resolve_workbook(self, info: graphene.ResolveInfo, workbook_id: str) -> models.Workbook:
        return models.Workbook.objects.get(id=workbook_id)

    def resolve_workbooks(self, info: graphene.ResolveInfo) -> Iterable[models.Workbook]:
        license = get_license(info.context.META)
        return models.Workbook.objects.filter(license=license)


class SaveWorkbook(graphene.Mutation):
    class Arguments:
        workbook_id = graphene.String()
        workbook_json = graphene.String()

    revision = graphene.Int()

    @classmethod
    @validate_license_for_workbook_mutation
    def mutate(
        cls,
        root: Any,
        info: graphene.ResolveInfo,
        workbook_id: str,
        workbook_json: str,
    ) -> Self:
        if len(workbook_json) > MAX_WORKBOOK_JSON_SIZE:
            raise GraphQLError("Workbook JSON too large")

        workbook = models.Workbook.objects.select_for_update().get(id=workbook_id)

        try:
            workbook_json = equalto.loads(workbook_json).json
        except equalto.exceptions.WorkbookError:
            raise GraphQLError("Could not parse workbook JSON")

        workbook.set_workbook_json(workbook_json)

        return SaveWorkbook(revision=workbook.revision)


class SetCellInput(graphene.Mutation):
    # simulates entering text in the spreadsheet widget
    class Arguments:
        workbook_id = graphene.String(required=True)
        sheet_id = graphene.Int()
        sheet_name = graphene.String()
        ref = graphene.String()
        row = graphene.Int()
        col = graphene.Int()

        input = graphene.String(required=True)

    workbook = graphene.Field(Workbook, required=True)

    @classmethod
    @validate_license_for_workbook_mutation
    def mutate(
        cls,
        root: Any,
        info: graphene.ResolveInfo,
        workbook_id: str,
        input: str,
        sheet_id: int | None = None,
        sheet_name: str | None = None,
        ref: str | None = None,
        row: int | None = None,
        col: int | None = None,
    ) -> Self:
        if len(input) > MAX_WORKBOOK_INPUT_SIZE:
            raise GraphQLError("Workbook input too large")

        workbook = models.Workbook.objects.select_for_update().get(id=workbook_id)

        if sheet_name is not None:
            assert sheet_id is None
            sheet = workbook.calc.sheets[sheet_name]
        else:
            assert sheet_id is not None
            sheet = workbook.calc.sheets.get_sheet_by_id(sheet_id)

        if ref is not None:
            assert row is None and col is None
            cell = sheet[ref]
        else:
            assert row is not None and col is not None
            cell = sheet.cell(row, col)

        cell.set_user_input(input)

        workbook.set_workbook_json(workbook.calc.json)
        return SetCellInput(workbook=workbook)


class CreateSheet(graphene.Mutation):
    class Arguments:
        workbook_id = graphene.String(required=True)
        sheet_name = graphene.String()

    workbook = graphene.Field(Workbook, required=True)
    sheet = graphene.Field(Sheet, required=True)

    @classmethod
    @validate_license_for_workbook_mutation
    def mutate(cls, root: Any, info: graphene.ResolveInfo, workbook_id: str, sheet_name: str | None = None) -> Self:
        workbook = models.Workbook.objects.select_for_update().get(id=workbook_id)

        calc_sheet = workbook.calc.sheets.add(sheet_name)

        workbook.set_workbook_json(workbook.calc.json)
        return cls(workbook=workbook, sheet=Sheet(calc_sheet=calc_sheet))


class DeleteSheet(graphene.Mutation):
    class Arguments:
        workbook_id = graphene.String(required=True)
        sheet_id = graphene.Int(required=True)

    workbook = graphene.Field(Workbook, required=True)

    @classmethod
    @validate_license_for_workbook_mutation
    def mutate(cls, root: Any, info: graphene.ResolveInfo, workbook_id: str, sheet_id: int) -> Self:
        workbook = models.Workbook.objects.select_for_update().get(id=workbook_id)

        workbook.calc.sheets.get_sheet_by_id(sheet_id).delete()

        workbook.set_workbook_json(workbook.calc.json)
        return cls(workbook=workbook)


class RenameSheet(graphene.Mutation):
    class Arguments:
        workbook_id = graphene.String(required=True)
        sheet_id = graphene.Int(required=True)
        new_name = graphene.String(required=True)

    workbook = graphene.Field(Workbook, required=True)
    sheet = graphene.Field(Sheet, required=True)

    @classmethod
    @validate_license_for_workbook_mutation
    def mutate(cls, root: Any, info: graphene.ResolveInfo, workbook_id: str, sheet_id: int, new_name: str) -> Self:
        workbook = models.Workbook.objects.select_for_update().get(id=workbook_id)

        calc_sheet = workbook.calc.sheets.get_sheet_by_id(sheet_id)
        calc_sheet.name = new_name

        workbook.set_workbook_json(workbook.calc.json)
        return cls(workbook=workbook, sheet=Sheet(calc_sheet=calc_sheet))


class CreateWorkbook(graphene.Mutation):
    # The class attributes define the response of the mutation
    workbook = graphene.Field(Workbook, required=True)

    @classmethod
    def mutate(cls, root: Any, info: graphene.ResolveInfo) -> Self:
        log.info("MUTATE CREATE WORKBOOK")
        # The /graphiql client doesn't seem to be inserting the Authorization header
        license = get_license(info.context.META)
        origin = info.context.META.get("HTTP_ORIGIN")
        log.info(
            "CreateWorkbook: auth=%s, license=%s, origin=%s"
            % (info.context.META.get("HTTP_AUTHORIZATION"), license, origin),
        )
        if not is_license_key_valid_for_host(license.key, origin):
            log.error("License key %s is not valid for %s." % (license.key, origin))
            raise GraphQLError("Invalid license key")
        if models.Workbook.objects.filter(license=license).count() >= MAX_WORKBOOKS_PER_LICENSE:
            log.error(
                "License key cannot be used to create any more workbooks (MAX_WORKBOOKS_PER_LICENSE=%s)"
                % MAX_WORKBOOKS_PER_LICENSE,
            )
            raise GraphQLError(
                "You cannot create more than %s workbooks with this license key." % MAX_WORKBOOKS_PER_LICENSE,
            )

        workbook = models.Workbook(license=license)
        workbook.save()
        # Notice we return an instance of this mutation
        return CreateWorkbook(workbook=workbook)


class Mutation(graphene.ObjectType):
    save_workbook = SaveWorkbook.Field()
    create_workbook = CreateWorkbook.Field()
    set_cell_input = SetCellInput.Field()
    create_sheet = CreateSheet.Field()
    delete_sheet = DeleteSheet.Field()
    rename_sheet = RenameSheet.Field()


schema = graphene.Schema(query=Query, mutation=Mutation)
