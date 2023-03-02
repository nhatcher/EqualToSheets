from typing import Any, Callable, Iterable, Self

import graphene
from graphene_django import DjangoObjectType
from graphql import GraphQLError

from serverless import log, models
from serverless.util import is_license_key_valid_for_host


def get_license(info: graphene.ResolveInfo) -> models.License:
    auth = info.context.META.get("HTTP_AUTHORIZATION", None)
    if auth is None or auth[:7] != "Bearer ":
        raise GraphQLError("Invalid license key")
    license_key = auth[7:]
    qs_license = models.License.objects.filter(key=license_key)
    if qs_license.count() == 1:
        return qs_license.get()
    log.error("Could not find license for license key %s, qs_license.count()=%s" % (license_key, qs_license.count()))
    raise GraphQLError("Invalid license key")


def _validate_license_for_workbook(info: graphene.ResolveInfo, workbook_id: str) -> None:
    license = get_license(info)
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


class Cell(graphene.ObjectType):
    text = graphene.String()
    number = graphene.Float()
    format = graphene.String()
    boolean = graphene.Boolean()
    type = graphene.String()
    formula = graphene.String()


class Sheet(graphene.ObjectType):
    name = graphene.String()
    id = graphene.String()

    cell = graphene.Field(Cell, ref=graphene.String(), row=graphene.Int(), col=graphene.Int())

    def resolve_cell(
        self,
        info: graphene.ResolveInfo,
        ref: str | None = None,
        row: int | None = None,
        col: int | None = None,
    ) -> Cell:
        if ref is not None and row is None and col is None:
            # resolving ref (eg "A1")
            # TODO: extract from workbook json
            return Cell(
                text="$1.23",
                number=1.234,
                format="$#,##0.00",
                boolean=None,
                type="float",
            )
        elif ref is None and row is not None and col is not None:
            # resolving row/col
            # TODO: extract from workbook JSON
            return Cell(
                text="$1.23",
                number=1.234,
                format="$#,##0.00",
                boolean=None,
                type="float",
            )
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
        )

    sheet = graphene.Field(Sheet, sheet_id=graphene.String(), name=graphene.String())
    sheets = graphene.List(Sheet)

    def resolve_sheet(self, info: graphene.ResolveInfo, sheet_id: str | None = None, name: str | None = None) -> Sheet:
        if sheet_id is not None and name is None:
            return Sheet(name="Sheet1", id=sheet_id)
        elif sheet_id is None and name is not None:
            return Sheet(name=name, id="id-1234-56")
        else:
            log.error("ERROR - name/id")
            raise Exception("ERROR - name/id")

    def resolve_sheets(self, info: graphene.ResolveInfo) -> Iterable[Sheet]:
        log.info("Origin: %s" % info.context.META.get("Origin"))
        # TODO: read from workbook JSON
        return [
            Sheet(name="Sheet1", id="id-1"),
            Sheet(name="Sheet1", id="id-2"),
            Sheet(name="Sheet1", id="id-3"),
        ]


class Query(graphene.ObjectType):
    workbook = graphene.Field(Workbook, workbook_id=graphene.String(required=True))
    workbooks = graphene.List(Workbook)

    @validate_license_for_workbook_query
    def resolve_workbook(self, info: graphene.ResolveInfo, workbook_id: str) -> models.Workbook:
        return models.Workbook.objects.get(id=workbook_id)

    def resolve_workbooks(self, info: graphene.ResolveInfo) -> Iterable[models.Workbook]:
        license = get_license(info)
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
        workbook_json: dict[str, Any],
    ) -> Self:
        # TODO: validate the JSON
        workbook = models.Workbook.objects.get(id=workbook_id)
        workbook.workbook_json = workbook_json
        workbook.revision += 1
        workbook.save()

        return SaveWorkbook(revision=workbook.revision)


class SetCellInput(graphene.Mutation):
    # simulates entering text in the spreadsheet widget
    class Arguments:
        workbook_id = graphene.String(required=True)
        sheet_id = graphene.String()
        sheet_name = graphene.String()
        ref = graphene.String()
        row = graphene.Int()
        col = graphene.Int()

        input = graphene.String(required=True)

    workbook = graphene.Field(Workbook)

    @classmethod
    @validate_license_for_workbook_mutation
    def mutate(
        cls,
        root: Any,
        info: graphene.ResolveInfo,
        workbook_id: str,
        input: str,
        sheet_id: str | None = None,
        sheet_name: str | None = None,
        ref: str | None = None,
        row: int | None = None,
        col: int | None = None,
    ) -> Self:
        # sheet_id XOR sheet_name must be specified
        assert (sheet_id is None and sheet_name is not None) or (sheet_id is not None and sheet_name is None)
        # ref XOR (row, col) must be specified
        assert (ref is None and row is not None and col is not None) or (
            ref is not None and row is None and col is None
        )
        workbook = models.Workbook.objects.get(id=workbook_id)
        # TODO - modify workbook data and reeval workbook
        workbook.revision += 1
        workbook.save()
        return SetCellInput(workbook=workbook)


class CreateWorkbook(graphene.Mutation):
    # The class attributes define the response of the mutation
    workbook = graphene.Field(Workbook, required=True)

    @classmethod
    def mutate(cls, root: Any, info: graphene.ResolveInfo) -> Self:
        log.info("MUTATE CREATE WORKBOOK")
        # The /graphiql client doesn't seem to be inserting the Authorization header
        license = get_license(info)
        origin = info.context.META.get("HTTP_ORIGIN")
        log.info(
            "CreateWorkbook: auth=%s, license=%s, origin=%s"
            % (info.context.META.get("HTTP_AUTHORIZATION"), license, origin),
        )
        if not is_license_key_valid_for_host(license.key, origin):
            log.error("License key %s is not valid for %s." % (license.key, origin))
            raise GraphQLError("Invalid license key")
        workbook = models.Workbook(license=license)
        workbook.save()
        # Notice we return an instance of this mutation
        return CreateWorkbook(workbook=workbook)


class Mutation(graphene.ObjectType):
    save_workbook = SaveWorkbook.Field()
    create_workbook = CreateWorkbook.Field()
    set_cell_input = SetCellInput.Field()


schema = graphene.Schema(query=Query, mutation=Mutation)
