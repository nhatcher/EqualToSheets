from serverless.rest_api.v1 import CreateWorkbookType


class CreateWorkbookTypeConverter:
    regex = ".*"

    def to_python(self, value: str) -> CreateWorkbookType:
        return CreateWorkbookType(value)

    def to_url(self, create_workbook_type: CreateWorkbookType) -> str:
        return create_workbook_type.value
