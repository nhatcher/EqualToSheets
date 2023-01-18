class PyCalcModel:
    def evaluate(self) -> None: ...
    def get_cell_value_by_index(self, sheet: int, row: int, column: int) -> str: ...
    def get_cell_type(self, sheet: int, row: int, column: int) -> int: ...
    def has_formula(self, sheet: int, row: int, column: int) -> bool: ...
    def get_formula_or_value(self, sheet: int, row: int, column: int) -> str: ...
    def get_formatted_cell_value(self, sheet: int, row: int, column: int) -> str: ...
    def get_worksheet_ids(self) -> list[int]: ...
    def get_worksheet_names(self) -> list[str]: ...
    def update_cell_with_formula(self, sheet: int, row: int, column: int, formula: str) -> None: ...
    def update_cell_with_text(self, sheet: int, row: int, column: int, value: str) -> None: ...
    def update_cell_with_number(self, sheet: int, row: int, column: int, value: float) -> None: ...
    def update_cell_with_bool(self, sheet: int, row: int, column: int, value: bool) -> None: ...
    def add_sheet(self, name: str) -> None: ...
    def new_sheet(self) -> None: ...
    def delete_sheet_by_sheet_id(self, sheet_id: int) -> None: ...
    def rename_sheet(self, sheet: int, new_name: str) -> None: ...
    def delete_cell(self, sheet: int, row: int, column: int) -> None: ...
    def get_timezone(self) -> str: ...
    def save_to_xlsx(self, file: str) -> None: ...
    def get_style_for_cell(self, sheet: int, row: int, column: int) -> str: ...
    def set_cell_style(self, sheet: int, row: int, column: int, style: str) -> None: ...

def create(name: str, locale: str, tz: str) -> PyCalcModel: ...
def load_excel(workbook_path: str, locale: str, tz: str) -> PyCalcModel: ...

class WorkbookError(Exception):
    """A generic workbook error."""
