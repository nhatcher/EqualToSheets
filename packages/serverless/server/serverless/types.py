from typing import Dict, List, Union

CellValue = Union[bool, float, int, str, None]
CellRangeValue = list[list[CellValue]]
SimulateInputType = Dict[str, Dict[str, Union[CellValue | CellRangeValue]]]
SimulateOutputType = Dict[str, List[str]]
SimulateResultType = SimulateInputType
