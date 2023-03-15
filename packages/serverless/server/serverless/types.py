from typing import Dict, List, Union

SimulateInputType = Dict[str, Dict[str, Union[bool, float, int, str]]]
SimulateOutputType = Dict[str, List[str]]
SimulateResultType = SimulateInputType
