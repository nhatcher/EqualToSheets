from util import common, log
import os
from functools import lru_cache
from typing import Dict

# functools.cache is only available on Python 3.9+, and the CI machine image has 3.8
@lru_cache
def env_with_docker_user() -> Dict[str, str]:  # noqa: TAN002 - py3.8 compat
    """Returns a copy of `os.environ` with EQUALTO_DOCKER_USER set."""
    env = os.environ.copy()
    env["EQUALTO_DOCKER_USER"] = f"{os.getuid()}:{os.getgid()}"
    return env

@log.func_context
def run_rust_tests() -> None:
    """Runs the EqualTo Calc tests."""
    env = env_with_docker_user()
    common.run_str("docker-compose run --rm rust make tests", env=env)
