VENV_NAME?=venv

PYTHON3_MINOR_VER = $(shell python3 --version | grep ^Python | sed -E 's/^Python 3\.([0-9]+)\.([0-9]+)$$/\1/g')

RUST_MAJOR_VER = $(shell rustc --version | grep ^rustc | sed -E 's/^rustc ([0-9]+)\.([0-9]+)\.([0-9]+) \(.+\)$$/\1/g')
RUST_MINOR_VER = $(shell rustc --version | grep ^rustc | sed -E 's/^rustc ([0-9]+)\.([0-9]+)\.([0-9]+) \(.+\)$$/\2/g')

PYTHON3_MIN_MINOR_VER = 9  # >= 3.9
RUST_MIN_MINOR_VER = 65  # >= 1.65

IS_SUPPORTED_PYTHON_VERSION = $(shell expr $(PYTHON3_MINOR_VER) ">=" $(PYTHON3_MIN_MINOR_VER))
IS_SUPPORTED_RUST_VERSION = $(shell expr $(RUST_MAJOR_VER) "=" 1 "&" $(RUST_MINOR_VER) ">=" $(RUST_MIN_MINOR_VER))

ifeq "$(IS_SUPPORTED_PYTHON_VERSION)" "0"
    $(error Unsupported python interpreter version, please use python >= 3.$(PYTHON3_MIN_MINOR_VER))
endif

ifeq "$(IS_SUPPORTED_RUST_VERSION)" "0"
    $(error Unsupported rust compiler version, please use rust >= 1.$(RUST_MIN_MINOR_VER))
endif

venv: clean $(VENV_NAME)/bin/activate

$(VENV_NAME)/bin/activate:
	pip3 install --upgrade pip virtualenv
	@test -d $(VENV_NAME) || python3 -m virtualenv --clear $(VENV_NAME)
	${VENV_NAME}/bin/python -m pip install -e .[dev]
	@touch $(VENV_NAME)/bin/activate

tests: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m pytest ./tests

coverage: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m pytest \
		--cov=equalto \
		--cov-report=html:./tests/htmlcov \
		--cov-report=term-missing:skip-covered \
		--no-cov-on-fail \
		--cov-fail-under=100 \
		./tests

lint: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m mypy ./equalto ./tests
	@${VENV_NAME}/bin/python -m flake8 ./equalto ./tests
	@${VENV_NAME}/bin/python -m black --check ./equalto ./tests
	@${VENV_NAME}/bin/python -m isort --check ./equalto ./tests
	cargo fmt --check
	cargo clippy --all-targets --all-features -- -D warnings

fmt: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m black ./equalto ./tests
	@${VENV_NAME}/bin/python -m isort ./equalto ./tests
	cargo fmt

clean:
	@rm -rf $(VENV_NAME) target/ dist/

.PHONY: venv tests coverage lint fmt clean
