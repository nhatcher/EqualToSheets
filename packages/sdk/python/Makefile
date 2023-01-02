VENV_NAME?=venv

venv: $(VENV_NAME)/bin/activate

$(VENV_NAME)/bin/activate:
	pip3 install --upgrade pip virtualenv
	@test -d $(VENV_NAME) || python3 -m virtualenv --clear $(VENV_NAME)
	${VENV_NAME}/bin/python -m pip install -e .[dev]
	@touch $(VENV_NAME)/bin/activate

tests: venv
	@${VENV_NAME}/bin/python -m pytest ./tests

coverage: venv
	@${VENV_NAME}/bin/python -m pytest \
		--cov=equalto \
		--cov-report=html:./tests/htmlcov \
		--cov-report=term-missing:skip-covered \
		--no-cov-on-fail \
		--cov-fail-under=100 \
		./tests

lint: venv
	@${VENV_NAME}/bin/python -m mypy ./equalto ./tests
	@${VENV_NAME}/bin/python -m flake8 ./equalto ./tests
	@${VENV_NAME}/bin/python -m black --check ./equalto ./tests
	@${VENV_NAME}/bin/python -m isort --check ./equalto ./tests

fmt: venv
	@${VENV_NAME}/bin/python -m black ./equalto ./tests
	@${VENV_NAME}/bin/python -m isort ./equalto ./tests

clean:
	@rm -rf $(VENV_NAME) target/

.PHONY: venv test clean
