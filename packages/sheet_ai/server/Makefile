VENV_NAME?=venv
LINT_FILES=./sheet_ai ./tests ./wsgi

venv: clean $(VENV_NAME)/bin/activate

$(VENV_NAME)/bin/activate:
	pip3 install --upgrade pip virtualenv
	@test -d $(VENV_NAME) || python3 -m virtualenv --clear $(VENV_NAME)
	${VENV_NAME}/bin/python -m pip install -r requirements.txt
	${VENV_NAME}/bin/python -m pip install -r dev-requirements.txt
	${VENV_NAME}/bin/python -m pip install ../../sdk/python
	@touch $(VENV_NAME)/bin/activate

tests: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m pytest ./tests

lint: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m mypy ${LINT_FILES}
	@${VENV_NAME}/bin/python -m flake8 ${LINT_FILES}
	@${VENV_NAME}/bin/python -m black --check ${LINT_FILES}
	@${VENV_NAME}/bin/python -m isort --check ${LINT_FILES}

fmt: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m black ${LINT_FILES}
	@${VENV_NAME}/bin/python -m isort ${LINT_FILES}

compile-requirements: $(VENV_NAME)/bin/activate
	@${VENV_NAME}/bin/python -m pip install pip-tools
	@${VENV_NAME}/bin/pip-compile -o requirements.txt requirements.in --resolver=backtracking
	@${VENV_NAME}/bin/pip-compile -o dev-requirements.txt dev-requirements.in --resolver=backtracking

clean:
	@rm -rf ${VENV_NAME}


.PHONY: venv tests lint fmt clean
