.PHONY: test
test:
	uv sync --group all
	uv run pytest

.PHONY: lint
lint:
	uv sync --group all
	uv run pre-commit run --all-files

.PHONY: docs
docs:
	uv run --group all sphinx-autobuild docs docs/_build/html --port 9000 -a -D llms_txt_enabled=0 --watch xlwings
