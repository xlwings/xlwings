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
	uv sync --group all
	uv run sphinx-autobuild docs docs/_build/html --port 9000 -E
