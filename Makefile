PY := ./.venv/bin/python
PIP := ./.venv/bin/pip

.PHONY: setup run clean

setup:
	@if [ ! -x "$(PY)" ]; then \
		./scripts/bootstrap.sh; \
	else \
		echo ".venv already exists (use .venv/bin/python or source .venv/bin/activate)"; \
	fi

run:
	$(PY) main.py

clean:
	rm -rf .venv downloads
