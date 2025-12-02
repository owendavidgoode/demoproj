.PHONY: install dev test lint build clean

PYTHON := python3
VENV := .venv
BIN := $(VENV)/bin
DRIVER_DIR := bin

install:
	$(PYTHON) -m venv $(VENV)
	$(BIN)/pip install --upgrade pip
	$(BIN)/pip install -r requirements.txt
	@echo "Note: If you need a portable chromedriver, place it in $(DRIVER_DIR)/"

dev:
	$(BIN)/python src/cli/main.py --help

test:
	$(BIN)/pytest tests/

lint:
	$(BIN)/pip install ruff
	$(BIN)/ruff check src/ tests/

build:
	@echo "Building distribution..."
	mkdir -p dist
	zip -r dist/local-inventory-tool.zip src/ config/ requirements.txt README.md Makefile $(DRIVER_DIR)

clean:
	rm -rf $(VENV) dist/
	find . -type d -name "__pycache__" -exec rm -rf {} +
