.PHONY: help venv install install-dev test test-cov lint format type-check clean setup-dev activate

VENV = venv
PYTHON = $(VENV)/bin/python
PIP = $(VENV)/bin/pip

help:  ## Show this help message
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "\033[36m%-20s\033[0m %s\n", $$1, $$2}'

venv:  ## Create virtual environment
	python -m venv $(VENV)
	$(PIP) install --upgrade pip

install: venv  ## Install production dependencies
	$(PIP) install -r requirements.txt

install-dev: venv  ## Install development dependencies
	$(PIP) install -r requirements-dev.txt

setup-dev: install-dev  ## Setup development environment
	$(VENV)/bin/pre-commit install

activate:  ## Show command to activate virtual environment
	@echo "Run: source $(VENV)/bin/activate"

test:  ## Run tests
	$(PYTHON) -m pytest

test-cov:  ## Run tests with coverage
	$(PYTHON) -m pytest --cov=. --cov-report=html --cov-report=term

lint:  ## Run linter
	$(PYTHON) -m ruff check .

format:  ## Format code
	$(PYTHON) -m ruff format .

type-check:  ## Run type checker
	$(PYTHON) -m mypy .

clean:  ## Clean build artifacts
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf .pytest_cache/
	rm -rf .mypy_cache/
	rm -rf htmlcov/
	rm -rf $(VENV)/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

check: lint type-check test  ## Run all checks (lint, type-check, test)

ci: format lint type-check test-cov  ## Run CI pipeline locally
