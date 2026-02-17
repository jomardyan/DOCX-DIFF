# Makefile for DOCX Diff development tasks
# Compatible with Windows (using PowerShell) and Unix-like systems

.PHONY: help install install-dev test lint format clean build docs run run-gui

# Determine the OS
ifeq ($(OS),Windows_NT)
    PYTHON := python
    RM := powershell -Command "Remove-Item -Recurse -Force"
else
    PYTHON := python3
    RM := rm -rf
endif

help:
	@echo "DOCX Diff - Development Commands"
	@echo ""
	@echo "  make install      - Install production dependencies"
	@echo "  make install-dev  - Install development dependencies"
	@echo "  make test         - Run tests with coverage"
	@echo "  make lint         - Run code quality checks"
	@echo "  make format       - Format code with Black and isort"
	@echo "  make clean        - Remove build artifacts"
	@echo "  make build        - Build distribution packages"
	@echo "  make run          - Run CLI with sample files"
	@echo "  make run-gui      - Run GUI mode"
	@echo ""

install:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements.txt

install-dev:
	$(PYTHON) -m pip install --upgrade pip
	$(PYTHON) -m pip install -r requirements-dev.txt

test:
	$(PYTHON) -m pytest tests/ -v --cov=docxdiff --cov-report=html --cov-report=term

lint:
	$(PYTHON) -m flake8 docxdiff.py --max-line-length=100 --statistics
	$(PYTHON) -m pylint docxdiff.py
	$(PYTHON) -m mypy docxdiff.py --ignore-missing-imports

format:
	$(PYTHON) -m black docxdiff.py --line-length=100
	$(PYTHON) -m isort docxdiff.py --profile black

clean:
	$(RM) build dist *.egg-info
	$(RM) __pycache__ .pytest_cache .mypy_cache
	$(RM) htmlcov .coverage
	$(RM) *.pyc *.pyo

build: clean
	$(PYTHON) setup.py sdist bdist_wheel

run:
	$(PYTHON) docxdiff.py --help

run-gui:
	$(PYTHON) docxdiff.py

check: lint test
	@echo "All checks passed!"

.DEFAULT_GOAL := help
