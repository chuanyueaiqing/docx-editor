# DOCX Editor Makefile
# For use with GNU Make (available in Git Bash, WSL, or Cygwin)

.PHONY: install install-dev test test-quick coverage clean build

install:
	pip install -e .

install-dev:
	pip install -e "."
	pip install pytest pytest-cov coverage build

test:
	python -m pytest tests/ -v --tb=short -k "not win32"

test-all:
	python -m pytest tests/ -v --tb=short

coverage:
	python -m pytest tests/ --cov=docx_editor --cov-report=term --cov-report=html

clean:
	rm -rf build/ dist/ *.egg-info/
	find . -type d -name __pycache__ -prune -exec rm -rf {} + 2>/dev/null; true
	find . -type f -name '*.pyc' -delete

build:
	python -m build

publish: build
	python -m twine upload dist/*
