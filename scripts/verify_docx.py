#!/usr/bin/env python
"""Entry point: py scripts/verify_docx.py path/to/file.docx [--json]"""
import importlib.util
import os
import sys

_PROJECT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
_SCRIPT = os.path.join(_PROJECT, 'skills', 'docx-mode', 'scripts', 'verify_docx.py')
sys.path.insert(0, _PROJECT)

spec = importlib.util.spec_from_file_location('verify_docx_main', _SCRIPT)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
mod.main()
