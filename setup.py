#!/usr/bin/env python
"""
Setup script for OfficeReader-MCP.

This is a minimal setup.py for backward compatibility with older tools that don't support pyproject.toml.
All configuration is read from pyproject.toml.

For modern installations, use:
    pip install .
or for development:
    pip install -e .
"""

from setuptools import setup

# All configuration is in pyproject.toml
# This setup.py exists only for backward compatibility
setup()
