"""
Kernel Startup Script — Pre-loaded Common Libraries
=====================================================

This module provides the startup code that runs automatically when
a new sandbox kernel is created. It pre-imports commonly needed
libraries so they're immediately available for execute_code calls.

All imports are wrapped in try/except so that missing packages
don't prevent kernel startup — the kernel will still work, it
just won't have that particular library pre-loaded.

Usage:
    This is imported and used by PersistentKernelManager in
    fix_kernel_persistence.py. You don't need to call it directly.

    If you want to customize which libraries are pre-loaded,
    edit the KERNEL_STARTUP_CODE string below.
"""


# =============================================================================
# STARTUP CODE — This string is executed in every new kernel
# =============================================================================

KERNEL_STARTUP_CODE = '''
# =============================================================
# MCP Sandbox Kernel — Automatic Library Pre-load
# =============================================================
# This runs once when the kernel is created.
# All imports use try/except — missing packages are logged
# but don't prevent kernel startup.
# =============================================================

import os
import sys
import json
import csv
import base64
import io
import re
import pathlib
import traceback
import hashlib
from datetime import datetime, timedelta
from collections import defaultdict, Counter

# ----- PDF Libraries -----
try:
    import fitz  # PyMuPDF — primary PDF handler
    _fitz_available = True
except ImportError:
    fitz = None
    _fitz_available = False

try:
    import PyPDF2
    _pypdf2_available = True
except ImportError:
    PyPDF2 = None
    _pypdf2_available = False

try:
    import pdfplumber
    _pdfplumber_available = True
except ImportError:
    pdfplumber = None
    _pdfplumber_available = False

# ----- Data Libraries -----
try:
    import pandas as pd
    _pandas_available = True
except ImportError:
    pd = None
    _pandas_available = False

try:
    import numpy as np
    _numpy_available = True
except ImportError:
    np = None
    _numpy_available = False

try:
    import openpyxl  # Excel read/write
    _openpyxl_available = True
except ImportError:
    openpyxl = None
    _openpyxl_available = False

# ----- Document Libraries -----
try:
    from docx import Document as DocxDocument  # python-docx
    _docx_available = True
except ImportError:
    DocxDocument = None
    _docx_available = False

# ----- Image Libraries -----
try:
    from PIL import Image
    _pillow_available = True
except ImportError:
    Image = None
    _pillow_available = False

# ----- HTTP/API Libraries -----
try:
    import requests
    _requests_available = True
except ImportError:
    requests = None
    _requests_available = False

# =============================================================
# Pre-load Summary (stored for introspection)
# =============================================================

__sandbox_preloaded__ = {
    "fitz (PyMuPDF)": _fitz_available,
    "PyPDF2": _pypdf2_available,
    "pdfplumber": _pdfplumber_available,
    "pandas": _pandas_available,
    "numpy": _numpy_available,
    "openpyxl": _openpyxl_available,
    "python-docx": _docx_available,
    "Pillow": _pillow_available,
    "requests": _requests_available,
}

# =============================================================
# Sandbox Data Directory
# =============================================================

SANDBOX_DATA_DIR = pathlib.Path("/app/sandbox_data")
if not SANDBOX_DATA_DIR.exists():
    SANDBOX_DATA_DIR.mkdir(parents=True, exist_ok=True)

# =============================================================
# Helper Functions (available in all execute_code calls)
# =============================================================

def list_sandbox_files(extension=None):
    """List files in the sandbox data directory, optionally filtered by extension."""
    files = []
    for f in SANDBOX_DATA_DIR.iterdir():
        if extension and not f.suffix.lower() == extension.lower():
            continue
        files.append({
            "name": f.name,
            "size_bytes": f.stat().st_size,
            "size_human": _human_size(f.stat().st_size),
            "extension": f.suffix,
            "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat(),
        })
    return sorted(files, key=lambda x: x["name"])


def list_sandbox_pdfs():
    """List all PDF files in the sandbox."""
    return list_sandbox_files(extension=".pdf")


def sandbox_status():
    """Print a summary of the sandbox environment."""
    print("=" * 60)
    print("SANDBOX ENVIRONMENT STATUS")
    print("=" * 60)
    print()
    print("Pre-loaded Libraries:")
    for lib, available in __sandbox_preloaded__.items():
        status = "✅" if available else "❌ Not installed"
        print(f"  {lib}: {status}")
    print()
    print(f"Sandbox Data Directory: {SANDBOX_DATA_DIR}")
    files = list_sandbox_files()
    print(f"Files in sandbox: {len(files)}")
    for f in files:
        print(f"  {f['name']} ({f['size_human']})")
    print()
    print(f"Python version: {sys.version}")
    print("=" * 60)


def _human_size(size_bytes):
    """Convert bytes to human-readable size string."""
    for unit in ["B", "KB", "MB", "GB"]:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def read_pdf_text(filepath, method="auto"):
    """
    Extract text from a PDF file using the best available library.
    
    Args:
        filepath: Path to the PDF file (string or Path)
        method: "auto" (try best available), "fitz", "pypdf2", or "pdfplumber"
    
    Returns:
        String containing all extracted text
    """
    filepath = str(filepath)
    
    if method == "auto":
        if _fitz_available:
            method = "fitz"
        elif _pdfplumber_available:
            method = "pdfplumber"
        elif _pypdf2_available:
            method = "pypdf2"
        else:
            raise ImportError(
                "No PDF library available. Install one of: "
                "PyMuPDF (fitz), pdfplumber, PyPDF2"
            )
    
    if method == "fitz":
        doc = fitz.open(filepath)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text
    
    elif method == "pdfplumber":
        with pdfplumber.open(filepath) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\\n"
        return text
    
    elif method == "pypdf2":
        reader = PyPDF2.PdfReader(filepath)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\\n"
        return text
    
    else:
        raise ValueError(f"Unknown PDF method: {method}. Use 'auto', 'fitz', 'pypdf2', or 'pdfplumber'")


# Print startup confirmation (visible in kernel logs)
_available_count = sum(1 for v in __sandbox_preloaded__.values() if v)
_total_count = len(__sandbox_preloaded__)
print(f"[Kernel Ready] {_available_count}/{_total_count} libraries pre-loaded | sandbox: {SANDBOX_DATA_DIR}")
'''


# =============================================================================
# DIAGNOSTIC FUNCTION (can be called from main.py for troubleshooting)
# =============================================================================

def get_startup_code() -> str:
    """Return the kernel startup code string."""
    return KERNEL_STARTUP_CODE


def get_diagnostic_code() -> str:
    """
    Return a diagnostic code snippet that can be executed in a kernel
    to check what's available and working.
    """
    return '''
# Run this to check sandbox health
sandbox_status()
print()
print("PDF files available:")
for pdf in list_sandbox_pdfs():
    print(f"  {pdf['name']} ({pdf['size_human']})")
print()
print("Quick PDF read test:")
pdfs = list_sandbox_pdfs()
if pdfs:
    try:
        text = read_pdf_text(f"/app/sandbox_data/{pdfs[0]['name']}")
        print(f"  ✅ Successfully read {pdfs[0]['name']}")
        print(f"  First 200 chars: {text[:200]}...")
    except Exception as e:
        print(f"  ❌ Failed: {e}")
else:
    print("  No PDF files in sandbox to test")
'''
