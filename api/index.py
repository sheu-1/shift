"""
Vercel WSGI entry point.
Vercel's @vercel/python runtime discovers the `app` (Flask) callable from this file.
"""
import sys
import os

# Ensure the parent directory (shift-scheduler/) is on sys.path so `app` can be imported
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app  # noqa: F401  – Vercel looks for `app` in this module

# Alias so Vercel can also find it as `handler`
handler = app
