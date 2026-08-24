#!/usr/bin/env python3
import sys
import os

# Ensure project directory is in python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.cli import app

if __name__ == "__main__":
    app()
