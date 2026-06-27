"""WSGI entry point: gunicorn imports ``summa.wsgi:app``."""

from flask import Flask

from summa import create_app

app: Flask = create_app()
