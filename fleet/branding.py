"""Publiczna nazwa i wersja produktu."""

APP_NAME = 'Floteka'
DEFAULT_VERSION = '1.1.2'
try:
    from ._build_version import VERSION
except ModuleNotFoundError as exc:
    if exc.name != 'fleet._build_version':
        raise
    VERSION = DEFAULT_VERSION
