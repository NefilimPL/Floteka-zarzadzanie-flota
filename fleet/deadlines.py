"""Wspólne kolory i granice terminów rozładunku (lokalne dni, bez godzin)."""

from datetime import date

OVERDUE_FILL = 'E53935'
SOON_FILL = 'FFA726'


def deadline_tag(value, today=None):
    if not value:
        return ''
    days = (date.fromisoformat(value) - (today or date.today())).days
    if days < 0:
        return 'overdue'
    return 'soon' if days <= 1 else ''
