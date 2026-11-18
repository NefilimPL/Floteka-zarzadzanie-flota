from datetime import date

import pytest

from fleet.deadlines import deadline_tag


@pytest.mark.parametrize(('value', 'expected'), [
    ('2026-11-16', 'overdue'), ('2026-11-17', 'soon'),
    ('2026-11-18', 'soon'), ('2026-11-19', ''), ('', ''), (None, ''),
])
def test_deadline_boundaries(value, expected):
    assert deadline_tag(value, today=date(2026, 11, 17)) == expected
