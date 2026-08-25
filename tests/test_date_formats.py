from datetime import date, datetime

import pytest

from excel_template_writer.date_formats import (
    DateFormatSyntaxError,
    DateFormatValueError,
    format_date,
    parse_date_format,
)


def test_formats_every_supported_date_field_deterministically() -> None:
    date_format = parse_date_format("YYYY|yy|MMMM|mmm|MM|m|DDDD|ddd|DD|d")

    rendered = format_date(date(2026, 8, 31), date_format)

    assert rendered == "2026|26|August|Aug|08|8|Monday|Mon|31|31"


def test_formats_datetime_calendar_fields_and_explicit_literals() -> None:
    date_format = parse_date_format('dd "of" mmmm yyyy \\Q')

    rendered = format_date(datetime(2026, 8, 31, 23, 45), date_format)

    assert rendered == "31 of August 2026 Q"


@pytest.mark.parametrize(
    ("source", "message"),
    [
        ("", "cannot be empty"),
        ("---", "at least one date field"),
        ("yyy-mm", "unsupported width"),
        ("yyyy-qq", "unsupported date field"),
        ('yyyy "month', "unterminated quoted literal"),
        ("yyyy-\\", "dangling escape"),
        ("{yyyy}-{mm}", "unsupported unquoted character"),
        ("%Y-%m", "unsupported unquoted character"),
        ("yyyy;mm", "unsupported unquoted character"),
    ],
)
def test_rejects_ambiguous_or_unsupported_date_formats(source: str, message: str) -> None:
    with pytest.raises(DateFormatSyntaxError, match=message):
        parse_date_format(source)


@pytest.mark.parametrize("value", [None, "2026-08-31", 46265])
def test_rejects_values_that_are_not_native_dates(value: object) -> None:
    with pytest.raises(DateFormatValueError, match="requires a date or datetime"):
        format_date(value, parse_date_format("yyyy-mm-dd"))
