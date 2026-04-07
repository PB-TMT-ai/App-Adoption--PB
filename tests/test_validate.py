"""Tests for the validation module."""

import os
import tempfile

import pandas as pd
import pytest

from scripts.config import get_column_names, SAMPLE_DATA
from scripts.validate import validate_upload, normalize_columns


@pytest.fixture
def tmp_dir():
    with tempfile.TemporaryDirectory() as d:
        yield d


def _create_excel(path, data, columns=None):
    """Helper to create an Excel file from a list of dicts."""
    df = pd.DataFrame(data)
    if columns:
        df = df[columns]
    df.to_excel(path, index=False, engine="openpyxl")
    return path


class TestValidateDealers:
    def test_valid_dealer_file(self, tmp_dir):
        path = os.path.join(tmp_dir, "dealers.xlsx")
        _create_excel(path, SAMPLE_DATA["dealers"])
        is_valid, messages = validate_upload(path, "dealers")
        assert is_valid
        assert any("passed" in m.lower() for m in messages)

    def test_missing_required_column(self, tmp_dir):
        path = os.path.join(tmp_dir, "dealers.xlsx")
        data = [{"dealer_name": "Test Dealer"}]  # Missing dealer_id
        _create_excel(path, data)
        is_valid, messages = validate_upload(path, "dealers")
        assert not is_valid
        assert any("dealer_id" in m for m in messages)

    def test_blank_required_field(self, tmp_dir):
        path = os.path.join(tmp_dir, "dealers.xlsx")
        data = [{"dealer_id": "DLR001", "dealer_name": None}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path, "dealers")
        assert not is_valid
        assert any("blank" in m.lower() for m in messages)

    def test_duplicate_primary_key(self, tmp_dir):
        path = os.path.join(tmp_dir, "dealers.xlsx")
        data = [
            {"dealer_id": "DLR001", "dealer_name": "Dealer A"},
            {"dealer_id": "DLR001", "dealer_name": "Dealer B"},
        ]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path, "dealers")
        # Duplicates are warnings, not errors
        assert is_valid
        assert any("duplicate" in m.lower() for m in messages)

    def test_extra_columns_warning(self, tmp_dir):
        path = os.path.join(tmp_dir, "dealers.xlsx")
        data = [{"dealer_id": "DLR001", "dealer_name": "Test", "extra_col": "value"}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path, "dealers")
        assert is_valid
        assert any("unexpected" in m.lower() or "extra_col" in m for m in messages)


class TestValidateInstallEvents:
    def test_valid_install_events(self, tmp_dir):
        path = os.path.join(tmp_dir, "events.xlsx")
        _create_excel(path, SAMPLE_DATA["install_events"])
        is_valid, messages = validate_upload(path, "install_events")
        assert is_valid

    def test_invalid_event_type(self, tmp_dir):
        path = os.path.join(tmp_dir, "events.xlsx")
        data = [{"dealer_id": "DLR001", "event_type": "upgrade", "event_date": "2025-01-01"}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path, "install_events")
        assert not is_valid
        assert any("event_type" in m for m in messages)


class TestValidateUsageMetrics:
    def test_valid_usage_metrics(self, tmp_dir):
        path = os.path.join(tmp_dir, "usage.xlsx")
        _create_excel(path, SAMPLE_DATA["usage_metrics"])
        is_valid, messages = validate_upload(path, "usage_metrics")
        assert is_valid


class TestEdgeCases:
    def test_file_not_found(self):
        is_valid, messages = validate_upload("/nonexistent/file.xlsx", "dealers")
        assert not is_valid
        assert any("not found" in m.lower() for m in messages)

    def test_unknown_dataset_type(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        _create_excel(path, [{"a": 1}])
        is_valid, messages = validate_upload(path, "unknown_type")
        assert not is_valid

    def test_empty_file(self, tmp_dir):
        path = os.path.join(tmp_dir, "empty.xlsx")
        pd.DataFrame().to_excel(path, index=False, engine="openpyxl")
        is_valid, messages = validate_upload(path, "dealers")
        assert not is_valid
        assert any("empty" in m.lower() for m in messages)


class TestNormalizeColumns:
    def test_normalize(self):
        df = pd.DataFrame(columns=["Dealer ID", " Event Type ", "APP_VERSION"])
        df = normalize_columns(df)
        assert list(df.columns) == ["dealer_id", "event_type", "app_version"]
