"""Tests for the validation module."""

import os
import tempfile

import pandas as pd
import pytest

from scripts.config import get_excel_columns, SAMPLE_DATA, VALID_STATUSES
from scripts.validate import validate_upload


@pytest.fixture
def tmp_dir():
    with tempfile.TemporaryDirectory() as d:
        yield d


def _create_excel(path, data):
    """Helper to create an Excel file from a list of dicts."""
    df = pd.DataFrame(data)
    df.to_excel(path, index=False, engine="openpyxl")
    return path


class TestValidateAppAdoption:
    def test_valid_file(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        _create_excel(path, SAMPLE_DATA)
        is_valid, messages = validate_upload(path)
        assert is_valid

    def test_missing_required_column(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [{"Name of the Dealer": "Test", "Status": "Installed"}]  # Missing Account Sf ID
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        assert not is_valid
        assert any("Account Sf ID" in m for m in messages)

    def test_blank_dealer_name_is_warning(self, tmp_dir):
        """Blank dealer name is a warning (not error) since real data has this."""
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [{"Account Sf ID": "001", "Name of the Dealer": None, "Status": "Installed"}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        assert is_valid
        assert any("blank" in m.lower() for m in messages)

    def test_blank_account_sf_id_is_error(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [{"Account Sf ID": None, "Name of the Dealer": "Test", "Status": "Installed"}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        assert not is_valid

    def test_invalid_status_warning(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [{"Account Sf ID": "001", "Name of the Dealer": "Test", "Status": "Unknown"}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        # Invalid status is a warning, not an error
        assert is_valid
        assert any("unexpected values" in m.lower() for m in messages)

    def test_duplicate_primary_key(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [
            {"Account Sf ID": "001", "Name of the Dealer": "A", "Status": "Installed"},
            {"Account Sf ID": "001", "Name of the Dealer": "B", "Status": "Installed"},
        ]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        assert is_valid  # Duplicates are warnings
        assert any("duplicate" in m.lower() for m in messages)

    def test_extra_columns_warning(self, tmp_dir):
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [{"Account Sf ID": "001", "Name of the Dealer": "Test", "Status": "Installed", "Extra Col": "x"}]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        assert is_valid
        assert any("unexpected" in m.lower() for m in messages)

    def test_dash_in_numeric_column(self, tmp_dir):
        """'-' is a valid placeholder in numeric columns and should not cause errors."""
        path = os.path.join(tmp_dir, "test.xlsx")
        data = [{
            "Account Sf ID": "001", "Name of the Dealer": "Test", "Status": "Installed",
            "Order quantity (total)": "-", "Quantity (via. app.)": "-",
        }]
        _create_excel(path, data)
        is_valid, messages = validate_upload(path)
        assert is_valid


class TestEdgeCases:
    def test_file_not_found(self):
        is_valid, messages = validate_upload("/nonexistent/file.xlsx")
        assert not is_valid
        assert any("not found" in m.lower() for m in messages)

    def test_empty_file(self, tmp_dir):
        path = os.path.join(tmp_dir, "empty.xlsx")
        pd.DataFrame().to_excel(path, index=False, engine="openpyxl")
        is_valid, messages = validate_upload(path)
        assert not is_valid
        assert any("empty" in m.lower() for m in messages)
