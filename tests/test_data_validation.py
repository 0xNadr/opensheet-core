"""Tests for data validation (read + write) — issue #8."""

import os
import tempfile

import opensheet_core


def test_write_and_read_list_validation():
    """Write a list validation and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.add_data_validation(
                "list",
                "A1:A100",
                formula1='"Option1,Option2,Option3"',
                allow_blank=True,
                show_input_message=True,
                show_error_message=True,
            )
            w.write_row(["Option1"])

        sheets = opensheet_core.read_xlsx(path)
        dvs = sheets[0]["data_validations"]
        assert len(dvs) == 1
        dv = dvs[0]
        assert dv["type"] == "list"
        assert dv["sqref"] == "A1:A100"
        assert dv["formula1"] == '"Option1,Option2,Option3"'
        assert dv["allow_blank"] is True
        assert dv["show_input_message"] is True
        assert dv["show_error_message"] is True
    finally:
        os.unlink(path)


def test_write_and_read_whole_number_validation():
    """Write a whole number between validation and read it back."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.add_data_validation(
                "whole",
                "B1:B50",
                formula1="1",
                formula2="100",
                operator="between",
                show_error_message=True,
                error_title="Invalid",
                error_message="Enter a number between 1 and 100",
                error_style="stop",
            )
            w.write_row([42])

        sheets = opensheet_core.read_xlsx(path)
        dvs = sheets[0]["data_validations"]
        assert len(dvs) == 1
        dv = dvs[0]
        assert dv["type"] == "whole"
        assert dv["operator"] == "between"
        assert dv["formula1"] == "1"
        assert dv["formula2"] == "100"
        assert dv["error_title"] == "Invalid"
        assert dv["error_message"] == "Enter a number between 1 and 100"
        assert dv["error_style"] == "stop"
    finally:
        os.unlink(path)


def test_multiple_validations():
    """Multiple data validations on the same sheet."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.add_data_validation("list", "A1:A10", formula1='"Yes,No"')
            w.add_data_validation("whole", "B1:B10", formula1="0", formula2="999", operator="between")
            w.add_data_validation("decimal", "C1:C10", formula1="0.0", operator="greaterThan")
            w.write_row(["Yes", 42, 3.14])

        sheets = opensheet_core.read_xlsx(path)
        dvs = sheets[0]["data_validations"]
        assert len(dvs) == 3
        assert dvs[0]["type"] == "list"
        assert dvs[1]["type"] == "whole"
        assert dvs[2]["type"] == "decimal"
    finally:
        os.unlink(path)


def test_validation_with_prompt():
    """Data validation with input prompt."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.add_data_validation(
                "list",
                "A1:A5",
                formula1='"Red,Green,Blue"',
                show_input_message=True,
                prompt_title="Choose color",
                prompt="Select from the list",
            )
            w.write_row(["Red"])

        sheets = opensheet_core.read_xlsx(path)
        dv = sheets[0]["data_validations"][0]
        assert dv["prompt_title"] == "Choose color"
        assert dv["prompt"] == "Select from the list"
    finally:
        os.unlink(path)


def test_no_validations():
    """Sheet with no validations returns empty list."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row(["data"])

        sheets = opensheet_core.read_xlsx(path)
        assert sheets[0]["data_validations"] == []
    finally:
        os.unlink(path)


def test_validation_requires_open_sheet():
    """Adding validation without an open sheet raises an error."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        w = opensheet_core.XlsxWriter(path)
        try:
            w.add_data_validation("list", "A1:A10", formula1='"Yes,No"')
            assert False, "Should have raised"
        except Exception as e:
            assert "No sheet is open" in str(e)
        w.close()
    finally:
        os.unlink(path)
