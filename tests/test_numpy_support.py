"""Tests for NumPy type support — issue #18."""

import datetime
import math
import os
import tempfile

import numpy as np

import opensheet_core


def test_numpy_int_types():
    """numpy integer types are written correctly."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row([np.int8(1), np.int16(2), np.int32(3), np.int64(4)])
            w.write_row([np.uint8(10), np.uint16(20), np.uint32(30), np.uint64(40)])

        rows = opensheet_core.read_sheet(path)
        assert rows[0] == [1, 2, 3, 4]
        assert rows[1] == [10, 20, 30, 40]
    finally:
        os.unlink(path)


def test_numpy_float_types():
    """numpy float types are written correctly."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row([np.float32(1.5), np.float64(2.5)])

        rows = opensheet_core.read_sheet(path)
        assert abs(rows[0][0] - 1.5) < 0.01
        assert rows[0][1] == 2.5
    finally:
        os.unlink(path)


def test_numpy_bool():
    """numpy.bool_ is written as boolean, not number."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row([np.bool_(True), np.bool_(False)])

        rows = opensheet_core.read_sheet(path)
        assert rows[0][0] is True
        assert rows[0][1] is False
    finally:
        os.unlink(path)


def test_numpy_str():
    """numpy.str_ is written as string."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row([np.str_("hello"), np.str_("world")])

        rows = opensheet_core.read_sheet(path)
        assert rows[0] == ["hello", "world"]
    finally:
        os.unlink(path)


def test_numpy_nan():
    """numpy NaN is written as empty cell (skipped in XLSX)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            # NaN between real values: value, NaN, value
            w.write_row(["start", np.float64("nan"), "end"])

        rows = opensheet_core.read_sheet(path)
        # NaN becomes empty cell (None in the gap)
        assert rows[0][0] == "start"
        assert rows[0][1] is None  # NaN → Empty → None
        assert rows[0][2] == "end"
    finally:
        os.unlink(path)


def test_numpy_inf():
    """numpy infinity is written as string."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row([np.float64("inf"), np.float64("-inf"), float("inf")])

        rows = opensheet_core.read_sheet(path)
        assert rows[0] == ["Infinity", "-Infinity", "Infinity"]
    finally:
        os.unlink(path)


def test_numpy_datetime64():
    """numpy.datetime64 is written as datetime."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        dt = np.datetime64("2025-06-15T14:30:00")
        with opensheet_core.XlsxWriter(path) as w:
            w.add_sheet("Sheet1")
            w.write_row([dt])

        rows = opensheet_core.read_sheet(path)
        result = rows[0][0]
        assert isinstance(result, datetime.datetime)
        assert result.year == 2025
        assert result.month == 6
        assert result.day == 15
        assert result.hour == 14
        assert result.minute == 30
    finally:
        os.unlink(path)
