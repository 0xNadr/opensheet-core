import datetime
import tempfile
import pytest

pd = pytest.importorskip("pandas")
import numpy as np
import opensheet_core


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


class TestReadXlsxDf:
    def test_basic_read(self, tmp_xlsx):
        """Read a simple sheet into a DataFrame with header."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["Name", "Age", "Active"])
            w.write_row(["Alice", 30, True])
            w.write_row(["Bob", 25, False])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(df.columns) == ["Name", "Age", "Active"]
        assert len(df) == 2
        assert df.iloc[0]["Name"] == "Alice"
        assert df.iloc[0]["Age"] == 30
        assert df.iloc[1]["Active"] == False

    def test_read_no_header(self, tmp_xlsx):
        """Read without treating first row as header."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["a", "b"])
            w.write_row(["c", "d"])

        df = opensheet_core.read_xlsx_df(tmp_xlsx, header=False)
        assert list(df.columns) == [0, 1]
        assert len(df) == 2
        assert df.iloc[0][0] == "a"

    def test_read_specific_sheet_by_name(self, tmp_xlsx):
        """Read a specific sheet by name."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("First")
            w.write_row(["x"])
            w.add_sheet("Second")
            w.write_row(["Col"])
            w.write_row([42])

        df = opensheet_core.read_xlsx_df(tmp_xlsx, sheet_name="Second")
        assert list(df.columns) == ["Col"]
        assert df.iloc[0]["Col"] == 42

    def test_read_specific_sheet_by_index(self, tmp_xlsx):
        """Read a specific sheet by 0-based index."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("First")
            w.write_row(["a"])
            w.add_sheet("Second")
            w.write_row(["Col"])
            w.write_row([99])

        df = opensheet_core.read_xlsx_df(tmp_xlsx, sheet_index=1)
        assert list(df.columns) == ["Col"]
        assert df.iloc[0]["Col"] == 99

    def test_read_empty_sheet(self, tmp_xlsx):
        """Reading an empty sheet returns an empty DataFrame."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Empty")

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert len(df) == 0

    def test_read_with_dates(self, tmp_xlsx):
        """Dates come through as date objects in the DataFrame."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Dates")
            w.write_row(["Event", "Date"])
            w.write_row(["Launch", datetime.date(2025, 3, 15)])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert df.iloc[0]["Event"] == "Launch"
        assert df.iloc[0]["Date"] == datetime.date(2025, 3, 15)

    def test_read_with_datetimes(self, tmp_xlsx):
        """Datetimes come through as datetime objects in the DataFrame."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Times")
            w.write_row(["Event", "Timestamp"])
            w.write_row(["Deploy", datetime.datetime(2025, 3, 15, 14, 30, 0)])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert df.iloc[0]["Timestamp"] == datetime.datetime(2025, 3, 15, 14, 30, 0)

    def test_read_with_formulas(self, tmp_xlsx):
        """Formulas are unwrapped to their cached values."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Calc")
            w.write_row(["A", "B", "Sum"])
            w.write_row([10, 20, opensheet_core.Formula("SUM(A2:B2)", cached_value=30)])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert df.iloc[0]["Sum"] == 30

    def test_read_with_formatted_cells(self, tmp_xlsx):
        """FormattedCells are unwrapped to their plain values."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Money")
            w.write_row(["Item", "Price"])
            w.write_row(["Widget", opensheet_core.FormattedCell(19.99, "$#,##0.00")])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert abs(df.iloc[0]["Price"] - 19.99) < 0.001

    def test_read_with_none_cells(self, tmp_xlsx):
        """None cells become NaN in the DataFrame."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Sparse")
            w.write_row(["A", "B"])
            w.write_row([1, None])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert df.iloc[0]["A"] == 1
        assert pd.isna(df.iloc[0]["B"])

    def test_read_header_only(self, tmp_xlsx):
        """A sheet with just a header row returns empty DataFrame with columns."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Headers")
            w.write_row(["X", "Y", "Z"])

        df = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(df.columns) == ["X", "Y", "Z"]
        assert len(df) == 0


class TestToXlsx:
    def test_basic_write(self, tmp_xlsx):
        """Write a simple DataFrame and read it back."""
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [30, 25]})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(result.columns) == ["Name", "Age"]
        assert len(result) == 2
        assert result.iloc[0]["Name"] == "Alice"
        assert result.iloc[0]["Age"] == 30

    def test_write_no_header(self, tmp_xlsx):
        """Write without header row."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        opensheet_core.to_xlsx(df, tmp_xlsx, header=False)

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert len(rows) == 2
        assert rows[0] == [1, 3]
        assert rows[1] == [2, 4]

    def test_write_with_index(self, tmp_xlsx):
        """Write with index column."""
        df = pd.DataFrame({"Val": [10, 20]}, index=["a", "b"])
        df.index.name = "Key"
        opensheet_core.to_xlsx(df, tmp_xlsx, index=True)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(result.columns) == ["Key", "Val"]
        assert result.iloc[0]["Key"] == "a"
        assert result.iloc[0]["Val"] == 10

    def test_write_custom_sheet_name(self, tmp_xlsx):
        """Write with a custom sheet name."""
        df = pd.DataFrame({"X": [1]})
        opensheet_core.to_xlsx(df, tmp_xlsx, sheet_name="Custom")

        names = opensheet_core.sheet_names(tmp_xlsx)
        assert names == ["Custom"]

    def test_write_int_types(self, tmp_xlsx):
        """numpy integer types are converted properly."""
        df = pd.DataFrame({"A": np.array([1, 2, 3], dtype=np.int64)})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["A"] == 1

    def test_write_float_types(self, tmp_xlsx):
        """numpy float types are converted properly."""
        df = pd.DataFrame({"A": np.array([1.5, 2.7], dtype=np.float64)})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert abs(result.iloc[0]["A"] - 1.5) < 0.001

    def test_write_bool_types(self, tmp_xlsx):
        """Boolean columns round-trip correctly."""
        df = pd.DataFrame({"Flag": [True, False, True]})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["Flag"] == True
        assert result.iloc[1]["Flag"] == False

    def test_write_datetime_column(self, tmp_xlsx):
        """datetime64 columns are written as Excel dates."""
        df = pd.DataFrame({
            "Timestamp": pd.to_datetime(["2025-03-15", "2025-06-01"]),
        })
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["Timestamp"] == datetime.date(2025, 3, 15)
        assert result.iloc[1]["Timestamp"] == datetime.date(2025, 6, 1)

    def test_write_datetime_with_time(self, tmp_xlsx):
        """datetime64 with time component preserves time."""
        df = pd.DataFrame({
            "Timestamp": pd.to_datetime(["2025-03-15 14:30:00"]),
        })
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["Timestamp"] == datetime.datetime(2025, 3, 15, 14, 30, 0)

    def test_write_nan_values(self, tmp_xlsx):
        """NaN values are written as empty cells."""
        df = pd.DataFrame({"A": [1.0, np.nan, 3.0]})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        rows = opensheet_core.read_sheet(tmp_xlsx)
        # Header + 3 data rows
        assert rows[1] == [1]  # row with 1.0
        # NaN row might be empty or have None
        assert len(rows[2]) == 0 or rows[2] == [None]
        assert rows[3] == [3]

    def test_write_mixed_types(self, tmp_xlsx):
        """DataFrame with mixed column types."""
        df = pd.DataFrame({
            "Name": ["Alice", "Bob"],
            "Age": [30, 25],
            "Score": [95.5, 88.0],
            "Active": [True, False],
        })
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["Name"] == "Alice"
        assert result.iloc[0]["Age"] == 30
        assert abs(result.iloc[0]["Score"] - 95.5) < 0.001
        assert result.iloc[0]["Active"] == True

    def test_write_empty_dataframe(self, tmp_xlsx):
        """Writing an empty DataFrame produces a file with just headers."""
        df = pd.DataFrame({"A": pd.Series(dtype="int64"), "B": pd.Series(dtype="str")})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(result.columns) == ["A", "B"]
        assert len(result) == 0

    def test_roundtrip_large_dataframe(self, tmp_xlsx):
        """Roundtrip a larger DataFrame."""
        n = 1000
        df = pd.DataFrame({
            "id": range(n),
            "value": np.random.randn(n),
            "label": [f"item_{i}" for i in range(n)],
        })
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert len(result) == n
        assert list(result.columns) == ["id", "value", "label"]

    def test_write_unnamed_index(self, tmp_xlsx):
        """Index without a name gets 'index' as column header."""
        df = pd.DataFrame({"Val": [10, 20]}, index=["a", "b"])
        opensheet_core.to_xlsx(df, tmp_xlsx, index=True)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(result.columns) == ["index", "Val"]

    def test_write_none_in_object_column(self, tmp_xlsx):
        """None values in object columns are handled."""
        df = pd.DataFrame({"A": ["hello", None, "world"]})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert rows[0] == ["A"]
        assert rows[1] == ["hello"]
        # None row
        assert len(rows[2]) == 0 or rows[2] == [None]
        assert rows[3] == ["world"]

    def test_write_multiindex(self, tmp_xlsx):
        """MultiIndex (tuple) index values are written correctly."""
        arrays = [["a", "a", "b"], [1, 2, 1]]
        idx = pd.MultiIndex.from_arrays(arrays, names=["letter", "number"])
        df = pd.DataFrame({"Val": [10, 20, 30]}, index=idx)
        opensheet_core.to_xlsx(df, tmp_xlsx, index=True)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert list(result.columns) == ["letter", "number", "Val"]
        assert result.iloc[0]["letter"] == "a"
        assert result.iloc[0]["number"] == 1
        assert result.iloc[0]["Val"] == 10

    def test_write_python_date_passthrough(self, tmp_xlsx):
        """Python datetime.date values pass through directly."""
        df = pd.DataFrame({"D": [datetime.date(2025, 1, 1)]})
        # Force object dtype so pandas doesn't convert to Timestamp
        df["D"] = df["D"].astype(object)
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["D"] == datetime.date(2025, 1, 1)

    def test_write_python_datetime_passthrough(self, tmp_xlsx):
        """Python datetime.datetime values pass through directly."""
        df = pd.DataFrame({"T": [datetime.datetime(2025, 1, 1, 10, 30)]})
        df["T"] = df["T"].astype(object)
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["T"] == datetime.datetime(2025, 1, 1, 10, 30)

    def test_write_unsupported_type_stringified(self, tmp_xlsx):
        """Unsupported types are stringified via str()."""
        df = pd.DataFrame({"A": [complex(1, 2)]})
        df["A"] = df["A"].astype(object)
        opensheet_core.to_xlsx(df, tmp_xlsx)

        result = opensheet_core.read_xlsx_df(tmp_xlsx)
        assert result.iloc[0]["A"] == "(1+2j)"

    def test_read_no_header_ragged_rows(self, tmp_xlsx):
        """No-header mode pads ragged rows correctly."""
        with opensheet_core.XlsxWriter(tmp_xlsx) as w:
            w.add_sheet("Data")
            w.write_row(["a", "b", "c"])
            w.write_row(["d"])

        df = opensheet_core.read_xlsx_df(tmp_xlsx, header=False)
        assert len(df) == 2
        assert len(df.columns) == 3
        assert pd.isna(df.iloc[1][1])

    def test_write_nat_values(self, tmp_xlsx):
        """pd.NaT values in datetime columns are written as empty cells."""
        df = pd.DataFrame({"T": pd.to_datetime(["2025-01-01", None])})
        opensheet_core.to_xlsx(df, tmp_xlsx)

        rows = opensheet_core.read_sheet(tmp_xlsx)
        assert rows[0] == ["T"]
        assert rows[1] == [datetime.date(2025, 1, 1)]
        assert len(rows[2]) == 0 or rows[2] == [None]


class TestImportError:
    def test_read_xlsx_df_without_pandas(self, monkeypatch):
        """read_xlsx_df raises ImportError when pandas is not installed."""
        import builtins
        real_import = builtins.__import__

        def mock_import(name, *args, **kwargs):
            if name == "pandas":
                raise ImportError("No module named 'pandas'")
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, "__import__", mock_import)
        with pytest.raises(ImportError, match="pandas is required"):
            from opensheet_core.pandas import _check_pandas
            _check_pandas()
