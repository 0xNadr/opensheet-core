# Known Edge Cases and Limitations

This document records edge cases discovered through fuzzing and corpus
testing, along with the reader's behavior in each case.

## Security Limits

The reader enforces hard limits to prevent denial-of-service:

| Limit | Value | Effect |
|-------|-------|--------|
| Max ZIP entry size | 256 MB (decompressed) | Returns `XlsxError` |
| Max shared strings | 2,000,000 | Returns `XlsxError` |
| Max rows per sheet | 1,048,576 (Excel limit) | Returns `XlsxError` |

## ZIP / Archive Edge Cases

| Scenario | Behavior |
|----------|----------|
| Truncated ZIP | Clean `ZipError` |
| Empty ZIP (no entries) | Error: missing `xl/workbook.xml` |
| Non-ZIP data (random bytes) | Clean `ZipError` |
| Missing `xl/workbook.xml` | Error: `InvalidStructure` |
| Missing `_rels/.rels` | Falls back gracefully |
| ZIP bomb (large compression ratio) | Caught by entry size limit |

## XML Edge Cases

| Scenario | Behavior |
|----------|----------|
| Malformed XML (unclosed tags) | Partial parse; may return truncated data |
| Missing `<sheetData>` element | Empty rows returned |
| Unknown XML elements | Ignored (forward-compatible) |
| XML with BOM | Handled by quick-xml |
| Very deeply nested XML | Parsed normally (no depth limit) |

## Cell & Data Type Edge Cases

| Scenario | Behavior |
|----------|----------|
| Shared string index out of bounds | Returns `"[invalid string index]"` |
| Empty cell element `<c/>` | Returns `Empty` |
| Cell with no `<v>` child | Returns `Empty` |
| Boolean cell with value not 0 or 1 | Treated as truthy/falsy |
| Inline string with rich text runs | Concatenated plain text |
| Formula with cached value | Returns `Formula { formula, cached_value }` |
| Error cell (`t="e"`) | Returns the error string (e.g. `"#REF!"`) |
| Number format code detection | Date serial numbers with date format codes → `Date` |

## Row & Structure Edge Cases

| Scenario | Behavior |
|----------|----------|
| Row number 0 | Treated as row 0 (no crash) |
| Sparse rows (gaps in numbering) | Gaps filled with empty rows |
| Row at max (1,048,576) | Accepted (at limit) |
| Duplicate row numbers | Last row wins |
| Columns beyond XFD (16,384) | Parsed if present |
| Merged cells with no data in merge area | Merge range recorded, cells empty |

## Sheet & Workbook Edge Cases

| Scenario | Behavior |
|----------|----------|
| Hidden/veryHidden sheets | Parsed with `state` field set |
| Sheet with no rows | Empty `rows` list |
| 100+ sheets in one workbook | All parsed |
| Sheet name with special characters | Preserved as-is |
| Defined names with `hidden="1"` | Included in output |
| Missing `xl/_rels/workbook.xml.rels` | Error: cannot resolve sheet paths |

## File Source Compatibility

The test corpus includes files structured to match output from:

- **Microsoft Excel** (standard OOXML)
- **LibreOffice Calc** (may use different XML namespaces)
- **Google Sheets** (export as .xlsx)
- **opensheet-core writer** (roundtrip testing)
- **Programmatically generated** (minimal valid XLSX)

## Fuzzing

Fuzz targets exercise all reader entry points:

- `fuzz_read_xlsx` — Full workbook parse
- `fuzz_read_single_sheet` — Single sheet extraction
- `fuzz_read_sheet_names` — Sheet name enumeration
- `fuzz_read_document_properties` — Document metadata parse

Run locally:
```bash
cargo +nightly fuzz run fuzz_read_xlsx -- -max_total_time=60
```

CI runs fuzzing weekly and on PRs touching `src/reader/` or `fuzz/`.
