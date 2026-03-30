# Benchmarking Methodology

This document explains how we benchmark OpenSheet Core against openpyxl, what we measure, and the design decisions behind our benchmark suite.

## Quick start

```bash
# Full benchmark (100k rows, 5 interleaved runs)
python benchmarks/benchmark.py

# Custom configuration
python benchmarks/benchmark.py --rows 50000 --cols 20 --runs 7

# Quick smoke test (1k rows, 1 run)
python benchmarks/benchmark.py --quick

# Specialized benchmarks (multiple dataset sizes)
python benchmarks/bench_read.py
python benchmarks/bench_write.py
```

## What we measure

### Timing

- **Clock**: `time.perf_counter()` — high-resolution monotonic clock
- **Reported value**: Minimum of N runs (least noisy, represents best achievable performance)
- **Also shown**: Mean and standard deviation across all runs, so readers can assess variance

The minimum is standard practice in microbenchmarking (used by Google Benchmark, Criterion.rs, etc.) because it filters out interference from other processes, GC pauses, and OS scheduling. The mean/stddev lets you verify the results are stable.

### Memory

- **Metric**: Current RSS (Resident Set Size) delta — memory actually in use before and after the workload
- **Platform APIs**:
  - **macOS**: `proc_pidinfo()` via `libproc` (returns `pti_resident_size`, the actual current RSS)
  - **Linux**: `/proc/self/statm` (page count * page size)
  - **Fallback**: `resource.getrusage(RUSAGE_SELF).ru_maxrss` (high-water mark, less accurate)
- **Reported value**: Median of N runs (robust against outliers)
- **Also shown**: Mean and standard deviation

#### Why not `ru_maxrss`?

Our earlier benchmarks used `resource.getrusage(RUSAGE_SELF).ru_maxrss`, which reports the **maximum** RSS the process has ever reached (a high-water mark). This has a critical flaw: it never decreases within a process lifetime. If importing a native library (like a Rust `.so`) temporarily spikes RSS during initialization, that spike becomes the baseline, and the workload's actual memory usage can appear artificially low — even zero — if it stays below that peak.

By measuring **current** RSS with platform-specific APIs, we get a true before/after picture of memory actually consumed by the workload.

## How we avoid bias

### Interleaved runs

Instead of running all opensheet measurements first and then all openpyxl measurements (`[A,A,A,B,B,B]`), we interleave them (`[A,B,A,B,A,B]`). This ensures both libraries experience similar system conditions:

- Thermal state (CPU may throttle after sustained load)
- Background process activity
- OS memory pressure and page cache state

### Subprocess isolation

Each individual measurement runs in a **fresh Python subprocess**. This provides:

- **Clean memory state**: No residual allocations from previous runs
- **Independent RSS**: Each process starts from zero, avoiding high-water-mark accumulation
- **No JIT/cache sharing**: Each run cold-starts the library

### Fair comparison modes

- **openpyxl write**: Uses `write_only=True` (streaming mode, openpyxl's fastest path)
- **openpyxl read**: Uses `read_only=True, data_only=True` (streaming mode with raw values)
- **Read benchmark**: Both libraries read the same openpyxl-generated file, avoiding format-specific advantages
- **Data generation**: Identical mixed-type rows (strings, integers, floats, booleans) for both

### Warm-up

Before measurement begins, both libraries perform a small warm-up run. This populates the OS page cache and ensures file I/O paths are primed. Warm-up results are discarded.

## Benchmark scripts

### `benchmark.py` — Main comparison

The primary benchmark. Compares read and write performance on a single dataset size.

| Parameter | Default | Description |
|-----------|---------|-------------|
| `--rows` | 100,000 | Number of data rows |
| `--cols` | 10 | Number of columns |
| `--runs` | 5 | Measurement runs per library |
| `--quick` | — | Quick mode: 1k rows, 1 run |

Output includes:
- Min time and mean +/- stddev for each library
- RSS delta (median) and memory +/- stddev
- Speedup ratio and memory comparison
- File sizes produced by each writer

### `bench_read.py` — Multi-size read comparison

Runs read benchmarks across multiple dataset sizes to show how performance scales:

| Config | Cells |
|--------|-------|
| 1,000 x 10 | 10K |
| 10,000 x 10 | 100K |
| 50,000 x 10 | 500K |
| 100,000 x 10 | 1M |
| 10,000 x 50 | 500K |

### `bench_write.py` — Multi-size write comparison

Same dataset configurations as the read benchmark, measuring write performance and output file sizes.

### `bench_utils.py` — Shared infrastructure

Core measurement functions used by all benchmark scripts:

- `BenchResult` — dataclass with min/mean/std for time and median/mean/std for memory
- `bench(func, *args, runs=5)` — measure a single function
- `bench_pair(func_a, args_a, func_b, args_b, runs=5)` — interleaved measurement of two functions
- `generate_row(r, cols)` — deterministic mixed-type row generation
- Platform-specific current RSS measurement in subprocess

## Interpreting the results

### What the numbers mean

| Metric | What it tells you |
|--------|-------------------|
| **Min time** | Best-case throughput, free from system noise |
| **Mean +/- stddev** | Typical performance and how much it varies |
| **RSS delta** | Memory consumed by the workload itself (excluding Python/library startup) |
| **File size** | Output file size (XLSX compression efficiency) |

### Known tradeoffs

- **Read memory**: OpenSheet Core materializes all rows into Python lists via `read_sheet()`. Despite this, it uses ~2.5x less memory than openpyxl thanks to deferred shared-string resolution (strings are stored as indices during Rust parsing and only converted to Python objects at the boundary via pre-interned lookup, avoiding duplicate allocations). A streaming iterator API is planned to bring constant-memory reads.
- **Write speed**: The speedup is more modest for writes because both libraries stream data — the bottleneck shifts toward Python-side row generation and data serialization.
- **File size**: OpenSheet Core files may be larger due to different XML formatting and compression settings. File content is equivalent.

### Memory optimization details

OpenSheet Core's read path uses several techniques to minimize memory:

1. **Deferred shared-string resolution**: During XML parsing, shared strings are stored as integer indices (`SharedString(idx)`) rather than cloned `String` values. This avoids O(N) string allocations during parsing.
2. **Pre-interned Python strings**: The shared string table is converted to Python objects once. When cells reference the same string, they reuse the existing Python object via `clone_ref()` instead of creating a new one.
3. **Convert-and-drop**: The Rust row data is consumed (taken by value) during Python conversion. As each row is converted, the Rust memory is freed immediately rather than holding both representations simultaneously.
4. **Single-sheet parsing**: `read_sheet()` only parses the requested worksheet, skipping all other sheets in the workbook.

### What could affect your results

- **Hardware**: SSD vs HDD affects write benchmarks significantly
- **Other processes**: Background activity adds noise (interleaving helps but doesn't eliminate this)
- **Python version**: CPython 3.11+ has measurably faster bytecode execution
- **Dataset shape**: Wide datasets (many columns) vs tall datasets (many rows) stress different code paths
- **Data types**: All-string workloads behave differently from mixed-type or all-numeric workloads

## Contributing to benchmarks

We welcome benchmark improvements. Some areas where we'd appreciate help:

- **Real-world datasets**: Benchmarks with actual spreadsheet patterns (sparse data, many sheets, formulas)
- **Additional libraries**: Comparisons against xlsxwriter, pylightxl, or other alternatives
- **Platform coverage**: Benchmark results from different hardware and OS configurations
- **Scaling analysis**: How performance changes from 1K to 10M rows

When submitting benchmark results, please include:
- Hardware (CPU model, RAM, disk type)
- OS and Python version
- OpenSheet Core and openpyxl versions
- Full benchmark output
