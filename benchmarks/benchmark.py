#!/usr/bin/env python3
"""
OpenSheet Core Benchmark Suite

Compare opensheet_core vs openpyxl for reading and writing XLSX files.

Usage:
    python benchmarks/benchmark.py              # default: 100k rows
    python benchmarks/benchmark.py --rows 50000 # custom row count
    python benchmarks/benchmark.py --quick      # fast smoke test (1k rows)
"""

import argparse
import os
import sys
import tempfile
import time
import tracemalloc

try:
    import openpyxl
except ImportError:
    print("openpyxl is required for benchmarking: pip install openpyxl")
    sys.exit(1)

import opensheet_core


COLS = 10


def format_bytes(n):
    if n < 1024:
        return f"{n} B"
    elif n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    else:
        return f"{n / (1024 * 1024):.1f} MB"


def format_time(seconds):
    if seconds < 1:
        return f"{seconds * 1000:.0f} ms"
    return f"{seconds:.2f} s"


def generate_row(r, cols):
    row = []
    for c in range(cols):
        match c % 4:
            case 0:
                row.append(f"text_{r}_{c}")
            case 1:
                row.append(r * cols + c)
            case 2:
                row.append((r * cols + c) * 0.123)
            case 3:
                row.append(r % 2 == 0)
    return row


# --- Write benchmarks ---

def write_opensheet(path, rows, cols):
    tracemalloc.start()
    t0 = time.perf_counter()
    with opensheet_core.XlsxWriter(path) as w:
        w.add_sheet("Benchmark")
        w.write_row([f"col_{i}" for i in range(cols)])
        for r in range(rows):
            w.write_row(generate_row(r, cols))
    elapsed = time.perf_counter() - t0
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed, peak


def write_openpyxl(path, rows, cols):
    tracemalloc.start()
    t0 = time.perf_counter()
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Benchmark")
    ws.append([f"col_{i}" for i in range(cols)])
    for r in range(rows):
        ws.append(generate_row(r, cols))
    wb.save(path)
    elapsed = time.perf_counter() - t0
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed, peak


# --- Read benchmarks ---

def read_opensheet(path):
    tracemalloc.start()
    t0 = time.perf_counter()
    rows = opensheet_core.read_sheet(path)
    _ = len(rows)
    elapsed = time.perf_counter() - t0
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed, peak


def read_openpyxl(path):
    tracemalloc.start()
    t0 = time.perf_counter()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        _ = list(row)
    wb.close()
    elapsed = time.perf_counter() - t0
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed, peak


# --- Runner ---

def bench(func, *args, runs=3):
    """Run a benchmark function multiple times, return (avg_time, avg_peak_mem)."""
    times, mems = [], []
    for _ in range(runs):
        t, m = func(*args)
        times.append(t)
        mems.append(m)
    return sum(times) / len(times), sum(mems) / len(mems)


def print_comparison(label, os_time, os_mem, op_time, op_mem):
    speedup = op_time / os_time if os_time > 0 else float("inf")
    mem_ratio = op_mem / os_mem if os_mem > 0 else float("inf")

    print(f"\n  {label}")
    print(f"  {'Library':<22} {'Time':<12} {'Peak Memory':<14}")
    print(f"  {'-'*48}")
    print(f"  {'opensheet_core':<22} {format_time(os_time):<12} {format_bytes(os_mem):<14}")
    print(f"  {'openpyxl':<22} {format_time(op_time):<12} {format_bytes(op_mem):<14}")
    print(f"  -> {speedup:.1f}x faster, {mem_ratio:.0f}x less memory")

    return speedup, mem_ratio


def main():
    parser = argparse.ArgumentParser(description="OpenSheet Core Benchmark Suite")
    parser.add_argument("--rows", type=int, default=100_000, help="Number of rows (default: 100000)")
    parser.add_argument("--cols", type=int, default=COLS, help="Number of columns (default: 10)")
    parser.add_argument("--runs", type=int, default=3, help="Runs per benchmark (default: 3)")
    parser.add_argument("--quick", action="store_true", help="Quick mode: 1000 rows, 1 run")
    args = parser.parse_args()

    if args.quick:
        args.rows = 1_000
        args.runs = 1

    rows, cols, runs = args.rows, args.cols, args.runs

    print("=" * 55)
    print("  OpenSheet Core Benchmark Suite")
    print("=" * 55)
    print(f"  opensheet_core  {opensheet_core.__version__}")
    print(f"  openpyxl        {openpyxl.__version__}")
    print(f"  Python          {sys.version.split()[0]}")
    print(f"  Dataset         {rows:,} rows x {cols} cols ({rows * cols:,} cells)")
    print(f"  Runs            {runs}")

    os_path = tempfile.mktemp(suffix=".xlsx")
    op_path = tempfile.mktemp(suffix=".xlsx")

    try:
        # Warm up
        write_opensheet(os_path, min(rows, 100), cols)
        write_openpyxl(op_path, min(rows, 100), cols)

        # Write benchmark
        os_wt, os_wm = bench(write_opensheet, os_path, rows, cols, runs=runs)
        op_wt, op_wm = bench(write_openpyxl, op_path, rows, cols, runs=runs)
        write_speed, write_mem = print_comparison("WRITE", os_wt, os_wm, op_wt, op_wm)

        os_size = os.path.getsize(os_path)
        op_size = os.path.getsize(op_path)
        print(f"  File sizes: opensheet {format_bytes(os_size)}, openpyxl {format_bytes(op_size)}")

        # Read benchmark (use the opensheet-written file)
        read_opensheet(os_path)  # warm up
        read_openpyxl(os_path)

        os_rt, os_rm = bench(read_opensheet, os_path, runs=runs)
        op_rt, op_rm = bench(read_openpyxl, os_path, runs=runs)
        read_speed, read_mem = print_comparison("READ", os_rt, os_rm, op_rt, op_rm)

        # Summary
        print(f"\n{'=' * 55}")
        print("  SUMMARY")
        print(f"{'=' * 55}")
        print(f"  {'Operation':<10} {'Speedup':<14} {'Memory Savings':<14}")
        print(f"  {'-'*38}")
        print(f"  {'Write':<10} {write_speed:.1f}x faster   {write_mem:.0f}x less")
        print(f"  {'Read':<10} {read_speed:.1f}x faster   {read_mem:.0f}x less")
        print()

    finally:
        for p in (os_path, op_path):
            if os.path.exists(p):
                os.unlink(p)


if __name__ == "__main__":
    main()
