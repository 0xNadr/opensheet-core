"""Benchmark: opensheet_core vs openpyxl for writing XLSX files."""

import os
import sys
import tempfile

import openpyxl
import opensheet_core

from bench_utils import bench, format_bytes, format_time, generate_row


def do_openpyxl_write(path, rows, cols):
    """Write with openpyxl."""
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Benchmark")

    ws.append([f"col_{i}" for i in range(cols)])

    for r in range(rows):
        ws.append(generate_row(r, cols))

    wb.save(path)


def do_opensheet_write(path, rows, cols):
    """Write with opensheet_core."""
    with opensheet_core.XlsxWriter(path) as w:
        w.add_sheet("Benchmark")

        w.write_row([f"col_{i}" for i in range(cols)])

        for r in range(rows):
            w.write_row(generate_row(r, cols))


def format_speed_relative(ratio):
    if ratio == float("inf"):
        return "inf faster"
    if ratio >= 1:
        return f"{ratio:.1f}x faster"
    return f"{1 / ratio:.1f}x slower"


def format_memory_relative(os_mem, op_mem):
    if os_mem == 0 and op_mem == 0:
        return "no measurable RSS delta"
    if os_mem == 0:
        return "opensheet ~0 RSS delta"
    ratio = op_mem / os_mem
    if ratio >= 1:
        return f"{ratio:.1f}x less RSS delta"
    return f"{1 / ratio:.1f}x more RSS delta"


def run_benchmark(rows, cols, runs=3):
    print(f"\n{'='*60}")
    print(f"Benchmark: {rows:,} rows x {cols} cols ({rows * cols:,} cells)")
    print(f"{'='*60}")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        os_path = f.name
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        op_path = f.name

    try:
        # Warm up
        do_opensheet_write(os_path, min(rows, 100), cols)
        do_openpyxl_write(op_path, min(rows, 100), cols)

        # Benchmark
        os_time, os_mem = bench(do_opensheet_write, os_path, rows, cols, runs=runs)
        op_time, op_mem = bench(do_openpyxl_write, op_path, rows, cols, runs=runs)

        os_size = os.path.getsize(os_path)
        op_size = os.path.getsize(op_path)

        speedup = op_time / os_time if os_time > 0 else float("inf")
        speed_text = format_speed_relative(speedup)
        mem_text = format_memory_relative(os_mem, op_mem)

        print(f"  {'Library':<22} {'Time (min)':<15} {'Peak RSS Δ':<15} {'File Size':<15}")
        print(f"  {'-'*67}")
        print(f"  {'opensheet_core':<22} {format_time(os_time):<15} {format_bytes(os_mem):<15} {format_bytes(os_size):<15}")
        print(f"  {'openpyxl (write_only)':<22} {format_time(op_time):<15} {format_bytes(op_mem):<15} {format_bytes(op_size):<15}")
        print()
        print(f"  Speed:  opensheet_core is {speed_text}")
        print(f"  Memory: opensheet_core uses {mem_text}")

        return {
            "rows": rows,
            "cols": cols,
            "opensheet_time": os_time,
            "openpyxl_time": op_time,
            "opensheet_mem": os_mem,
            "openpyxl_mem": op_mem,
            "speedup": speedup,
        }
    finally:
        os.unlink(os_path)
        os.unlink(op_path)


def main():
    print("OpenSheet Core vs openpyxl — Write Benchmark")
    print(f"opensheet_core {opensheet_core.__version__}")
    print(f"openpyxl {openpyxl.__version__}")
    print(f"Python {sys.version.split()[0]}")
    print(f"Memory: peak RSS delta over pre-workload baseline")

    configs = [
        (1_000, 10),
        (10_000, 10),
        (50_000, 10),
        (100_000, 10),
        (10_000, 50),
    ]

    results = []
    for rows, cols in configs:
        result = run_benchmark(rows, cols)
        results.append(result)

    print(f"\n{'='*60}")
    print("Summary")
    print(f"{'='*60}")
    print(f"  {'Config':<20} {'Speed':<16} {'Memory':<24}")
    print(f"  {'-'*56}")
    for r in results:
        config = f"{r['rows']:,} x {r['cols']}"
        print(
            f"  {config:<20} "
            f"{format_speed_relative(r['speedup']):<16} "
            f"{format_memory_relative(r['opensheet_mem'], r['openpyxl_mem']):<24}"
        )


if __name__ == "__main__":
    main()
