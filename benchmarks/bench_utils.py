"""Shared utilities for the OpenSheet Core benchmark suite."""

import gc
import json
import os
import subprocess
import sys
import statistics
import textwrap
import time
import tracemalloc


def format_bytes(n):
    """Format a byte count for human display."""
    if n < 1024:
        return f"{n} B"
    elif n < 1024 * 1024:
        return f"{n / 1024:.1f} KB"
    else:
        return f"{n / (1024 * 1024):.1f} MB"


def format_time(seconds):
    """Format seconds for human display."""
    if seconds < 1:
        return f"{seconds * 1000:.0f} ms"
    return f"{seconds:.2f} s"


def _measure_in_subprocess(func_module, func_name, args):
    """Run a benchmark function in a fresh subprocess to get clean RSS.

    Each run gets a fresh process, avoiding the ru_maxrss high-water-mark
    problem for repeated in-process measurements. Memory is reported as
    peak RSS above the pre-workload baseline in that subprocess.
    """
    script = textwrap.dedent("""\
        import gc
        import importlib
        import json
        import resource
        import sys
        import time

        mod = importlib.import_module(sys.argv[1])
        func = getattr(mod, sys.argv[2])
        args = json.loads(sys.argv[3])

        gc.collect()
        gc.collect()
        rss_before = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        if sys.platform == "linux":
            rss_before *= 1024
        t0 = time.perf_counter()
        func(*args)
        elapsed = time.perf_counter() - t0
        rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        if sys.platform == "linux":
            rss *= 1024
        rss_delta = max(0, rss - rss_before)
        print(json.dumps({"time": elapsed, "rss": rss_delta}))
    """)
    result = subprocess.run(
        [sys.executable, "-c", script, func_module, func_name, json.dumps(list(args))],
        capture_output=True, text=True,
        cwd=os.path.dirname(os.path.abspath(__file__)),
    )
    if result.returncode != 0:
        raise RuntimeError(f"Subprocess failed:\n{result.stderr}")
    lines = [line for line in result.stdout.splitlines() if line.strip()]
    if not lines:
        raise RuntimeError("Subprocess produced no stdout output.")
    try:
        data = json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Subprocess did not produce valid JSON.\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
        ) from exc
    return data["time"], data["rss"]


def measure_inprocess(func, *args):
    """In-process measurement using tracemalloc (Python allocations only).

    Useful as a fallback and for quick checks. Note: does not capture
    native/Rust allocations.
    """
    gc.collect()
    gc.collect()
    tracemalloc.start()
    t0 = time.perf_counter()
    func(*args)
    elapsed = time.perf_counter() - t0
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return elapsed, peak


def bench(func, *args, runs=3, subprocess_mode=True):
    """Run a benchmark multiple times, return (min_time, median_mem).

    Reports min time (least noisy) and median peak memory.

    When subprocess_mode=True (default), each run executes in a fresh
    subprocess so that RSS measurements are independent and accurate
    for native/Rust code. The function must be importable by name from
    its module.
    """
    times, mems = [], []

    if subprocess_mode:
        func_name = func.__name__
        func_module = func.__module__
        if func_module == "__main__":
            # Derive module name from the file
            import inspect
            source_file = inspect.getfile(func)
            func_module = os.path.splitext(os.path.basename(source_file))[0]
        for _ in range(runs):
            t, m = _measure_in_subprocess(func_module, func_name, args)
            times.append(t)
            mems.append(m)
    else:
        for _ in range(runs):
            t, m = measure_inprocess(func, *args)
            times.append(t)
            mems.append(m)

    return min(times), int(statistics.median(mems))


def generate_row(r, cols):
    """Generate a benchmark row with mixed types."""
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
