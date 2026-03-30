"""Shared utilities for the OpenSheet Core benchmark suite."""

import dataclasses
import gc
import json
import os
import statistics
import subprocess
import sys
import textwrap
import time
import tracemalloc


@dataclasses.dataclass
class BenchResult:
    """Benchmark result with full statistics."""

    min_time: float
    mean_time: float
    std_time: float
    median_mem: int
    mean_mem: float
    std_mem: float


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

    Each run gets a fresh process, avoiding high-water-mark problems.
    Memory is reported as current RSS delta (not ru_maxrss) using
    platform-specific APIs for accurate measurement.
    """
    script = textwrap.dedent("""\
        import gc
        import importlib
        import json
        import os
        import sys
        import time

        def _get_rss():
            \"\"\"Get current RSS in bytes (not high-water mark).\"\"\"
            if sys.platform == "linux":
                with open("/proc/self/statm") as f:
                    return int(f.read().split()[1]) * os.sysconf("SC_PAGE_SIZE")
            elif sys.platform == "darwin":
                import ctypes
                import ctypes.util
                lib = ctypes.util.find_library("proc")
                if lib:
                    libproc = ctypes.CDLL(lib)
                    buf = ctypes.create_string_buffer(128)
                    # PROC_PIDTASKINFO = 4
                    ret = libproc.proc_pidinfo(os.getpid(), 4, 0, buf, 128)
                    if ret > 0:
                        # pti_resident_size is at offset 8 (after pti_virtual_size)
                        return int.from_bytes(buf[8:16], byteorder="little")
            # Fallback: ru_maxrss (high-water mark, less accurate)
            import resource
            rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
            if sys.platform == "linux":
                rss *= 1024
            return rss

        mod = importlib.import_module(sys.argv[1])
        func = getattr(mod, sys.argv[2])
        args = json.loads(sys.argv[3])

        gc.collect()
        gc.collect()
        rss_before = _get_rss()
        t0 = time.perf_counter()
        func(*args)
        elapsed = time.perf_counter() - t0
        rss_after = _get_rss()
        rss_delta = max(0, rss_after - rss_before)
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


def _resolve_func_info(func):
    """Resolve module name and function name for subprocess invocation."""
    func_name = func.__name__
    func_module = func.__module__
    if func_module == "__main__":
        import inspect
        source_file = inspect.getfile(func)
        func_module = os.path.splitext(os.path.basename(source_file))[0]
    return func_module, func_name


def _make_result(times, mems):
    """Build a BenchResult from collected samples."""
    mean_t = statistics.mean(times)
    std_t = statistics.stdev(times) if len(times) > 1 else 0.0
    mean_m = statistics.mean(mems)
    std_m = statistics.stdev(mems) if len(mems) > 1 else 0.0
    return BenchResult(
        min_time=min(times),
        mean_time=mean_t,
        std_time=std_t,
        median_mem=int(statistics.median(mems)),
        mean_mem=mean_m,
        std_mem=std_m,
    )


def bench(func, *args, runs=5, subprocess_mode=True):
    """Run a benchmark multiple times, return BenchResult.

    Reports min time (least noisy), mean +/- stddev, and median memory.

    When subprocess_mode=True (default), each run executes in a fresh
    subprocess so that RSS measurements are independent and accurate
    for native/Rust code. The function must be importable by name from
    its module.
    """
    times, mems = [], []

    if subprocess_mode:
        func_module, func_name = _resolve_func_info(func)
        for _ in range(runs):
            t, m = _measure_in_subprocess(func_module, func_name, args)
            times.append(t)
            mems.append(m)
    else:
        for _ in range(runs):
            t, m = measure_inprocess(func, *args)
            times.append(t)
            mems.append(m)

    return _make_result(times, mems)


def bench_pair(func_a, args_a, func_b, args_b, *, runs=5, subprocess_mode=True):
    """Run two benchmarks with interleaved runs to avoid ordering bias.

    Instead of [A,A,A,B,B,B], runs [A,B,A,B,A,B,...] so both libraries
    experience similar system conditions (thermal state, background load,
    memory pressure).
    """
    times_a, mems_a = [], []
    times_b, mems_b = [], []

    if subprocess_mode:
        info_a = _resolve_func_info(func_a)
        info_b = _resolve_func_info(func_b)
        for _ in range(runs):
            t, m = _measure_in_subprocess(info_a[0], info_a[1], args_a)
            times_a.append(t)
            mems_a.append(m)
            t, m = _measure_in_subprocess(info_b[0], info_b[1], args_b)
            times_b.append(t)
            mems_b.append(m)
    else:
        for _ in range(runs):
            t, m = measure_inprocess(func_a, *args_a)
            times_a.append(t)
            mems_a.append(m)
            t, m = measure_inprocess(func_b, *args_b)
            times_b.append(t)
            mems_b.append(m)

    return _make_result(times_a, mems_a), _make_result(times_b, mems_b)


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
