# Contributing to OpenSheet Core

Thanks for your interest in contributing! This guide will help you get set up and make your first contribution.

## Development setup

### Prerequisites

- **Python 3.9+**
- **Rust toolchain** — install via [rustup](https://rustup.rs/)
- **maturin** — `pip install maturin`

### Getting started

```bash
# Clone the repo
git clone https://github.com/0xNadr/opensheet-core
cd opensheet-core

# Create a virtual environment
python -m venv .venv
source .venv/bin/activate  # or .venv\Scripts\activate on Windows

# Install dev dependencies
pip install maturin pytest openpyxl

# Build the Rust extension in development mode
maturin develop --release

# Run the tests
pytest tests/ -v
```

### Running Rust tests

```bash
cargo test
```

### Running benchmarks

```bash
# Full benchmark (100k rows, 5 interleaved runs)
python benchmarks/benchmark.py

# Quick smoke test
python benchmarks/benchmark.py --quick

# Multi-size read/write benchmarks
python benchmarks/bench_read.py
python benchmarks/bench_write.py
```

See [docs/benchmarking.md](docs/benchmarking.md) for methodology details.

## Making changes

### Code style

- **Rust**: Run `cargo fmt` before committing. CI enforces this.
- **Rust linting**: Run `cargo clippy` and address any warnings.
- **Python**: Keep the Python layer thin — most logic lives in Rust.

### Testing

- Add tests for any new functionality in `tests/`
- Run the full test suite before opening a PR: `pytest tests/ -v`
- If you're adding a new reader feature, include a sample `.xlsx` file in `tests/fixtures/`

### Commit messages

Write clear, concise commit messages. Use the imperative mood:

- "Add freeze panes support" (not "Added" or "Adds")
- "Fix date parsing for 1900 leap year bug"

## Opening a pull request

1. Fork the repo and create a branch from `main`
2. Make your changes with tests
3. Ensure CI passes (`cargo fmt`, `cargo clippy`, `cargo test`, `pytest`)
4. Open a PR with a clear description of what and why

## What to work on

Check the [issues](https://github.com/0xNadr/opensheet-core/issues) — items labeled `good first issue` are great starting points. If you want to tackle something larger, open an issue first to discuss the approach.

## Reporting bugs

When reporting a bug, include:

- Python version and OS
- OpenSheet Core version (`python -c "import opensheet_core; print(opensheet_core.__version__)"`)
- Minimal code to reproduce the issue
- If possible, attach the `.xlsx` file that triggers the bug

## Questions?

Open a [discussion](https://github.com/0xNadr/opensheet-core/discussions) or an issue — happy to help.
