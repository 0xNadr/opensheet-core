# OpenSheet Core

OpenSheet Core is an open source spreadsheet I/O core for Python focused on low-memory reading, streaming XLSX writing, and easy adoption through pip-installable Python bindings.

The goal is to build a reusable spreadsheet infrastructure layer for Python and adjacent data tooling that reduces the usual tradeoffs between compatibility, speed, memory usage, and installation friction.

## Why this project exists

Spreadsheet files remain a common interchange format across data analysis, public sector workflows, finance, research, and business automation. In practice, existing Python tools often force users to choose between:

- broad feature support
- good performance on large files
- low memory usage
- simple installation and deployment

OpenSheet Core aims to improve that situation with a native core and a stable Python-facing API.

## Project goals

The first phase of the project is focused on practical, high-value workflows rather than full spreadsheet feature parity.

### Core goals

- low-memory reading of common spreadsheet formats
- streaming XLSX writing for large files
- pip-installable Python package with prebuilt wheels
- predictable typed cell extraction
- public benchmarks and compatibility reports
- regression testing and parser hardening
- clear documentation and examples

### Non-goals for the first phase

OpenSheet Core is **not** initially trying to be a full replacement for every existing spreadsheet library feature. The first phase does not promise:

- full Excel feature parity
- complete macro or VBA support
- perfect round-tripping of every workbook edge case
- advanced spreadsheet editing features beyond core read/write workflows

## Initial scope

Planned initial scope:

- fast reading of common spreadsheet formats
- streaming XLSX writer for large files
- Python bindings with prebuilt wheels
- compatibility corpus, benchmarks, regression tests, and fuzzing

## Planned feature areas

The exact implementation may evolve, but the initial roadmap is centered on the following areas.

### Reading

- workbook and worksheet discovery
- row-wise worksheet iteration
- typed cell extraction
- formula values and raw formulas where available
- merged-cell metadata
- defined names and workbook metadata
- support for common spreadsheet input workflows

### Writing

- XLSX output
- streaming row-by-row writing
- shared strings
- formulas
- basic styling support for common cases
- memory-efficient write mode for large datasets

### Python integration

- clean Python API
- normal `pip install` experience
- prebuilt wheels for major platforms
- optional integration helpers for adjacent Python data tools

### Quality and reliability

- compatibility corpus built from real-world examples
- regression testing
- differential testing where useful
- fuzzing and parser hardening
- clear support matrix and documented limitations

## Design principles

OpenSheet Core is intended to follow these principles:

- **practical scope first**  
  Ship a useful core before chasing long-tail features.

- **low memory by design**  
  Large spreadsheets should not require excessive memory overhead.

- **easy adoption**  
  End users should not need a Rust toolchain or custom build setup.

- **open infrastructure**  
  This project should be useful as a building block for other tools, not only as a standalone package.

- **clear compatibility boundaries**  
  Supported and unsupported features should be explicit.

- **public development**  
  Roadmap, issues, and progress should be visible and open to feedback.

## Proposed architecture

The working direction is:

- a native core for spreadsheet parsing and writing
- Python bindings for end-user adoption
- packaging that supports prebuilt wheels on major platforms
- a test and benchmark suite to validate behavior and performance over time

This may change as the prototype develops, but the main intention is to keep the Python API easy to use while moving performance-critical work into a lower-level core.

## Roadmap

### Phase 1: project foundation

- define scope
- publish roadmap
- set up repository structure
- choose licensing
- prepare benchmark plan
- collect representative spreadsheet samples

### Phase 2: reader prototype

- implement basic workbook reading
- support worksheet iteration
- expose typed cell values
- add Python bindings
- validate against initial test corpus

### Phase 3: writer prototype

- implement streaming XLSX writer
- support large-file row writing
- add formulas and common metadata handling
- expand tests and examples

### Phase 4: hardening and packaging

- prebuilt wheels for major platforms
- regression tests
- fuzzing
- benchmark publication
- documentation and migration guidance

## Status

Early project planning / prototype stage.

This repository is currently being prepared for initial implementation. The current focus is:

- scope definition
- roadmap refinement
- packaging design
- test corpus planning
- milestone planning for an initial public prototype

## Installation

Not yet available.

The long-term goal is a normal installation flow such as:

```bash
pip install opensheet-core
````

with prebuilt wheels for major platforms so that ordinary users do not need to compile from source.

## Development status

The project is not production-ready yet.

Until the first prototype is published, everything should be treated as exploratory and subject to change, including:

* API design
* supported formats
* internal architecture
* milestone ordering

## Contributing

Contributions, feedback, and ecosystem input are welcome.

Early useful contributions include:

* issue reports about real spreadsheet pain points
* representative sample files for testing
* compatibility edge cases
* packaging suggestions
* benchmark scenarios
* review of scope and roadmap

Once the implementation begins, contribution guidelines and development setup instructions will be added.

## Intended users and ecosystem

OpenSheet Core is intended for:

* Python developers
* data engineers
* analysts
* researchers
* automation developers
* maintainers of higher-level data and spreadsheet tooling
* public sector and civic tech workflows
* teams dealing with large workbook ingestion or export

The project is meant to support both direct end-user usage and reuse by other open-source tools.

## Funding and sustainability

This project is being positioned as open digital infrastructure. The intention is to build it in public, keep it openly licensed, and make it sustainable through transparent development and community engagement.

## License

MIT

See the [LICENSE](LICENSE) file for details.
