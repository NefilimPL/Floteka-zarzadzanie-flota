# Local generator implementation plan

> **For agentic workers:** Use superpowers:subagent-driven-development or superpowers:executing-plans to implement the tasks.

**Goal:** One Buduj.bat produces a verified EXE and existing macro workbooks without Office installed.
**Architecture:** A pure Python template generator and a Windows build orchestrator/package helper.
**Tech Stack:** Python 3.13, zipfile, XML DOM, PowerShell, PyInstaller, pytest, openpyxl.
**Spec:** docs/superpowers/specs/2026-11-18-local-generator-design.md

## Global constraints

- Preserve current VBA and layout, previous changes, and all user files.
- Fresh output directories; fail rather than overwrite.
- Normal generation must not import or launch COM/Excel.
- Python/Tcl/Tk required; dependencies installed only in build environment.
- Template/source mismatch fails before output; version and source provenance match EXE.
- Reports must disclose whether live macro execution occurred.

## Review focus

- ZIP XML edits preserve namespaces, buttons, VBA and valid workbook relationships.
- Changes to VBA/layout cannot silently reuse stale templates.
- Partial/failed builds and nonempty destinations never produce an apparent successful package.
- Paths with spaces and archive-based builds without Git work.
- Package contains no private databases, migrated workbooks, stale reports or build environment.

## Task 1: Macro-preserving template generator

Files: tools/generate_excel.py, tools/refresh_excel_templates.py,
excel/templates/*, tests/test_generate_excel.py, .gitignore.

- [x] Add failing tests for generation without COM, two tagged versions, unchanged VBA and drawings, clean/synthetic data, stale hashes and existing output preservation.
- [x] Implement generate(output, tag=None, *, root=ROOT) returning the output excel-build.json report.
- [x] Prepare and verify canonical templates with the existing Excel builder; save manifest source and file hashes.
- [x] Run structural tests and live Excel regression tests on generated copies if the development environment supports Excel.

## Task 2: One-file entry point and verified packaging

Files: Buduj.bat, tools/build_local.ps1, tools/package_local.py,
tests/test_local_package.py, README.md, docs/verification.md.

- [x] Add tests proving a valid package and rejection of missing/mismatched reports or existing output.
- [x] Implement package(output, *, root=ROOT): validate EXE and template build reports and packaged files; create only explicitly allowed assets, sources ZIP and hashes.
- [x] Implement orchestrator using the generator from Task 1, test suite, existing build_exe.ps1, make_demo.py and package helper. A new output folder is required; log errors and return nonzero.
- [x] Document Python prerequisite, one-click build, tags, output location and template-refresh limit.

## Final verification

- [x] Run full tests and YAML/PowerShell syntax checks.
- [x] Run the actual launcher end to end; inspect generated workbooks and run packaged EXE self-test.
- [x] Review diff and report evidence separately for Office-free generation and live macro tests.
