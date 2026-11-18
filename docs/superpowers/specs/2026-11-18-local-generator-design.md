# Local EXE and XLSM generator

The user approved one `Buduj.bat` entry point producing Floteka.exe and both
XLSM workbooks with their existing macros. Building must not require Microsoft
Excel, Microsoft 365, or a self-hosted runner. Using the XLSM macros still
requires desktop Excel at the recipient. The existing application remains
independent of Office.

Use verified, empty and synthetic DEMO XLSM templates. Copy ZIP members without
changing VBA, controls, relationships or layout; update only version metadata,
the START version label and page footers. Track template and relevant source
hashes in a manifest and reject stale or modified templates before writing.
Refresh templates through the existing COM builder on a development machine
with Excel; this is a maintenance tool, not part of normal local builds.
This preserves the current workbook rather than replacing its behavior with a
macro-free XLSX or relying on incomplete VBA compatibility in LibreOffice.

`Buduj.bat` calls a PowerShell orchestrator. It finds Python 3.13 with Tcl/Tk,
creates an isolated build environment, installs the declared dependencies,
generates XLSM, runs application/layout/version tests, builds and self-tests
the real EXE, and produces a ZIP plus SHA256 hashes in a fresh output folder.
Python and first-run internet access are documented prerequisites. Existing
output files and user data must never be overwritten. Local reports explicitly
distinguish structural template validation from live Excel macro execution.
Source ZIPs include the launcher, generator, templates and manifest.

Existing COM build and GitHub release behavior remain available. The earlier
optional-Excel workflow fix remains in the checkout. No push or publication is
part of this request.
