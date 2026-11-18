"""Pakuje wyłącznie jawnie wybrane pliki, bez baz użytkownika i katalogu .git."""
import hashlib
import argparse
import json
import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'dist' / 'Floteka'


def package(output=None):
    global OUT
    if output is not None:
        OUT = Path(output).resolve()
    reports = ['excel-test.json', 'excel-regressions.json', 'excel-deadlines.json', 'excel-optional.json', 'excel-export.json', 'exe-test.json']
    for name in reports:
        report = json.loads((ROOT / 'docs' / name).read_text(encoding='utf-8'))
        if not report.get('ok'):
            raise RuntimeError(f'Test nie przeszedł: {name}')
        if name == 'excel-export.json' and not report.get('frozen_export_checked'):
            raise RuntimeError('Sprawdź eksport gotowego EXE w Excelu: tools/test_export_excel.py --exe ...')
    shutil.copyfile(ROOT / 'README.md', OUT / 'INSTRUKCJA.md')
    shutil.copyfile(ROOT / 'docs' / 'verification.md', OUT / 'WERYFIKACJA.md')
    (OUT / 'Uruchom-DEMO.bat').write_bytes(
        b'@echo off\r\ncd /d "%~dp0"\r\nstart "" "%~dp0Floteka.exe" --db "%~dp0Floteka-DEMO.sqlite3"\r\n')
    folders = ['fleet', 'excel', 'tests', '.github']
    files = [ROOT / name for name in ['main.py', 'README.md', 'requirements.txt', 'requirements-build.txt', 'Uruchom.bat', '.gitignore', 'docs/verification.md']]
    files += [ROOT / 'docs' / name for name in reports]
    files += [ROOT / 'tools' / name for name in ['build_excel.py', 'build_exe.ps1', 'release_version.py', 'excel_paths.py', 'build_release.ps1', 'test_excel.py', 'test_excel_regressions.py', 'test_deadlines_excel.py', 'test_optional_excel.py', 'test_export_excel.py', 'make_demo.py', 'package_release.py', 'check_excel_filters_speed.py', 'repair_excel_copy.py', 'capture_preview.py']]
    for folder in folders:
        files += [p for p in (ROOT / folder).rglob('*') if p.is_file() and '__pycache__' not in p.parts]
    with zipfile.ZipFile(OUT / 'Zrodla.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(set(files)):
            archive.write(path, path.relative_to(ROOT).as_posix())
    names = ['Floteka.exe', 'Floteka.xlsm', 'Floteka-DEMO.xlsm', 'Floteka-DEMO.sqlite3',
             'Uruchom-DEMO.bat', 'INSTRUKCJA.md', 'WERYFIKACJA.md', 'Zrodla.zip']
    with zipfile.ZipFile(OUT / 'Floteka-pakiet.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.write(OUT / name, 'Floteka/' + name)
    names += ['Floteka-pakiet.zip']
    sums = [hashlib.sha256((OUT / name).read_bytes()).hexdigest() + '  ' + name for name in names]
    (OUT / 'SUMY-SHA256.txt').write_text('\n'.join(sums) + '\n', encoding='utf-8')
    print('Pakiet:', OUT / 'Floteka-pakiet.zip')
    print('Pliki:', len(names))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path)
    package(parser.parse_args().output)
