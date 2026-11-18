"""Package the verified outputs of the local, Office-free build."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile

from tools.release_version import parse_tag

ROOT = Path(__file__).resolve().parents[1]


WORKBOOKS = ['Floteka.xlsm', 'Floteka-DEMO.xlsm']
REPORTS = ['exe-test.json', 'build-info.json', 'excel-build.json']
SOURCE_FILES = ['main.py', 'README.md', 'requirements.txt', 'requirements-build.txt',
                'Buduj.bat', 'Uruchom.bat', '.gitignore', 'docs/verification.md',
                'tools/build_local.ps1', 'tools/package_local.py', 'tools/generate_excel.py',
                'tools/refresh_excel_templates.py', 'tools/build_exe.ps1',
                'tools/release_version.py', 'tools/make_demo.py',
                'excel/templates/Floteka.xlsm', 'excel/templates/Floteka-DEMO.xlsm',
                'excel/templates/manifest.json']


def sha256(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def checked_file(path, base):
    if path.is_symlink() or not path.resolve().is_relative_to(base.resolve()):
        raise ValueError(f'Unsafe package path: {path}')
    if not path.is_file():
        raise FileNotFoundError(path)
    return path


def read_report(output, name):
    report = json.loads(checked_file(output / name, output).read_text(encoding='utf-8-sig'))
    if not isinstance(report, dict):
        raise ValueError(f'Invalid report: {name}')
    return report


def validate(output):
    exe = read_report(output, 'exe-test.json')
    build = read_report(output, 'build-info.json')
    excel = read_report(output, 'excel-build.json')
    if (exe.get('ok') is not True or exe.get('frozen') is not True or
            exe.get('app') != 'Floteka' or build.get('app') != 'Floteka' or
            build.get('self_test') is not True):
        raise ValueError('EXE self-test does not confirm a packaged Floteka application.')
    version = build.get('version')
    if not isinstance(version, str) or not version or exe.get('version') != version:
        raise ValueError('EXE version mismatch.')
    if excel.get('ok') is not True or excel.get('engine') != 'template' or excel.get('macro_tests_run') is not False:
        raise ValueError('Expected a successful template report disclosing that macros were not executed.')
    for field in ['version', 'tag', 'commit']:
        if field not in build or field not in excel or build[field] != excel[field]:
            raise ValueError(f'EXE/XLSM provenance mismatch: {field}')
    tag = build['tag']
    if tag and parse_tag(tag).version != version:
        raise ValueError('Tag does not match the packaged version.')
    if not tag and re.fullmatch(r'\d+\.\d+\.\d+(?:\.\d+)?', version) is None:
        raise ValueError('Invalid default version.')
    commit = build['commit']
    if commit is not None and (not isinstance(commit, str) or re.fullmatch(r'[0-9a-f]{40,64}', commit) is None):
        raise ValueError('Invalid source commit.')
    if sha256(checked_file(output / 'Floteka.exe', output)) != build.get('sha256'):
        raise ValueError('EXE SHA256 mismatch.')
    if excel.get('files') != WORKBOOKS:
        raise ValueError('Template report must list both workbooks.')
    hashes = excel.get('hashes')
    if not isinstance(hashes, dict) or set(hashes) != set(WORKBOOKS):
        raise ValueError('Template report must hash both workbooks.')
    for name in WORKBOOKS:
        if sha256(checked_file(output / name, output)) != hashes[name]:
            raise ValueError(f'XLSM SHA256 mismatch: {name}')


def source_files(root):
    paths = {checked_file(root / name, root) for name in SOURCE_FILES}
    # Only code extensions in known source directories; no arbitrary local files,
    # databases, output workbooks, archived reports or build environments.
    for folder, extensions in [('fleet', {'.py'}), ('tools', {'.py', '.ps1'}),
                               ('tests', {'.py'}), ('excel', {'.bas', '.cls'}),
                               ('.github', {'.yml', '.yaml'})]:
        for path in (root / folder).rglob('*'):
            if '__pycache__' not in path.parts and path.suffix.lower() in extensions and path.is_file():
                paths.add(checked_file(path, root))
    return sorted(paths)


def checksum_text(folder, names):
    return ''.join(f'{sha256(folder / name)}  {name}\n' for name in names)


def package(output, *, root=ROOT):
    output, root = Path(output).resolve(), Path(root).resolve()
    validate(output)
    sources = source_files(root)
    created = ['WERYFIKACJA.md', 'Zrodla.zip', 'Floteka-pakiet.zip']
    demo = (output / 'Floteka-DEMO.sqlite3').exists()
    if demo:
        checked_file(output / 'Floteka-DEMO.sqlite3', output)
        created.append('Uruchom-DEMO.bat')
    for name in created:
        if (output / name).exists():
            raise FileExistsError(f'Refusing to replace existing delivery: {output / name}')
    instruction = (root / 'README.md').read_bytes()
    # build_exe.ps1 creates these two files. Accept only its exact, verified
    # contents before replacing the EXE-only checksum list with the full list.
    if (output / 'INSTRUKCJA.md').exists():
        if checked_file(output / 'INSTRUKCJA.md', output).read_bytes() != instruction:
            raise FileExistsError('Existing INSTRUKCJA.md differs from source README.')
    if (output / 'SUMY-SHA256.txt').exists():
        previous = checked_file(output / 'SUMY-SHA256.txt', output).read_text(encoding='utf-8-sig').strip()
        if previous != checksum_text(output, ['Floteka.exe']).strip():
            raise FileExistsError('Existing checksum file is not the EXE build checksum.')
    names = ['Floteka.exe', *WORKBOOKS, *REPORTS, 'INSTRUKCJA.md', 'WERYFIKACJA.md', 'Zrodla.zip']
    if demo:
        names += ['Floteka-DEMO.sqlite3', 'Uruchom-DEMO.bat']
    if (output / 'test-results.xml').exists():
        checked_file(output / 'test-results.xml', output)
        names.append('test-results.xml')
    with tempfile.TemporaryDirectory(prefix='floteka-package-', dir=output.parent) as folder:
        staging = Path(folder)
        for name in ['Floteka.exe', *WORKBOOKS, *REPORTS] + (['Floteka-DEMO.sqlite3'] if demo else []) + (['test-results.xml'] if 'test-results.xml' in names else []):
            shutil.copyfile(output / name, staging / name)
        (staging / 'INSTRUKCJA.md').write_bytes(instruction)
        shutil.copyfile(root / 'docs/verification.md', staging / 'WERYFIKACJA.md')
        if demo:
            (staging / 'Uruchom-DEMO.bat').write_bytes(
                b'@echo off\r\ncd /d "%~dp0"\r\nstart "" "%~dp0Floteka.exe" --db "%~dp0Floteka-DEMO.sqlite3"\r\n')
        with ZipFile(staging / 'Zrodla.zip', 'w', ZIP_DEFLATED) as archive:
            for path in sources:
                archive.write(path, path.relative_to(root).as_posix())
        (staging / 'SUMY-SHA256.txt').write_text(checksum_text(staging, names), encoding='utf-8')
        with ZipFile(staging / 'Floteka-pakiet.zip', 'w', ZIP_DEFLATED) as archive:
            for name in [*names, 'SUMY-SHA256.txt']:
                archive.write(staging / name, 'Floteka/' + name)
        (staging / 'SUMY-SHA256.txt').write_text(checksum_text(staging, [*names, 'Floteka-pakiet.zip']), encoding='utf-8')
        # Publish the delivery ZIP last: any earlier failure cannot leave an
        # apparently complete package behind.
        for name in [*[name for name in created if name != 'Floteka-pakiet.zip'],
                     'INSTRUKCJA.md', 'SUMY-SHA256.txt', 'Floteka-pakiet.zip']:
            if name == 'Floteka-pakiet.zip':
                # Staging shares the output filesystem. Windows rename is
                # atomic and refuses existing targets; POSIX link also provides
                # atomic no-overwrite publication (rename would overwrite).
                if os.name == 'nt':
                    os.rename(staging / name, output / name)
                else:
                    os.link(staging / name, output / name)
            elif name in ['INSTRUKCJA.md', 'SUMY-SHA256.txt'] and (output / name).exists():
                (staging / name).replace(output / name)
            else:
                with (output / name).open('xb') as target, (staging / name).open('rb') as source:
                    shutil.copyfileobj(source, target)
    return output / 'Floteka-pakiet.zip'


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    print('Pakiet:', package(parser.parse_args().output))
