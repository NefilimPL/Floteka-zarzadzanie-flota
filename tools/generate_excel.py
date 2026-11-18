"""Generate versioned XLSM files from verified templates without Microsoft Office."""
import argparse
import hashlib
import io
import json
import os
from pathlib import Path
import re
import sys
from xml.dom import minidom
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tools.release_version import prepare_version

NAMES = ('Floteka.xlsm', 'Floteka-DEMO.xlsm')
SOURCE_FILES = ('excel/Flota.bas', 'excel/ThisWorkbook.cls', 'tools/build_excel.py')
SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CUSTOM_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
DC_NS = 'http://purl.org/dc/elements/1.1/'


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def source_hashes(root):
    # Git may check out text files with CRLF or LF; both describe the same VBA.
    return {name: sha256((Path(root) / name).read_text(encoding='utf-8-sig').encode('utf-8'))
            for name in SOURCE_FILES}


def verified_templates(root):
    folder = Path(root) / 'excel' / 'templates'
    try:
        manifest = json.loads((folder / 'manifest.json').read_text(encoding='utf-8'))
        if manifest['schema'] != 1 or manifest['source_hashes'] != source_hashes(root):
            raise ValueError('The XLSM templates are stale. Refresh templates after changing VBA or layout.')
        templates = {name: (folder / name).read_bytes() for name in NAMES}
        for name, content in templates.items():
            if sha256(content) != manifest['files'][name]['sha256']:
                raise ValueError(f'Modified XLSM template: {name}. Refresh the templates.')
    except (OSError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise ValueError('Missing or invalid XLSM templates. Run tools/refresh_excel_templates.py on a development machine with Excel.') from exc
    return manifest, templates


def _text(node):
    return ''.join(child.data for child in node.childNodes if child.nodeType == child.TEXT_NODE)


def _replace_text(node, value):
    for child in list(node.childNodes):
        node.removeChild(child)
    node.appendChild(node.ownerDocument.createTextNode(value))


def versioned_workbook(content, old_version, version):
    """Leave binary VBA, controls, drawing anchors and ZIP metadata untouched."""
    result = io.BytesIO()
    counts = {'title': 0, 'property': 0, 'label': 0, 'footer': 0}
    with ZipFile(io.BytesIO(content)) as before, ZipFile(result, 'w') as after:
        if before.testzip() is not None or 'xl/vbaProject.bin' not in before.namelist():
            raise ValueError('The template is not an intact macro workbook.')
        for member in before.infolist():
            data = before.read(member.filename)
            sheet = bool(re.fullmatch(r'xl/worksheets/sheet\d+\.xml', member.filename))
            if member.filename in ('docProps/core.xml', 'docProps/custom.xml', 'xl/sharedStrings.xml') or sheet:
                # DOM retains namespace declarations including the prefixes referenced
                # by mc:Ignorable. Rewriting sheets with ElementTree can remove them.
                doc = minidom.parseString(data)
                changed = False
                if member.filename == 'docProps/core.xml':
                    for node in doc.getElementsByTagNameNS(DC_NS, 'title'):
                        if _text(node) == f'Floteka {old_version}':
                            _replace_text(node, f'Floteka {version}')
                            counts['title'] += 1
                            changed = True
                elif member.filename == 'docProps/custom.xml':
                    for prop in doc.getElementsByTagNameNS(CUSTOM_NS, 'property'):
                        if prop.getAttribute('name') == 'FlotekaVersion':
                            values = [child for child in prop.childNodes if child.nodeType == child.ELEMENT_NODE]
                            if len(values) != 1 or _text(values[0]) != old_version:
                                raise ValueError('Invalid template version property.')
                            _replace_text(values[0], version)
                            counts['property'] += 1
                            changed = True
                for node in doc.getElementsByTagNameNS(SHEET_NS, 't'):
                    text = _text(node)
                    prefix = f'Wersja {old_version} |'
                    if text.startswith(prefix):
                        _replace_text(node, f'Wersja {version} |' + text[len(prefix):])
                        counts['label'] += 1
                        changed = True
                for node in doc.getElementsByTagNameNS(SHEET_NS, 'oddFooter'):
                    text = _text(node)
                    if f'Floteka {old_version}' in text:
                        _replace_text(node, text.replace(f'Floteka {old_version}', f'Floteka {version}'))
                        counts['footer'] += 1
                        changed = True
                if changed:
                    data = doc.toxml(encoding='UTF-8')
                doc.unlink()
            after.writestr(member, data)
        after.comment = before.comment
    if counts != {'title': 1, 'property': 1, 'label': 1, 'footer': 7}:
        raise ValueError(f'Unexpected template version fields: {counts}. Refresh the templates.')
    return result.getvalue()


def generate(output, tag=None, *, root=ROOT):
    root, output = Path(root).resolve(), Path(output).resolve()
    manifest, templates = verified_templates(root)
    names = (*NAMES, 'excel-build.json')
    for name in names:
        if (output / name).exists():
            raise FileExistsError(f'Existing file will not be overwritten: {output / name}')
    release = prepare_version(root, tag, require_commit=os.environ.get('GITHUB_ACTIONS') == 'true')
    files = {name: versioned_workbook(content, manifest['version'], release.version)
             for name, content in templates.items()}
    report = {'ok': True, 'engine': 'template', 'macro_tests_run': False,
              'version': release.version, 'tag': release.tag, 'commit': release.commit,
              'files': list(NAMES), 'hashes': {name: sha256(data) for name, data in files.items()},
              'template_version': manifest['version'],
              'template_verified_in_excel': manifest['verified_in_excel'],
              'source_hashes': manifest['source_hashes']}
    files['excel-build.json'] = (json.dumps(report, ensure_ascii=False, indent=2) + '\n').encode('utf-8')
    output.mkdir(parents=True, exist_ok=True)
    created = []
    try:
        for name, data in files.items():
            target = output / name
            with target.open('xb') as file:
                created.append(target)
                file.write(data)
    except BaseException:
        for target in created:
            target.unlink(missing_ok=True)
        raise
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--tag', default=os.environ.get('FLOTEKA_BUILD_TAG') or None)
    args = parser.parse_args()
    try:
        report = generate(args.output, args.tag)
    except (OSError, ValueError) as exc:
        parser.exit(1, f'XLSM generation failed: {exc}\n')
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
