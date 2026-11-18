"""Development maintenance: rebuild and verify XLSM templates using installed Excel.

Normal users run Buduj.bat; they never need this maintenance command or Office.
"""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tools.generate_excel import NAMES, sha256, source_hashes, versioned_workbook


def refresh():
    folder = ROOT / 'excel' / 'templates'
    (ROOT / 'build').mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='excel-templates-', dir=ROOT / 'build') as staging:
        staging = Path(staging)
        subprocess.run([sys.executable, str(ROOT / 'tools' / 'build_excel.py'), '--output', str(staging)], cwd=ROOT, check=True)
        env = dict(os.environ, FLOTA_EXCEL_FILE=str(staging / NAMES[0]), FLOTEKA_EXCEL_VERSION_DIR=str(staging))
        subprocess.run([sys.executable, '-m', 'pytest', 'tests/test_excel_layout.py', 'tests/test_excel_version.py',
                        '--capture=sys', '-q', '-p', 'no:cacheprovider'], cwd=ROOT, env=env, check=True)
        checks = ('test_excel.py', 'test_excel_regressions.py', 'test_deadlines_excel.py', 'test_optional_excel.py')
        for name in checks:
            subprocess.run([sys.executable, str(ROOT / 'tools' / name)], cwd=ROOT, env=env, check=True)
        report = json.loads((ROOT / 'docs' / 'excel-build.json').read_text(encoding='utf-8'))
        if not report['ok']:
            raise ValueError('The Excel template build failed.')
        files = {}
        for name in NAMES:
            content = (staging / name).read_bytes()
            # Exercise the exact ZIP/XML path used by the Office-free generator.
            versioned_workbook(content, report['version'], report['version'])
            files[name] = {'sha256': sha256(content)}
        manifest = {'schema': 1, 'version': report['version'], 'source_hashes': source_hashes(ROOT),
                    'files': files, 'verified_in_excel': {
                        'version': report['excel'], 'utc': datetime.now(timezone.utc).isoformat(),
                        'checks': list(checks), 'layout_and_version_tests': True}}
        folder.mkdir(parents=True, exist_ok=True)
        for name in NAMES:
            shutil.copyfile(staging / name, folder / name)
        (folder / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        print('Refreshed verified XLSM templates:', folder)


if __name__ == '__main__':
    refresh()
