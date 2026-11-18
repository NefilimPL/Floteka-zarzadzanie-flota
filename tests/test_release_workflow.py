"""Exercise the publication script with a local substitute for the GitHub CLI."""
import json
import os
from pathlib import Path
import shutil
import subprocess

import pytest
import yaml


ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = shutil.which('pwsh') or shutil.which('powershell')
pytestmark = pytest.mark.skipif(not POWERSHELL, reason='PowerShell is required for the publication script.')
COMMIT = 'a' * 40


def run_publication(tmp_path, *, dev=True, scenario='new', mismatch=False):
    version = '1.2.3.4-dev' if dev else '1.2.3.4'
    tag = 'v' + version
    folder = tmp_path / 'dist' / 'release'
    folder.mkdir(parents=True)
    for name in ('Floteka.exe', 'Floteka.xlsm', 'Floteka-DEMO.xlsm', 'INSTRUKCJA.md', 'exe-test.json'):
        (folder / name).write_text('test asset', encoding='utf-8')
    report = {'version': version, 'tag': tag, 'commit': COMMIT}
    (folder / 'build-info.json').write_text(json.dumps(report), encoding='utf-8')
    if mismatch:
        report['commit'] = 'b' * 40
    (folder / 'excel-build.json').write_text(json.dumps(report), encoding='utf-8')
    workflow = yaml.safe_load((ROOT / '.github' / 'workflows' / 'build-exe.yml').read_text(encoding='utf-8'))
    script = workflow['jobs']['release']['steps'][-1]['run']
    # This function shadows gh, records every API command and has no network code.
    mock = r'''
$ErrorActionPreference = 'Stop'
function global:gh {
    $flotekaCall = @($args)
    ConvertTo-Json -InputObject $flotekaCall -Compress | Add-Content calls.jsonl
    $global:LASTEXITCODE = 0
    if ($args[0] -eq 'api') {
        if ($env:TEST_SCENARIO -eq 'moved-tag') { 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb' }
        else { $env:FLOTEKA_SOURCE }
    } elseif ($args[1] -eq 'view') {
        switch ($env:TEST_SCENARIO) {
            'new' { $global:LASTEXITCODE = 1 }
            'moved-tag' { $global:LASTEXITCODE = 1 }
            'draft' { '{"isDraft":true,"assets":[{"name":"Floteka.exe"}]}' }
            'extra-asset' { '{"isDraft":true,"assets":[{"name":"private.sqlite3"}]}' }
            'published' { '{"isDraft":false,"assets":[]}' }
        }
    }
}
'''
    harness = tmp_path / 'publish-test.ps1'
    harness.write_text(mock + '\n' + script, encoding='utf-8')
    env = dict(os.environ, GH_TOKEN='unused-test-token', GH_REPO='example/floteka',
               FLOTEKA_TAG=tag, FLOTEKA_VERSION=version, FLOTEKA_SOURCE=COMMIT,
               FLOTEKA_PRERELEASE=str(dev).lower(), TEST_SCENARIO=scenario)
    result = subprocess.run([POWERSHELL, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(harness)],
                            cwd=tmp_path, env=env, capture_output=True, text=True)
    calls_file = tmp_path / 'calls.jsonl'
    calls = [json.loads(line) for line in calls_file.read_text(encoding='utf-8-sig').splitlines()] if calls_file.exists() else []
    return result, calls


@pytest.mark.parametrize('dev', [False, True])
def test_release_uploads_both_workbooks_and_exe_before_publishing(tmp_path, dev):
    result, calls = run_publication(tmp_path, dev=dev)
    assert result.returncode == 0, result.stderr
    create = next(call for call in calls if call[:2] == ['release', 'create'])
    edit = next(call for call in calls if call[:2] == ['release', 'edit'])
    assert '--draft' in create
    assert 'dist\\release\\Floteka.exe' in create or 'dist/release/Floteka.exe' in create
    assert any(str(arg).endswith('Floteka.xlsm') for arg in create)
    assert any(str(arg).endswith('Floteka-DEMO.xlsm') for arg in create)
    assert calls.index(create) < calls.index(edit)
    assert '--draft=false' in edit
    assert ('--prerelease' in create) is dev
    assert f'--prerelease={str(dev).lower()}' in edit
    assert ('--latest=false' in edit) is dev
    sums = (tmp_path / 'dist' / 'release' / 'SUMY-SHA256.txt').read_text(encoding='utf-8-sig')
    assert '  Floteka.xlsm' in sums and '  Floteka.exe' in sums


def test_retry_completes_only_the_expected_draft_assets(tmp_path):
    result, calls = run_publication(tmp_path, scenario='draft')
    assert result.returncode == 0, result.stderr
    assert any(call[:2] == ['release', 'upload'] for call in calls)
    assert calls[-1][:2] == ['release', 'edit']


@pytest.mark.parametrize('scenario', ['extra-asset', 'published'])
def test_existing_release_is_not_published_or_overwritten_unsafely(tmp_path, scenario):
    result, calls = run_publication(tmp_path, scenario=scenario)
    assert result.returncode != 0
    assert not any(call[0] == 'release' and call[1] in ('create', 'upload', 'edit') for call in calls)


def test_different_exe_and_excel_commits_stop_before_github_calls(tmp_path):
    result, calls = run_publication(tmp_path, mismatch=True)
    assert result.returncode != 0
    assert calls == []


def test_moved_stable_tag_cannot_publish_the_wrong_source(tmp_path):
    result, calls = run_publication(tmp_path, dev=False, scenario='moved-tag')
    assert result.returncode != 0
    assert not any(call[:2] == ['release', 'create'] for call in calls)
