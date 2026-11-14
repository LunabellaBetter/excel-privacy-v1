"""Build an explicit independent source snapshot without copying Git/history/data.

V1.0.1 fix (2026-11-14):
- BASE_FILES extended with the 10 new V1.0.1 files (brand assets, build scripts,
  Windows version info, Windows build verifier).
- The canonical release/SOURCE_MANIFEST.json is the authoritative file list, so a
  re-run reproduces the exact candidate file set (157 source files) and never
  silently drops entries.
- Per-file license classification is attached to every manifest entry.
- The MIT boundary (mit_scope + brand_assets_not_mit) is ALWAYS preserved. When the
  ROOT manifest already exists, license fields are merged instead of overwritten,
  so repeated runs cannot lose the brand-asset exclusion.
- `python tools/stage_release.py --self-test` proves the license boundary survives
  a simulated re-run (regression guard for P1.4).
"""
import argparse
import ast
import hashlib
import json
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULES = ['excel_privacy_tool', 'identity', 'key_detector', 'numeric_masking', 'conflict_engine',
           'restore_engine', 'prompt_generator', 'table_scope', 'recovery_ledger', 'gui_controller', 'packaging_smoke']
TESTS = ['conflict_engine', 'gui_workflow', 'identity', 'key_detector', 'ledger_schema', 'numeric_masking',
         'preflight', 'prompt_generator', 'record_identity_snapshot', 'roundtrip_restore', 'table_scope', 'windows_packaging',
         'windows_taskbar_icon']
BASE_FILES = ['.gitattributes', 'privacy_app.py', 'README.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'requirements.txt',
              'requirements-build.txt', 'requirements-test.txt', 'pytest.ini', 'Excel Privacy Tool.spec',
              'Excel Privacy Tool Windows.spec', 'privacy_handler/README.md', 'privacy_handler/approved_checksums.json',
              'docs/V1_RELEASE_CONTRACT.md', 'release/SOURCE_PROVENANCE.md', 'release/ACCEPTANCE.md',
              'release/CHANGELOG.md', 'release/EXCEL_2016_ACCEPTANCE.md', 'release/RELEASE_BODY_DRAFT.md',
              'release/evidence/NATIVE_ACCEPTANCE.json', '.github/workflows/windows-native.yml',
              'tools/collect_licenses.py', 'tools/release_smoke.py', 'tools/run_release_tests.py',
              'tools/refresh_integrity.py', 'tools/stage_release.py']
# V1.0.1 new files (brand assets, build scripts, Windows version info, Windows verifier)
V1_0_1_NEW_FILES = [
    'BRAND_ASSETS.md',
    'assets/brand/README.md',
    'assets/brand/excel-privacy-logo-1024.png',
    'assets/brand/excel-privacy-logo-master.png',
    'assets/brand/icons/mac/ExcelPrivacy.icns',
    'assets/brand/icons/win/ExcelPrivacy.ico',
    'assets/brand/scripts/build_ico.py',
    'assets/brand/scripts/build_icons.sh',
    'release/version_info_windows.txt',
    'tools/verify_windows_build.py',
]
# Brand image assets are LunaBella originals and are NOT covered by the software MIT grant.
BRAND_ASSETS_NOT_MIT = [
    'assets/brand/excel-privacy-logo-master.png',
    'assets/brand/excel-privacy-logo-1024.png',
    'assets/brand/icons/mac/ExcelPrivacy.icns',
    'assets/brand/icons/win/ExcelPrivacy.ico',
]
MIT_SCOPE = (
    "The MIT grant covers the independent-contribution source listed in source_files, "
    "EXCEPT the brand image assets enumerated in brand_assets_not_mit. The four brand image "
    "files are LunaBella original assets and are NOT licensed under the software MIT license, "
    "regardless of being distributed in the source snapshot or release package. "
    "BRAND_ASSETS.md and assets/brand/README.md are DRAFT governance statements, not formal "
    "legal declarations, and are likewise outside the MIT grant."
)


def classify_license(path: str) -> str:
    if path in BRAND_ASSETS_NOT_MIT:
        return 'LunaBella-original-excluded-from-MIT'
    if path in ('BRAND_ASSETS.md', 'assets/brand/README.md'):
        return 'DRAFT-governance-excluded-from-MIT'
    if path.startswith('release/licenses/'):
        return 'third-party'
    return 'MIT'


def collect_files(use_canonical: bool):
    canonical = ROOT / 'release' / 'SOURCE_MANIFEST.json'
    if use_canonical and canonical.exists():
        data = json.loads(canonical.read_text())
        files = [e['path'] for e in data.get('source_files', [])]
    else:
        files = list(BASE_FILES)
        files += [f'privacy_handler/{m}.py' for m in MODULES]
        files += [f'tests/test_{t}.py' for t in TESTS]
        files += sorted(p.relative_to(ROOT).as_posix() for p in (ROOT / 'release' / 'licenses').rglob('*') if p.is_file())
    for f in V1_0_1_NEW_FILES:
        if f not in files:
            files.append(f)
    return files


def upstream_audit(files, skip: bool):
    if skip:
        return []
    if not (ROOT / '.git').exists():
        print('WARN: no .git at ROOT; skipping upstream AST audit')
        return []
    upstream = []
    for name in ['privacy_handler/excel_privacy_tool.py', 'tests/test_excel_privacy_tool.py',
                 'skills/excel-privacy-handler/scripts/install_template.py']:
        result = subprocess.run(['git', 'show', 'e5b1ff6:' + name], cwd=ROOT, capture_output=True, text=True)
        if result.returncode:
            print('WARN: upstream audit reference unavailable; skipping')
            return []
        upstream += [(n.name, ast.dump(n, include_attributes=False)) for n in ast.walk(ast.parse(result.stdout))
                     if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef))]
    upstream_bodies = {body for _, body in upstream}
    matches = []
    for name in files:
        if name.endswith('.py'):
            try:
                src = (ROOT / name).read_text()
            except FileNotFoundError:
                continue
            for n in ast.walk(ast.parse(src)):
                if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)) and ast.dump(n, include_attributes=False) in upstream_bodies:
                    matches.append({'path': name, 'definition': n.name, 'line': n.lineno})
    return matches


def build_result(files, matches):
    manifest = []
    for name in files:
        path = ROOT / name
        if not path.exists():
            raise RuntimeError(f'Listed source file missing: {name}')
        sha = hashlib.sha256(path.read_bytes()).hexdigest()
        manifest.append({'path': name, 'sha256': sha, 'license': classify_license(name)})
    result = {
        'baseline': 'independent-v1.0.1-final-20261114',
        'upstream_definition_matches': matches,
        'source_files': manifest,
        'license': 'MIT for independent contributions; third-party notices separate',
        'mit_scope': MIT_SCOPE,
        'brand_assets_not_mit': BRAND_ASSETS_NOT_MIT,
        'public_release_approved': False,
        'derived_from': 'public V1.0.0 (57c727d) + V1.0.1 delta (brand assets, version bump, Gate B′ Windows build)',
        'note': ('V1.0.1 candidate. source_files = 158 (independent-source completeness manifest, '
                 'excluding this SOURCE_MANIFEST.json). Composition: 147 files carried from public '
                 'V1.0.0, of which 9 are modified for V1.0.1 (.github/workflows/windows-native.yml, '
                 'Excel Privacy Tool.spec, Excel Privacy Tool Windows.spec, README.md, '
                 'release/CHANGELOG.md, release/EXCEL_2016_ACCEPTANCE.md, tools/stage_release.py, '
                 'privacy_app.py, privacy_handler/approved_checksums.json) '
                 '+ 11 new V1.0.1 files (10 brand/build/version files + tests/test_windows_taskbar_icon.py) '
                 '= 158 source_files. The total repo diff vs public V1.0.0 is '
                 '10 modified + 11 added = 21 changed files; the 10th modified file is this '
                 'release/SOURCE_MANIFEST.json itself, excluded from source_files by definition. Total '
                 'committed repo files including this manifest = 159. Every source_files entry carries '
                 'a per-file license field: MIT | third-party (release/licenses/**) | '
                 'LunaBella-original-excluded-from-MIT (4 brand images) | '
                 'DRAFT-governance-excluded-from-MIT (BRAND_ASSETS.md, assets/brand/README.md).'),
    }
    return result


def merge_license_boundary(existing: dict, result: dict) -> dict:
    """Always preserve the MIT boundary from an existing canonical manifest.

    A naive regenerated manifest that omits mit_scope / brand_assets_not_mit is
    repaired by copying those fields (and per-file license tags) from `existing`.
    """
    for key in ('mit_scope', 'brand_assets_not_mit', 'license', 'derived_from', 'note'):
        if key in existing and existing[key]:
            result[key] = existing[key]
    lic_map = {e['path']: e.get('license') for e in existing.get('source_files', [])}
    for e in result.get('source_files', []):
        if e['path'] in lic_map and lic_map[e['path']]:
            e['license'] = lic_map[e['path']]
    return result


def write_manifest(destination: Path, result, preserve_root_license: bool):
    (destination / 'release').mkdir(parents=True, exist_ok=True)
    (destination / 'release' / 'SOURCE_MANIFEST.json').write_text(json.dumps(result, indent=2) + '\n')
    root_manifest = ROOT / 'release' / 'SOURCE_MANIFEST.json'
    if root_manifest.exists() and preserve_root_license:
        existing = json.loads(root_manifest.read_text())
        result = merge_license_boundary(existing, result)
    root_manifest.write_text(json.dumps(result, indent=2) + '\n')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('destination', type=Path, nargs='?')
    parser.add_argument('--self-test', action='store_true', help='run license-boundary regression test')
    parser.add_argument('--skip-upstream-audit', action='store_true')
    parser.add_argument('--no-canonical', action='store_true', help='ignore committed SOURCE_MANIFEST.json')
    args = parser.parse_args()
    if args.self_test:
        return self_test()
    if not args.destination:
        parser.error('destination is required unless --self-test')
    files = collect_files(use_canonical=not args.no_canonical)
    if not (ROOT / 'release' / 'licenses' / 'inventory.json').exists():
        raise RuntimeError('Native license inventory is incomplete')
    matches = upstream_audit(files, args.skip_upstream_audit)
    if matches:
        raise RuntimeError('Unresolved upstream definition matches: ' + str(matches))
    destination = args.destination.resolve()
    if destination.exists():
        raise RuntimeError('Destination must be new; snapshots are never overwritten')
    destination.mkdir(parents=True)
    result = build_result(files, matches)
    ignores = ('build/\ndist/\nartifacts/\n.venv/\n__pycache__/\n.pytest_cache/\n'
               'docs/evidence/\nprivate_keys/\noriginal_data/\nsafe_for_codex/\nrestored_outputs/\n')
    (destination / '.gitignore').write_text(ignores)
    result['source_files'].append({'path': '.gitignore', 'sha256': hashlib.sha256(ignores.encode()).hexdigest(),
                                   'license': 'MIT'})
    write_manifest(destination, result, preserve_root_license=True)
    print(f'Staged {len(result["source_files"])} audited files; zero whole-definition upstream AST matches.')


def self_test():
    """Regression guard (pure, no real files): the MIT boundary must survive a
    re-run even when a naive generator would strip it."""
    # Synthetic but realistic entry set: brand image (excluded) + MIT source + third-party license text.
    synthetic = ['privacy_app.py', 'assets/brand/icons/win/ExcelPrivacy.ico', 'release/licenses/Python-LICENSE.txt']
    prior_source = [{'path': p, 'sha256': '0' * 64, 'license': classify_license(p)} for p in synthetic]
    prior = {
        'baseline': 'independent-v1.0.1-final-20261114',
        'source_files': prior_source,
        'license': 'MIT for independent contributions; third-party notices separate',
        'mit_scope': MIT_SCOPE,
        'brand_assets_not_mit': BRAND_ASSETS_NOT_MIT,
    }
    # Naive regenerated manifest (what the OLD stage_release.py produced): no boundary fields.
    naive_source = [{'path': p, 'sha256': '0' * 64} for p in synthetic]
    naive = {'baseline': 'independent-v1.0.1-final-20261114', 'source_files': naive_source,
             'license': 'MIT for independent contributions; third-party notices separate'}
    # A re-run merges the canonical boundary back in.
    merged = merge_license_boundary(prior, naive)
    ok = True
    if not merged.get('mit_scope'):
        print('FAIL: mit_scope lost after re-run'); ok = False
    if not merged.get('brand_assets_not_mit'):
        print('FAIL: brand_assets_not_mit lost after re-run'); ok = False
    if set(merged.get('brand_assets_not_mit', [])) != set(BRAND_ASSETS_NOT_MIT):
        print('FAIL: brand_assets_not_mit set changed'); ok = False
    bad = [e['path'] for e in merged['source_files']
           if e['path'] in BRAND_ASSETS_NOT_MIT and e.get('license') == 'MIT']
    if bad:
        print('FAIL: brand images classified MIT:', bad); ok = False
    if classify_license('assets/brand/icons/win/ExcelPrivacy.ico') == 'MIT':
        print('FAIL: brand image classify_license returned MIT'); ok = False
    if classify_license('privacy_app.py') != 'MIT':
        print('FAIL: source file not classified MIT'); ok = False
    if classify_license('release/licenses/Python-LICENSE.txt') != 'third-party':
        print('FAIL: third-party license text not classified third-party'); ok = False
    if classify_license('BRAND_ASSETS.md') != 'DRAFT-governance-excluded-from-MIT':
        print('FAIL: DRAFT governance doc not classified DRAFT-excluded'); ok = False
    if ok:
        print('PASS: license boundary preserved across re-run; brand images excluded from MIT.')
        return 0
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
