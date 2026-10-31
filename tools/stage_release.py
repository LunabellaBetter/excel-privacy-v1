"""Build an explicit independent source snapshot without copying Git/history/data."""
import argparse
import ast
import hashlib
import json
from pathlib import Path
import shutil
import subprocess

ROOT = Path(__file__).resolve().parents[1]
MODULES = ['excel_privacy_tool','identity','key_detector','numeric_masking','conflict_engine',
           'restore_engine','prompt_generator','table_scope','recovery_ledger','gui_controller','packaging_smoke']
TESTS = ['conflict_engine','gui_workflow','identity','key_detector','ledger_schema','numeric_masking',
         'preflight','prompt_generator','record_identity_snapshot','roundtrip_restore','table_scope','windows_packaging']
BASE_FILES = ['.gitattributes','privacy_app.py','README.md','LICENSE','THIRD_PARTY_NOTICES.md','requirements.txt',
              'requirements-build.txt','requirements-test.txt','pytest.ini','Excel Privacy Tool.spec',
              'Excel Privacy Tool Windows.spec','privacy_handler/README.md','privacy_handler/approved_checksums.json',
              'docs/V1_RELEASE_CONTRACT.md','release/SOURCE_PROVENANCE.md','release/ACCEPTANCE.md',
              'release/CHANGELOG.md','release/EXCEL_2016_ACCEPTANCE.md','release/RELEASE_BODY_DRAFT.md',
              'release/evidence/NATIVE_ACCEPTANCE.json','.github/workflows/windows-native.yml',
              'tools/collect_licenses.py','tools/release_smoke.py','tools/run_release_tests.py',
              'tools/refresh_integrity.py','tools/stage_release.py']


def main():
    parser=argparse.ArgumentParser();parser.add_argument('destination',type=Path);args=parser.parse_args()
    files=BASE_FILES+[f'privacy_handler/{m}.py' for m in MODULES]+[f'tests/test_{t}.py' for t in TESTS]
    files+=sorted(p.relative_to(ROOT).as_posix() for p in (ROOT/'release/licenses').rglob('*') if p.is_file())
    if not (ROOT/'release/licenses/inventory.json').exists():raise RuntimeError('Native license inventory is incomplete')
    upstream=[]
    for name in ['privacy_handler/excel_privacy_tool.py','tests/test_excel_privacy_tool.py','skills/excel-privacy-handler/scripts/install_template.py']:
        result=subprocess.run(['git','show','e5b1ff6:'+name],cwd=ROOT,capture_output=True,text=True)
        if result.returncode:raise RuntimeError('Upstream audit reference unavailable')
        upstream += [(n.name,ast.dump(n,include_attributes=False)) for n in ast.walk(ast.parse(result.stdout))
                     if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef,ast.ClassDef))]
    upstream_bodies={body for name,body in upstream}
    matches=[]
    for name in files:
        if name.endswith('.py'):
            for n in ast.walk(ast.parse((ROOT/name).read_text())):
                if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef,ast.ClassDef)) and ast.dump(n,include_attributes=False) in upstream_bodies:
                    matches.append({'path':name,'definition':n.name,'line':n.lineno})
    if matches:raise RuntimeError('Unresolved upstream definition matches: '+str(matches))
    destination=args.destination.resolve()
    if destination.exists():raise RuntimeError('Destination must be new; snapshots are never overwritten')
    destination.mkdir(parents=True)
    manifest=[]
    for name in files:
        path=ROOT/name;target=destination/name;target.parent.mkdir(parents=True,exist_ok=True)
        shutil.copyfile(path,target)
        manifest.append({'path':name,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
    ignores='build/\ndist/\nartifacts/\n.venv/\n__pycache__/\n.pytest_cache/\ndocs/evidence/\nprivate_keys/\noriginal_data/\nsafe_for_codex/\nrestored_outputs/\n'
    (destination/'.gitignore').write_text(ignores)
    manifest.append({'path':'.gitignore','sha256':hashlib.sha256(ignores.encode()).hexdigest()})
    result={'baseline':'independent-v1.0.0','upstream_definition_matches':matches,
            'source_files':manifest,'license':'MIT for independent contributions; third-party notices separate',
            'public_release_approved':False}
    (destination/'release/SOURCE_MANIFEST.json').write_text(json.dumps(result,indent=2)+'\n')
    (ROOT/'release/SOURCE_MANIFEST.json').write_text(json.dumps(result,indent=2)+'\n')
    print(f'Staged {len(manifest)} audited files; zero whole-definition upstream AST matches.')


if __name__=='__main__':main()
