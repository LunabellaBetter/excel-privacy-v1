"""Run V1 regressions; isolate Tk roots in fresh interpreter processes.

Tcl/Tk 9 on macOS can crash when many tests create/destroy separate Tcl
interpreters in a single process. Production starts one root per app process.
This runner reports subprocess failures, never converts a crash into a pass.
"""
import json
from pathlib import Path
import subprocess
import re
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'docs' / 'evidence'


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    started = time.time()
    args = [sys.executable, '-B', '-m', 'pytest']
    core = subprocess.run(args + ['tests', '--ignore=tests/test_gui_workflow.py', '-q',
                                  '--junitxml=' + str(OUT / 'pytest-core.xml')], cwd=ROOT)
    if core.returncode:
        return core.returncode
    collected = subprocess.run(args + ['tests/test_gui_workflow.py', '--collect-only', '-q'],
                               cwd=ROOT, capture_output=True, text=True, check=True)
    nodes = [line for line in collected.stdout.splitlines() if line.startswith('tests/') and '::' in line]
    if not nodes:
        raise RuntimeError('No GUI tests collected')
    results = []
    for index, node in enumerate(nodes, 1):
        child = subprocess.run(args + [node, '-q'], cwd=ROOT, capture_output=True, text=True, timeout=90)
        skipped = bool(re.search(r'\b\d+ skipped\b', child.stdout))
        failed = child.returncode != 0
        results.append({'test': node, 'exit': child.returncode,
                        'status': 'failed' if failed else 'skipped' if skipped else 'passed'})
        if failed:
            print(child.stdout, child.stderr, flush=True)
        if index % 20 == 0 or failed or index == len(nodes):
            print(f'GUI isolated: {index}/{len(nodes)}; failures={sum(r["exit"]!=0 for r in results)}', flush=True)
    report = {'core_exit': core.returncode, 'gui_processes': len(results),
              'gui_failed': sum(r['exit']!=0 for r in results),
              'gui_passed': sum(r['status']=='passed' for r in results),
              'gui_skipped': sum(r['status']=='skipped' for r in results),
              'elapsed_seconds': round(time.time() - started, 2), 'gui_results': results,
              'note': 'Each GUI test is run in a new interpreter; not a single-process full-suite pass.'}
    (OUT / 'pytest-gui-isolated.json').write_text(json.dumps(report, indent=2) + '\n')
    return 1 if report['gui_failed'] else 0


if __name__ == '__main__':
    raise SystemExit(main())
