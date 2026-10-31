"""Development-only regeneration of the reviewed V1 resource manifest."""
from pathlib import Path
import json
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from privacy_handler.excel_privacy_tool import APPROVED_CHECKSUMS, current_checksums

if __name__ == '__main__':
    APPROVED_CHECKSUMS.write_text(json.dumps(current_checksums(), indent=2) + '\n', encoding='utf-8')
    print('Reviewed resource manifest regenerated; rerun tests and verify-self.')
