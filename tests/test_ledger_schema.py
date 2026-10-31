import datetime as dt
import json
import sys
from pathlib import Path
import pytest
from cryptography.fernet import Fernet

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import privacy_handler.excel_privacy_tool as tool
import privacy_handler.recovery_ledger as ledger_mod
from privacy_handler.recovery_ledger import (
    LedgerConfig,
    LedgerEntry,
    LedgerRunInfo,
    ProtectedField,
    RecoveryLedger,
    RecoveryLedgerError,
    load_encrypted_ledger,
    save_encrypted_ledger,
)


def sample_ledger(run_id="run_test_001"):
    return RecoveryLedger(
        protocol_version="1.0",
        run=LedgerRunInfo(
            run_id=run_id,
            created_at=dt.datetime.now(dt.timezone.utc).isoformat(),
            workbook_identity="wb_hash_test_12345",
            source_extension=".xlsx",
        ),
        configuration=LedgerConfig(
            confirmed_key_field="员工工号",
            numeric_mode=None,
        ),
        protected_fields=[
            ProtectedField(logical_field="员工工号", protection_type="text_token"),
            ProtectedField(logical_field="姓名", protection_type="text_token"),
            ProtectedField(logical_field="基本工资", protection_type="numeric_scale"),
        ],
        entries=[
            LedgerEntry(
                sheet="人员总表",
                logical_field="员工工号",
                key_masked_value=None,
                original_type="str",
                original_value="00001234",
                masked_type="str",
                masked_value="MASK_EMP_001",
                protection_type="text_token",
            ),
            LedgerEntry(
                sheet="人员总表",
                logical_field="姓名",
                key_masked_value="MASK_EMP_001",
                original_type="str",
                original_value="张三",
                masked_type="str",
                masked_value="MASK_NAME_001",
                protection_type="text_token",
            ),
            LedgerEntry(
                sheet="人员总表",
                logical_field="基本工资",
                key_masked_value="MASK_EMP_001",
                original_type="float",
                original_value=8500.50,
                masked_type="float",
                masked_value=11662.68,
                protection_type="numeric_scale",
            ),
        ],
        numeric_groups=[
            {
                "group_id": "group_comp_01",
                "mode": "positive_scale",
                "fields": ["基本工资"],
                "scale_factor": 1.372,
            }
        ],
    )


def test_v1_ledger_roundtrip(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "test.recovery_ledger.enc"
    orig_ledger = sample_ledger()

    save_encrypted_ledger(orig_ledger, ledger_path, secret_key)
    assert ledger_path.exists()

    loaded = load_encrypted_ledger(ledger_path, secret_key)
    assert loaded.protocol_version == "1.0"
    assert loaded.run.run_id == orig_ledger.run.run_id
    assert loaded.run.workbook_identity == orig_ledger.run.workbook_identity
    assert loaded.configuration.confirmed_key_field == "员工工号"
    assert len(loaded.protected_fields) == 3
    assert len(loaded.entries) == 3
    assert loaded.numeric_groups == orig_ledger.numeric_groups

    # Compare dictionary representation
    assert loaded.to_dict() == orig_ledger.to_dict()


def test_ledger_typed_values_roundtrip(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "typed.recovery_ledger.enc"

    test_date = dt.date(2026, 10, 7)
    test_datetime = dt.datetime(2026, 10, 7, 11, 30, 45)

    entries = [
        LedgerEntry("S1", "IntField", None, "int", 42, "int", 84),
        LedgerEntry("S1", "FloatField", None, "float", 123.456, "float", 246.912),
        LedgerEntry("S1", "BoolField", None, "bool", True, "bool", False),
        LedgerEntry("S1", "StrField", None, "str", "Hello World", "str", "MASK_TXT"),
        LedgerEntry("S1", "DateField", None, "date", test_date.isoformat(), "str", "MASK_DATE"),
        LedgerEntry("S1", "DateTimeField", None, "datetime", test_datetime.isoformat(), "str", "MASK_DT"),
    ]

    ledger = sample_ledger()
    ledger.entries = entries

    save_encrypted_ledger(ledger, ledger_path, secret_key)
    loaded = load_encrypted_ledger(ledger_path, secret_key)

    assert loaded.entries[0].get_decoded_original_value() == 42
    assert isinstance(loaded.entries[0].get_decoded_original_value(), int)

    assert loaded.entries[1].get_decoded_original_value() == 123.456
    assert isinstance(loaded.entries[1].get_decoded_original_value(), float)

    assert loaded.entries[2].get_decoded_original_value() is True
    assert isinstance(loaded.entries[2].get_decoded_original_value(), bool)

    assert loaded.entries[3].get_decoded_original_value() == "Hello World"
    assert isinstance(loaded.entries[3].get_decoded_original_value(), str)

    assert loaded.entries[4].get_decoded_original_value() == test_date
    assert isinstance(loaded.entries[4].get_decoded_original_value(), dt.date)

    assert loaded.entries[5].get_decoded_original_value() == test_datetime
    assert isinstance(loaded.entries[5].get_decoded_original_value(), dt.datetime)


def test_ledger_preserves_leading_zero_strings(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "leading_zero.recovery_ledger.enc"

    ledger = sample_ledger()
    ledger.entries = [
        LedgerEntry("S1", "CodeField", None, "str", "00001234", "str", "MASK_001"),
        LedgerEntry("S1", "ZipField", None, "str", "010000", "str", "MASK_002"),
    ]

    save_encrypted_ledger(ledger, ledger_path, secret_key)
    loaded = load_encrypted_ledger(ledger_path, secret_key)

    val1 = loaded.entries[0].get_decoded_original_value()
    assert val1 == "00001234"
    assert isinstance(val1, str)
    assert val1 != 1234

    val2 = loaded.entries[1].get_decoded_original_value()
    assert val2 == "010000"
    assert isinstance(val2, str)


def test_ledger_wrong_secret_fails(tmp_path):
    secret_key1 = Fernet.generate_key()
    secret_key2 = Fernet.generate_key()
    ledger_path = tmp_path / "test.recovery_ledger.enc"

    save_encrypted_ledger(sample_ledger(), ledger_path, secret_key1)

    with pytest.raises(RecoveryLedgerError) as exc:
        load_encrypted_ledger(ledger_path, secret_key2)
    assert "secret key may be incorrect" in str(exc.value)


def test_ledger_corruption_fails(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "corrupt.recovery_ledger.enc"

    save_encrypted_ledger(sample_ledger(), ledger_path, secret_key)
    raw = ledger_path.read_bytes()

    # Corrupt some bytes in the middle of ciphertext
    corrupted = raw[:20] + b"XXXXCORRUPT" + raw[31:]
    ledger_path.write_bytes(corrupted)

    with pytest.raises(RecoveryLedgerError):
        load_encrypted_ledger(ledger_path, secret_key)


def test_unsupported_protocol_version_fails(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "version2.recovery_ledger.enc"

    # Manually construct future v2.0 JSON payload
    data = sample_ledger().to_dict()
    data["protocol_version"] = "2.0"
    json_bytes = json.dumps(data).encode("utf-8")
    ciphertext = Fernet(secret_key).encrypt(json_bytes)
    ledger_path.write_bytes(ciphertext)

    with pytest.raises(RecoveryLedgerError) as exc:
        load_encrypted_ledger(ledger_path, secret_key)
    assert "Unsupported recovery ledger protocol version" in str(exc.value)




def test_ledger_secret_not_embedded(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "secret_check.recovery_ledger.enc"
    orig_ledger = sample_ledger()

    save_encrypted_ledger(orig_ledger, ledger_path, secret_key)

    ciphertext_bytes = ledger_path.read_bytes()
    assert secret_key not in ciphertext_bytes

    loaded = load_encrypted_ledger(ledger_path, secret_key)
    json_text = loaded.to_json()
    assert secret_key.decode("ascii") not in json_text


def test_ledger_file_is_not_plaintext_json(tmp_path):
    secret_key = Fernet.generate_key()
    ledger_path = tmp_path / "not_plain.recovery_ledger.enc"
    orig_ledger = sample_ledger(run_id="super_secret_run_999")

    save_encrypted_ledger(orig_ledger, ledger_path, secret_key)
    raw_disk_content = ledger_path.read_bytes()

    # None of the sensitive metadata should be plain text on disk
    assert b"super_secret_run_999" not in raw_disk_content
    assert "张三".encode("utf-8") not in raw_disk_content
    assert "员工工号".encode("utf-8") not in raw_disk_content
    assert b"00001234" not in raw_disk_content
    assert b"protocol_version" not in raw_disk_content
    assert b"{" not in raw_disk_content[:10]  # Not raw JSON start


def test_recovery_ledger_module_is_integrity_locked():
    # Verify recovery_ledger.py is included in LOCKED_FILES
    locked_names = [p.name for p in tool.LOCKED_FILES]
    assert "recovery_ledger.py" in locked_names

    # Verify checksum can be computed for all locked files
    checksums = tool.current_checksums()
    assert "privacy_handler/recovery_ledger.py" in checksums
    assert len(checksums["privacy_handler/recovery_ledger.py"]) == 64


@pytest.mark.parametrize("suffix", [".csv", ".CSV", ".txt"])
def test_non_enc_recovery_rejected_before_decryption(tmp_path, monkeypatch, suffix):
    path = tmp_path / ("unsupported" + suffix)
    path.write_text("synthetic legacy CSV content", encoding="utf-8")
    def forbidden_key(*args, **kwargs):
        raise AssertionError("Unsupported input must be rejected before key access")
    monkeypatch.setattr(ledger_mod, "coerce_fernet", forbidden_key)
    with pytest.raises(RecoveryLedgerError, match="CSV recovery is not supported"):
        load_encrypted_ledger(path, b"unused")


def test_renamed_csv_enc_does_not_fall_back(tmp_path):
    path = tmp_path / "renamed.recovery_ledger.enc"
    path.write_text("sheet_name,column_header,masked_value,encrypted_original_value\n", encoding="utf-8")
    with pytest.raises(RecoveryLedgerError, match="Failed to decrypt recovery ledger"):
        load_encrypted_ledger(path, Fernet.generate_key())
