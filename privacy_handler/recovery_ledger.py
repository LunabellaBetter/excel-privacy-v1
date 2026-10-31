"""Locked local recovery ledger data models, encryption, and persistence foundation.

Provides structured manifest storage for V1 AI round-trip privacy protocol,
whole-file Fernet encryption, atomic file I/O, and type preservation for V1 encrypted ledgers.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
import datetime as dt
import json
import os
from pathlib import Path
import uuid
from typing import Any

try:
    from cryptography.fernet import Fernet, InvalidToken
except Exception:  # pragma: no cover - dependency availability is environment-specific.
    Fernet = None
    InvalidToken = Exception


CURRENT_PROTOCOL_VERSION = "1.0"
SUPPORTED_PROTOCOL_VERSIONS = {CURRENT_PROTOCOL_VERSION}


class RecoveryLedgerError(Exception):
    """Safe error for recovery ledger operations. Must not expose secrets."""


def require_fernet() -> Any:
    if Fernet is None:
        raise RecoveryLedgerError("cryptography package is required for encrypted recovery ledger operations")
    return Fernet


def coerce_fernet(key: Any) -> Any:
    fernet_cls = require_fernet()
    if isinstance(key, fernet_cls):
        return key
    if isinstance(key, (bytes, str)):
        raw = key.encode("ascii") if isinstance(key, str) else key
        return fernet_cls(raw)
    raise RecoveryLedgerError("Invalid secret key provider for recovery ledger")


def atomic_write_bytes(target_path: Path, data: bytes) -> None:
    """Atomically write binary data to target path using a temporary sibling file."""
    target = Path(target_path).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_name = f".{target.name}.tmp_{uuid.uuid4().hex[:8]}"
    tmp_path = target.with_name(tmp_name)
    try:
        with open(tmp_path, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        tmp_path.replace(target)
    except Exception as exc:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass
        raise exc


def encode_typed_value(value: Any) -> tuple[str, Any]:
    """Serialize Python native value into (type_tag, json_serializable_value)."""
    if isinstance(value, dt.datetime):
        return "datetime", value.isoformat()
    if isinstance(value, dt.date):
        return "date", value.isoformat()
    if isinstance(value, bool):
        return "bool", bool(value)
    if isinstance(value, int):
        return "int", int(value)
    if isinstance(value, float):
        return "float", float(value)
    return "str", str(value)


def decode_typed_value(type_tag: str, raw_value: Any) -> Any:
    """Deserialize (type_tag, raw_value) back into Python native type."""
    if type_tag == "datetime":
        return dt.datetime.fromisoformat(str(raw_value))
    if type_tag == "date":
        return dt.date.fromisoformat(str(raw_value))
    if type_tag == "bool":
        return bool(raw_value)
    if type_tag == "int":
        return int(raw_value)
    if type_tag == "float":
        return float(raw_value)
    return str(raw_value)


@dataclass
class LedgerRunInfo:
    run_id: str
    created_at: str
    workbook_identity: str
    source_extension: str = ".xlsx"

    def validate(self) -> None:
        if not self.run_id or not isinstance(self.run_id, str):
            raise RecoveryLedgerError("Invalid run.run_id")
        if not self.created_at or not isinstance(self.created_at, str):
            raise RecoveryLedgerError("Invalid run.created_at")
        if not self.workbook_identity or not isinstance(self.workbook_identity, str):
            raise RecoveryLedgerError("Invalid run.workbook_identity")


@dataclass
class LedgerConfig:
    confirmed_key_field: str | None = None
    numeric_mode: str | None = None
    sheet_scopes: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class ProtectedField:
    logical_field: str
    protection_type: str = "text_token"
    protected_field_id: str | None = None
    sheet_name: str | None = None
    source_header_row: int | None = None
    source_column_letter: str | None = None
    raw_header: str | None = None
    display_header: str | None = None
    numeric_analysis_group: str | None = None

    def validate(self) -> None:
        if not self.logical_field or not isinstance(self.logical_field, str):
            raise RecoveryLedgerError("Invalid protected_field.logical_field")


@dataclass
class LedgerEntry:
    sheet: str
    logical_field: str
    key_masked_value: str | None
    original_type: str
    original_value: Any
    masked_type: str
    masked_value: Any
    protection_type: str = "text_token"
    protected_field_id: str | None = None
    record_identity_masked: list[dict[str, Any]] | None = None

    def validate(self) -> None:
        if not self.sheet or not isinstance(self.sheet, str):
            raise RecoveryLedgerError("Invalid entry.sheet")
        if not self.logical_field or not isinstance(self.logical_field, str):
            raise RecoveryLedgerError("Invalid entry.logical_field")
        if not self.original_type or not isinstance(self.original_type, str):
            raise RecoveryLedgerError("Invalid entry.original_type")
        if not self.masked_type or not isinstance(self.masked_type, str):
            raise RecoveryLedgerError("Invalid entry.masked_type")

    def get_decoded_original_value(self) -> Any:
        return decode_typed_value(self.original_type, self.original_value)

    def get_decoded_masked_value(self) -> Any:
        return decode_typed_value(self.masked_type, self.masked_value)


@dataclass
class RecoveryLedger:
    run: LedgerRunInfo
    configuration: LedgerConfig = field(default_factory=LedgerConfig)
    protected_fields: list[ProtectedField] = field(default_factory=list)
    entries: list[LedgerEntry] = field(default_factory=list)
    numeric_groups: list[dict[str, Any]] = field(default_factory=list)
    protocol_version: str = CURRENT_PROTOCOL_VERSION

    def validate(self) -> None:
        if self.protocol_version not in SUPPORTED_PROTOCOL_VERSIONS:
            raise RecoveryLedgerError(
                f"Unsupported recovery ledger protocol version: '{self.protocol_version}'. "
                f"Supported versions: {sorted(SUPPORTED_PROTOCOL_VERSIONS)}"
            )
        self.run.validate()
        for pf in self.protected_fields:
            pf.validate()
        for entry in self.entries:
            entry.validate()

    def to_dict(self) -> dict[str, Any]:
        self.validate()
        return {
            "protocol_version": self.protocol_version,
            "run": asdict(self.run),
            "configuration": asdict(self.configuration),
            "protected_fields": [asdict(pf) for pf in self.protected_fields],
            "entries": [asdict(entry) for entry in self.entries],
            "numeric_groups": list(self.numeric_groups),
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> RecoveryLedger:
        if not isinstance(data, dict):
            raise RecoveryLedgerError("Recovery ledger root must be a JSON object")

        version = data.get("protocol_version")
        if version not in SUPPORTED_PROTOCOL_VERSIONS:
            raise RecoveryLedgerError(
                f"Unsupported recovery ledger protocol version: '{version}'. "
                f"Supported versions: {sorted(SUPPORTED_PROTOCOL_VERSIONS)}"
            )

        run_data = data.get("run")
        if not isinstance(run_data, dict):
            raise RecoveryLedgerError("Missing or invalid 'run' section in recovery ledger")
        run = LedgerRunInfo(
            run_id=run_data.get("run_id", ""),
            created_at=run_data.get("created_at", ""),
            workbook_identity=run_data.get("workbook_identity", ""),
            source_extension=run_data.get("source_extension", ".xlsx"),
        )

        cfg_data = data.get("configuration") or {}
        configuration = LedgerConfig(
            confirmed_key_field=cfg_data.get("confirmed_key_field"),
            numeric_mode=cfg_data.get("numeric_mode"),
            sheet_scopes=cfg_data.get("sheet_scopes", []),
        )

        pf_list: list[ProtectedField] = []
        for pf_item in data.get("protected_fields", []):
            if isinstance(pf_item, dict):
                pf_list.append(
                    ProtectedField(
                        logical_field=pf_item.get("logical_field", ""),
                        protection_type=pf_item.get("protection_type", "text_token"),
                        protected_field_id=pf_item.get("protected_field_id"),
                        sheet_name=pf_item.get("sheet_name"),
                        source_header_row=pf_item.get("source_header_row"),
                        source_column_letter=pf_item.get("source_column_letter"),
                        raw_header=pf_item.get("raw_header"),
                        display_header=pf_item.get("display_header"),
                        numeric_analysis_group=pf_item.get("numeric_analysis_group"),
                    )
                )

        entries_list: list[LedgerEntry] = []
        raw_entries = data.get("entries")
        if not isinstance(raw_entries, list):
            raise RecoveryLedgerError("Recovery ledger 'entries' must be a list")

        for item in raw_entries:
            if not isinstance(item, dict):
                raise RecoveryLedgerError("Each ledger entry must be an object")
            entries_list.append(
                LedgerEntry(
                    sheet=item.get("sheet", ""),
                    logical_field=item.get("logical_field", ""),
                    key_masked_value=item.get("key_masked_value"),
                    original_type=item.get("original_type", "str"),
                    original_value=item.get("original_value"),
                    masked_type=item.get("masked_type", "str"),
                    masked_value=item.get("masked_value"),
                    protection_type=item.get("protection_type", "text_token"),
                    protected_field_id=item.get("protected_field_id"),
                    record_identity_masked=item.get("record_identity_masked"),
                )
            )

        num_groups = data.get("numeric_groups", [])
        if not isinstance(num_groups, list):
            num_groups = []

        ledger = cls(
            protocol_version=version,
            run=run,
            configuration=configuration,
            protected_fields=pf_list,
            entries=entries_list,
            numeric_groups=num_groups,
        )
        ledger.validate()
        return ledger

    @classmethod
    def from_json(cls, json_str: str) -> RecoveryLedger:
        try:
            data = json.loads(json_str)
        except Exception as exc:
            raise RecoveryLedgerError("Corrupted recovery ledger: invalid JSON") from exc
        return cls.from_dict(data)

    def to_column_mappings(self) -> dict[tuple[str, str], dict[Any, Any]]:
        """Extract mapping dictionary matching (sheet, logical_field) -> {masked_val: decoded_original_val}."""
        mappings: dict[tuple[str, str], dict[Any, Any]] = {}
        for entry in self.entries:
            k = (entry.sheet, entry.logical_field)
            masked_val = entry.get_decoded_masked_value()
            orig_val = entry.get_decoded_original_value()
            mappings.setdefault(k, {})[masked_val] = orig_val
        return mappings


def save_encrypted_ledger(ledger: RecoveryLedger, target_path: Path, secret_key: Any) -> None:
    """Serialize and encrypt whole ledger with Fernet, saving atomically to disk."""
    fernet = coerce_fernet(secret_key)
    json_text = ledger.to_json()

    # Safety check: secret key must never be embedded in the plaintext payload
    secret_bytes = getattr(fernet, "_signing_key", b"")
    if secret_bytes and secret_bytes in json_text.encode("utf-8"):
        raise RecoveryLedgerError("Safety invariant violated: secret key found inside ledger text")

    ciphertext = fernet.encrypt(json_text.encode("utf-8"))
    atomic_write_bytes(target_path, ciphertext)


def load_encrypted_ledger(target_path: Path, secret_key: Any) -> RecoveryLedger:
    """Read encrypted ledger bytes, decrypt with Fernet, validate, and return RecoveryLedger."""
    path = Path(target_path).resolve()
    if path.suffix.lower() != ".enc":
        raise RecoveryLedgerError("V1 requires an encrypted .enc recovery ledger; CSV recovery is not supported")
    if not path.exists():
        raise RecoveryLedgerError(f"Recovery ledger not found: {path.name}")

    fernet = coerce_fernet(secret_key)
    ciphertext = path.read_bytes()

    try:
        decrypted_bytes = fernet.decrypt(ciphertext)
    except InvalidToken as exc:
        raise RecoveryLedgerError(
            "Failed to decrypt recovery ledger: secret key may be incorrect or file is corrupted"
        ) from exc
    except Exception as exc:
        raise RecoveryLedgerError("Failed to decrypt recovery ledger") from exc

    try:
        json_str = decrypted_bytes.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise RecoveryLedgerError("Decrypted recovery ledger is not valid UTF-8 text") from exc

    return RecoveryLedger.from_json(json_str)

