#!/usr/bin/env python3
"""
Extract metadata, body content, and attachments from Outlook .msg files.

Dependencies:
  pip install extract-msg
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

try:
    import extract_msg
except ImportError:
    print(
        "Missing dependency: extract-msg\nInstall with: pip install extract-msg",
        file=sys.stderr,
    )
    sys.exit(1)

try:
    import olefile
except ImportError:
    olefile = None


IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".webp",
    ".heic",
}

OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def get_attr(obj: Any, *names: str) -> Any:
    for name in names:
        if hasattr(obj, name):
            value = getattr(obj, name)
            if value is not None:
                return value
    return None


def json_safe(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, bytes):
        return f"<{len(value)} bytes>"
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, (list, tuple)):
        return [json_safe(v) for v in value]
    if isinstance(value, dict):
        return {str(k): json_safe(v) for k, v in value.items()}
    return str(value)


def sanitize_name(name: str, fallback: str = "item") -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", name or "")
    cleaned = cleaned.strip().strip(".")
    return cleaned or fallback


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def coerce_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, bytes):
        for encoding in ("utf-8", "utf-16", "latin-1"):
            try:
                return value.decode(encoding)
            except UnicodeDecodeError:
                continue
        return value.decode("utf-8", errors="replace")
    return str(value)


def detect_zip_password_protection(zip_path: Path) -> dict[str, Any]:
    details: dict[str, Any] = {
        "is_zip": True,
        "password_protected": False,
        "entry_count": 0,
        "error": None,
    }
    try:
        with zipfile.ZipFile(zip_path) as zf:
            infos = zf.infolist()
            details["entry_count"] = len(infos)
            details["password_protected"] = any(info.flag_bits & 0x1 for info in infos)
    except (zipfile.BadZipFile, OSError) as exc:
        details["error"] = str(exc)
    return details


def flatten_candidates(value: Any) -> Iterable[str]:
    if isinstance(value, (str, Path)):
        yield str(value)
        return
    if isinstance(value, dict):
        for v in value.values():
            yield from flatten_candidates(v)
        return
    if isinstance(value, (list, tuple, set)):
        for v in value:
            yield from flatten_candidates(v)


def parse_saved_path(save_result: Any, attachment_dir: Path) -> Path | None:
    for candidate in flatten_candidates(save_result):
        p = Path(candidate)
        if not p.is_absolute():
            p = attachment_dir / p
        if p.exists():
            return p
    return None


def attachment_filename(attachment: Any, index: int) -> str:
    candidate = get_attr(
        attachment,
        "longFilename",
        "filename",
        "name",
        "shortFilename",
    )
    if candidate:
        return sanitize_name(str(candidate), fallback=f"attachment_{index}")
    return f"attachment_{index}"


def save_attachment(attachment: Any, index: int, attachment_dir: Path) -> tuple[Path | None, str | None]:
    save_error: str | None = None

    if hasattr(attachment, "save"):
        for kwargs in (
            {"customPath": str(attachment_dir), "skipHidden": False},
            {"customPath": str(attachment_dir)},
        ):
            try:
                before = {p.name for p in attachment_dir.iterdir() if p.is_file()}
                result = attachment.save(**kwargs)
                path = parse_saved_path(result, attachment_dir)
                if path is not None:
                    return path, None
                after = [p for p in attachment_dir.iterdir() if p.is_file() and p.name not in before]
                if after:
                    newest = max(after, key=lambda p: p.stat().st_mtime)
                    return newest, None
            except TypeError:
                continue
            except Exception as exc:  # noqa: BLE001
                save_error = str(exc)
                break

    data = get_attr(attachment, "data")
    if isinstance(data, (bytes, bytearray)):
        filename = attachment_filename(attachment, index)
        target = unique_path(attachment_dir / filename)
        target.write_bytes(bytes(data))
        return target, save_error

    return None, save_error


def build_message_metadata(msg: Any, source_file: Path) -> dict[str, Any]:
    metadata = {
        "source_file": str(source_file),
        "extracted_at_utc": datetime.now(timezone.utc).isoformat(),
        "subject": json_safe(get_attr(msg, "subject")),
        "sender_name": json_safe(get_attr(msg, "sender")),
        "sender_email": json_safe(get_attr(msg, "senderEmail", "sender_email")),
        "to": json_safe(get_attr(msg, "to")),
        "cc": json_safe(get_attr(msg, "cc")),
        "bcc": json_safe(get_attr(msg, "bcc")),
        "date": json_safe(get_attr(msg, "date")),
        "message_id": json_safe(get_attr(msg, "messageId", "message_id")),
        "in_reply_to": json_safe(get_attr(msg, "inReplyTo", "in_reply_to")),
        "reply_to": json_safe(get_attr(msg, "replyTo", "reply_to")),
        "importance": json_safe(get_attr(msg, "importance")),
        "priority": json_safe(get_attr(msg, "priority")),
    }
    return metadata


def write_body_files(msg: Any, output_dir: Path) -> dict[str, str]:
    outputs: dict[str, str] = {}

    body_text = coerce_text(get_attr(msg, "body"))
    if body_text:
        text_path = output_dir / "body.txt"
        text_path.write_text(body_text, encoding="utf-8")
        outputs["plain_text"] = str(text_path)

    html_body = get_attr(msg, "htmlBody")
    if html_body:
        html_path = output_dir / "body.html"
        if isinstance(html_body, bytes):
            html_path.write_bytes(html_body)
        else:
            html_path.write_text(str(html_body), encoding="utf-8")
        outputs["html"] = str(html_path)

    rtf_body = get_attr(msg, "rtfBody")
    if rtf_body:
        rtf_path = output_dir / "body.rtf"
        if isinstance(rtf_body, bytes):
            rtf_path.write_bytes(rtf_body)
        else:
            rtf_path.write_text(str(rtf_body), encoding="utf-8")
        outputs["rtf"] = str(rtf_path)

    headers = coerce_text(get_attr(msg, "header"))
    if headers:
        headers_path = output_dir / "headers.txt"
        headers_path.write_text(headers, encoding="utf-8")
        outputs["headers"] = str(headers_path)

    return outputs


def is_embedded_msg_file(path: Path | None) -> bool:
    return bool(path and path.suffix.lower() == ".msg")


def validate_msg_container(path: Path) -> tuple[bool, str | None]:
    if not path.exists() or not path.is_file():
        return False, "File not found"

    try:
        with path.open("rb") as handle:
            signature = handle.read(8)
    except OSError as exc:
        return False, f"Unable to read file header: {exc}"

    if signature != OLE_SIGNATURE:
        return False, "Not an OLE MSG file (invalid signature)"

    if olefile is None:
        return True, None

    try:
        if not olefile.isOleFile(str(path)):
            return False, "Not an OLE file"
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)

    return True, None


def process_msg_file(
    msg_path: Path,
    output_dir: Path,
    recursive: bool,
    visited: set[Path],
    logger: logging.Logger,
) -> dict[str, Any]:
    msg_path = msg_path.resolve()
    if msg_path in visited:
        return {
            "source_file": str(msg_path),
            "status": "skipped_already_processed",
            "output_dir": str(output_dir),
        }

    visited.add(msg_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    attachments_dir = output_dir / "attachments"
    attachments_dir.mkdir(parents=True, exist_ok=True)

    msg = None
    result: dict[str, Any] = {
        "source_file": str(msg_path),
        "status": "ok",
        "output_dir": str(output_dir),
        "body_files": {},
        "attachments": [],
        "embedded_messages": [],
    }

    is_valid_msg, invalid_reason = validate_msg_container(msg_path)
    if not is_valid_msg:
        result["status"] = "skipped_invalid_msg"
        result["error"] = invalid_reason
        logger.warning("Skipping invalid .msg file %s: %s", msg_path, invalid_reason)
        return result

    try:
        msg = extract_msg.Message(str(msg_path))
        metadata = build_message_metadata(msg, msg_path)
        body_files = write_body_files(msg, output_dir)

        attachment_records: list[dict[str, Any]] = []
        attachments = list(get_attr(msg, "attachments") or [])
        for idx, attachment in enumerate(attachments, start=1):
            record: dict[str, Any] = {
                "index": idx,
                "filename": attachment_filename(attachment, idx),
                "long_filename": json_safe(get_attr(attachment, "longFilename")),
                "short_filename": json_safe(get_attr(attachment, "shortFilename")),
                "content_id": json_safe(get_attr(attachment, "contentId", "cid")),
                "mime_type": json_safe(get_attr(attachment, "mimetype", "mimeType")),
                "hidden": bool(get_attr(attachment, "hidden", "isHidden") or False),
                "save_error": None,
                "saved_path": None,
                "saved_size_bytes": None,
                "zip_details": None,
            }

            saved_path, save_error = save_attachment(attachment, idx, attachments_dir)
            record["save_error"] = save_error
            if saved_path and saved_path.exists():
                record["saved_path"] = str(saved_path)
                record["saved_size_bytes"] = saved_path.stat().st_size
                suffix = saved_path.suffix.lower()
                record["is_image"] = suffix in IMAGE_EXTENSIONS or str(record["mime_type"]).startswith("image/")

                if suffix == ".zip":
                    record["zip_details"] = detect_zip_password_protection(saved_path)
                    if record["zip_details"]["password_protected"]:
                        logger.warning("Password-protected zip detected: %s", saved_path)

                if recursive and is_embedded_msg_file(saved_path):
                    embedded_valid, embedded_reason = validate_msg_container(saved_path)
                    if embedded_valid:
                        nested_base = output_dir / "embedded_messages"
                        nested_name = sanitize_name(saved_path.stem, fallback=f"embedded_{idx}")
                        nested_dir = unique_path(nested_base / nested_name)
                        nested_result = process_msg_file(saved_path, nested_dir, recursive, visited, logger)
                        result["embedded_messages"].append(nested_result)
                    else:
                        record["embedded_msg_status"] = "skipped_invalid_msg"
                        record["embedded_msg_reason"] = embedded_reason
                        logger.warning(
                            "Skipping embedded .msg attachment %s: %s",
                            saved_path,
                            embedded_reason,
                        )

            attachment_records.append(record)

        metadata["attachment_count"] = len(attachment_records)
        metadata["attachments"] = attachment_records
        metadata["body_files"] = body_files

        metadata_file = output_dir / "metadata.json"
        metadata_file.write_text(json.dumps(json_safe(metadata), indent=2), encoding="utf-8")

        result["body_files"] = body_files
        result["attachments"] = attachment_records
        result["metadata_file"] = str(metadata_file)

    except Exception as exc:  # noqa: BLE001
        error_text = str(exc)
        if "property stream" in error_text.lower():
            result["status"] = "skipped_invalid_msg"
            result["error"] = "Unsupported .msg structure (missing property stream)"
            logger.warning("Skipping unsupported .msg file %s: %s", msg_path, error_text)
        else:
            result["status"] = "error"
            result["error"] = error_text
            logger.exception("Failed to process %s", msg_path)
    finally:
        if msg is not None and hasattr(msg, "close"):
            try:
                msg.close()
            except Exception:  # noqa: BLE001
                pass

    return result


def discover_msg_files(input_path: Path, exclude_roots: Iterable[Path] | None = None) -> list[Path]:
    excludes = [p.resolve() for p in (exclude_roots or [])]

    def is_excluded(path: Path) -> bool:
        candidate = path.resolve()
        return any(candidate.is_relative_to(root) for root in excludes)

    if input_path.is_file():
        if input_path.suffix.lower() == ".msg" and not is_excluded(input_path):
            return [input_path]
        return []
    if input_path.is_dir():
        return sorted(
            path
            for path in input_path.rglob("*")
            if path.suffix.lower() == ".msg" and not is_excluded(path)
        )
    return []


def make_case_output_dir(output_root: Path, source_msg: Path) -> Path:
    label = sanitize_name(source_msg.stem, fallback="email")
    return unique_path(output_root / label)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract metadata, email body content, and attachments from one .msg file "
            "or all .msg files in a directory."
        )
    )
    parser.add_argument("input", help="Path to a .msg file or folder containing .msg files")
    parser.add_argument(
        "-o",
        "--output",
        default="msg_extraction_output",
        help="Output directory (default: ./msg_extraction_output)",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Do not recursively process embedded .msg attachments",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )
    logger = logging.getLogger("msg-extractor")

    input_path = Path(args.input).expanduser().resolve()
    output_root = Path(args.output).expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    msg_files = discover_msg_files(input_path, exclude_roots=[output_root])
    if not msg_files:
        logger.error("No .msg files found in: %s", input_path)
        return 1

    visited: set[Path] = set()
    summary: list[dict[str, Any]] = []

    for msg_file in msg_files:
        case_dir = make_case_output_dir(output_root, msg_file)
        logger.info("Processing: %s", msg_file)
        result = process_msg_file(
            msg_path=msg_file,
            output_dir=case_dir,
            recursive=not args.no_recursive,
            visited=visited,
            logger=logger,
        )
        summary.append(result)

    summary_path = output_root / "summary.json"
    summary_path.write_text(json.dumps(json_safe(summary), indent=2), encoding="utf-8")
    logger.info("Done. Output written to: %s", output_root)
    logger.info("Summary: %s", summary_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
