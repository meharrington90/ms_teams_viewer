#!/usr/bin/env python3

import argparse
import csv
import json
import re
from pathlib import Path


SAFE_RE = re.compile(r"[^A-Za-z0-9._-]+")


def load_export(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def safe_name(value: str) -> str:
    value = SAFE_RE.sub("_", value or "")
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "untitled"


def sort_key(message: dict) -> tuple[bool, str, str]:
    return (message.get("timestamp") is None, message.get("timestamp") or "", message.get("id") or "")


def thread_filename(thread: dict) -> str:
    label = thread.get("label") or thread.get("id") or "thread"
    return f"{safe_name(thread.get('category') or 'thread')}__{safe_name(label)}__{safe_name(thread['id'])[:12]}.csv"


def write_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def export_conversations(export_data: dict, output_dir: Path) -> list[dict]:
    messages_dir = output_dir / "conversations"
    ensure_dir(messages_dir)

    manifest = []
    all_rows = []
    fields = [
        "thread_id",
        "thread_label",
        "thread_category",
        "metadata_quality",
        "message_id",
        "client_message_id",
        "timestamp",
        "sender_display_name",
        "sender_id",
        "message_type",
        "content_type",
        "quality",
        "content_text",
        "source",
    ]

    for thread in export_data.get("threads", []):
        rows = []
        for message in sorted(thread.get("messages", []), key=sort_key):
            row = {
                "thread_id": thread["id"],
                "thread_label": thread.get("label"),
                "thread_category": thread.get("category"),
                "metadata_quality": thread.get("metadata_quality"),
                "message_id": message.get("id"),
                "client_message_id": message.get("client_message_id"),
                "timestamp": message.get("timestamp"),
                "sender_display_name": message.get("sender_display_name"),
                "sender_id": message.get("sender_id"),
                "message_type": message.get("message_type"),
                "content_type": message.get("content_type"),
                "quality": message.get("quality"),
                "content_text": message.get("content_text"),
                "source": message.get("source"),
            }
            rows.append(row)
            all_rows.append(row)

        filename = thread_filename(thread)
        if rows:
            write_csv(messages_dir / filename, fields, rows)

        manifest.append(
            {
                "thread_id": thread["id"],
                "thread_label": thread.get("label"),
                "thread_category": thread.get("category"),
                "metadata_quality": thread.get("metadata_quality"),
                "message_count": thread.get("message_count"),
                "participant_count": len(thread.get("participants", [])),
                "participants": "; ".join(thread.get("participants", [])),
                "csv": str((messages_dir / filename).relative_to(output_dir)) if rows else "",
            }
        )

    write_csv(output_dir / "conversations_all_flat.csv", fields, sorted(all_rows, key=lambda row: (row["thread_label"] or "", row["timestamp"] or "", row["message_id"] or "")))
    write_csv(
        output_dir / "conversation_manifest.csv",
        ["thread_id", "thread_label", "thread_category", "metadata_quality", "message_count", "participant_count", "participants", "csv"],
        manifest,
    )
    return manifest


def export_calls(export_data: dict, output_dir: Path) -> None:
    fields = [
        "call_id",
        "shared_correlation_id",
        "start_time",
        "connect_time",
        "end_time",
        "direction",
        "call_type",
        "call_state",
        "originator_display_name",
        "originator_id",
        "originator_endpoint",
        "originator_phone_number",
        "target_display_name",
        "target_id",
        "target_endpoint",
        "target_phone_number",
        "participant_display_names",
        "participant_ids",
        "conversation_id",
        "group_chat_thread_id",
        "meeting_subject",
        "meeting_start_time",
        "meeting_end_time",
        "meeting_series_kind",
        "quality",
        "summary_text",
        "source",
    ]
    rows = []
    for call in export_data.get("calls", []):
        row = {field: call.get(field) for field in fields}
        row["participant_display_names"] = "; ".join(call.get("participant_display_names", []))
        row["participant_ids"] = "; ".join(call.get("participant_ids", []))
        rows.append(row)
    write_csv(output_dir / "call_history.csv", fields, rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Export the CCL canonical Teams JSON to per-conversation CSVs.")
    parser.add_argument("--input", default="teams_ccl_canonical_v1.json", help="Path to the canonical CCL JSON.")
    parser.add_argument("--output-dir", default="teams_ccl_csv_v1", help="Directory for CSV exports.")
    args = parser.parse_args()

    export_data = load_export(Path(args.input).resolve())
    output_dir = Path(args.output_dir).resolve()
    ensure_dir(output_dir)

    export_conversations(export_data, output_dir)
    export_calls(export_data, output_dir)
    (output_dir / "export_summary.json").write_text(json.dumps(export_data.get("summary") or {}, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
