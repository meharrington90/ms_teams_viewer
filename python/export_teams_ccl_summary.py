#!/usr/bin/env python3

import argparse
import io
import json
from contextlib import ExitStack, contextmanager, redirect_stderr, redirect_stdout
from pathlib import Path

from dump_teams_indexeddb_ccl import iter_wrapper_databases, load_ccl, resolve_teams_source


def safe_get_store(db, store_name: str):
    try:
        return db[store_name]
    except Exception:
        return None


@contextmanager
def decode_output_context(show_decode_errors: bool):
    if show_decode_errors:
        yield
        return

    sink = io.StringIO()
    with ExitStack() as stack:
        stack.enter_context(redirect_stdout(sink))
        stack.enter_context(redirect_stderr(sink))
        yield


def collect_store_records(store, limit: int | None = None, show_decode_errors: bool = True) -> list[dict]:
    records = []
    with decode_output_context(show_decode_errors):
        for index, record in enumerate(store.iterate_records(errors_to_stdout=True)):
            records.append(record.value)
            if limit is not None and index + 1 >= limit:
                break
    return records


def build_summary(root: Path | str | None = None, show_decode_errors: bool = True) -> dict:
    source = resolve_teams_source(root)
    leveldb_path = source.leveldb_path
    blob_path = source.blob_path
    ccl = load_ccl()
    wrapper = ccl.WrappedIndexDB(str(leveldb_path), str(blob_path) if blob_path.exists() else None)

    summary = {
        "backend": "ccl_chromium_reader.wrapper",
        "platform": source.platform_name,
        "search_root": str(source.search_root),
        "profile_root": str(source.profile_root),
        "leveldb_path": str(leveldb_path),
        "blob_path": str(blob_path) if blob_path.exists() else None,
        "local_storage_dir": str(source.local_storage_dir) if source.local_storage_dir else None,
        "discovery_method": source.discovery_method,
        "databases_total": 0,
        "key_stores": {},
        "samples": {},
    }

    targets = {
        "people": ("Teams:substrate-suggestions-manager", "people"),
        "conversations": ("Teams:conversation-manager", "conversations"),
        "replychains": ("Teams:replychain-manager", "replychains"),
        "call_history": ("Teams:call-history-manager", "call-history"),
        "threads_internal": ("Teams:messaging-slice-manager", "threads-internal-items"),
        "drafts_internal": ("Teams:messaging-slice-manager", "drafts-internal-items"),
        "system_messages": ("Teams:channel-info-pane-manager", "system-messages-store"),
    }

    dbs = iter_wrapper_databases(wrapper)
    summary["databases_total"] = len(dbs)

    for label, (db_prefix, store_name) in targets.items():
        matched = None
        for _, db in dbs:
            db_name = getattr(db, "name", "") or ""
            if db_name.startswith(db_prefix):
                store = safe_get_store(db, store_name)
                if store is None:
                    continue
                matched = store
                break

        if matched is None:
            summary["key_stores"][label] = {"found": False}
            continue

        records = collect_store_records(matched, show_decode_errors=show_decode_errors)
        summary["key_stores"][label] = {"found": True, "record_count": len(records)}
        if records:
            sample = records[0]
            if label == "people":
                summary["samples"][label] = {
                    "DisplayName": sample.get("DisplayName"),
                    "EmailAddresses": sample.get("EmailAddresses"),
                    "ExternalDirectoryObjectId": sample.get("ExternalDirectoryObjectId"),
                }
            elif label == "conversations":
                summary["samples"][label] = {
                    "id": sample.get("id"),
                    "type": sample.get("type"),
                    "title": sample.get("title"),
                    "member_count": len(sample.get("members", [])),
                }
            elif label == "replychains":
                message_map = sample.get("messageMap") or {}
                first_message = next(iter(message_map.values()), {})
                summary["samples"][label] = {
                    "conversationId": sample.get("conversationId"),
                    "replyChainId": sample.get("replyChainId"),
                    "messageMap_count": len(message_map),
                    "first_message": {
                        "id": first_message.get("id"),
                        "imDisplayName": first_message.get("imDisplayName"),
                        "messageType": first_message.get("messageType"),
                        "content_preview": (first_message.get("content") or "")[:300],
                    },
                }
            elif label == "call_history":
                summary["samples"][label] = {
                    "id": sample.get("id"),
                    "callDirection": sample.get("callDirection"),
                    "callType": sample.get("callType"),
                    "callState": sample.get("callState"),
                    "originator": (sample.get("originatorParticipant") or {}).get("displayName"),
                    "target": (sample.get("targetParticipant") or {}).get("displayName"),
                }
            else:
                summary["samples"][label] = sample

    # Derive aggregate message count directly from the structured replychains store.
    replychain_store = summary["key_stores"].get("replychains") or {}
    if replychain_store.get("found"):
        total_messages = 0
        message_threads = 0
        for _, db in dbs:
            db_name = getattr(db, "name", "") or ""
            if not db_name.startswith("Teams:replychain-manager"):
                continue
            store = safe_get_store(db, "replychains")
            if store is None:
                continue
            with decode_output_context(show_decode_errors):
                for record in store.iterate_records(errors_to_stdout=True):
                    value = record.value or {}
                    message_map = value.get("messageMap") or {}
                    if message_map:
                        message_threads += 1
                        total_messages += len(message_map)
            break
        summary["key_stores"]["replychains"]["messages_total"] = total_messages
        summary["key_stores"]["replychains"]["threads_with_messages"] = message_threads

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize key Teams IndexedDB stores using ccl_chromium_reader.")
    parser.add_argument("--root", help="Teams source root, profile directory, copied evidence root, or IndexedDB directory.")
    parser.add_argument("--output", default="teams_ccl_summary.json", help="Path to write JSON output.")
    args = parser.parse_args()

    payload = build_summary(Path(args.root).resolve() if args.root else None)
    Path(args.output).resolve().write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
