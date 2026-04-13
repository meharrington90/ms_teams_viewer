#!/usr/bin/env python3

import argparse
import io
import json
from contextlib import ExitStack, contextmanager, redirect_stderr, redirect_stdout
from datetime import datetime, timezone
from pathlib import Path

from dump_teams_indexeddb_ccl import iter_wrapper_databases, load_ccl, resolve_teams_source
from teams_ccl_common import (
    GUID_RE,
    clean_text,
    html_to_text,
    init_thread_record,
    normalize_thread_id,
    parse_profile,
)


def safe_text(value) -> str | None:
    if isinstance(value, str):
        return clean_text(value)
    if isinstance(value, (bytes, bytearray)):
        return clean_text(bytes(value).decode("utf-8", "ignore"))
    return None


def to_iso_from_millis(value) -> str | None:
    if value in {None, "", "<Undefined>"}:
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    try:
        return datetime.fromtimestamp(number / 1000, tz=timezone.utc).isoformat()
    except (ValueError, OSError):
        return None


def extract_guid(value: str | None) -> str | None:
    if not value:
        return None
    match = GUID_RE.search(value)
    if not match:
        return None
    return match.group(0).lower()


def classify_quality(message_type: str | None, content_text: str | None) -> str:
    if message_type and message_type.startswith(("ThreadActivity/", "Event/")):
        return "event"
    if not content_text:
        return "residual"
    if len(content_text.strip()) <= 1:
        return "residual"
    return "curated"


def flatten_members(members: list[dict], guid_to_name: dict[str, str]) -> tuple[list[str], list[str]]:
    participant_ids = []
    participants = []
    for member in members or []:
        raw_id = member.get("id")
        guid = extract_guid(raw_id)
        if guid and guid not in participant_ids:
            participant_ids.append(guid)
        display_name = None
        name_hint = member.get("nameHint") or {}
        if isinstance(name_hint, dict):
            display_name = clean_text(name_hint.get("displayName"))
        if not display_name and guid:
            display_name = guid_to_name.get(guid)
        if display_name and display_name not in participants:
            participants.append(display_name)
    return participant_ids, participants


def merge_thread_participant(thread: dict, participant_id: str | None = None, participant_name: str | None = None) -> None:
    normalized_id = extract_guid(participant_id)
    cleaned_name = clean_text(participant_name)

    if normalized_id and normalized_id not in thread["participant_ids"]:
        thread["participant_ids"].append(normalized_id)
    if cleaned_name and cleaned_name not in thread["participants"]:
        thread["participants"].append(cleaned_name)


def parse_embedded_json(value):
    if isinstance(value, dict):
        return value
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            return None
        if isinstance(parsed, dict):
            return parsed
    return None


def is_placeholder_label(label: str | None, thread_id: str) -> bool:
    cleaned = clean_text(label)
    if not cleaned:
        return True
    if cleaned == thread_id:
        return True
    return cleaned.startswith(("19:", "48:")) or cleaned.endswith(("@thread.v2", "@thread.skype", "@unq.gbl.spaces"))


def build_participant_label(participants: list[str], current_user_name: str | None) -> str | None:
    ordered = []
    for name in participants or []:
        cleaned = clean_text(name)
        if cleaned and cleaned not in ordered:
            ordered.append(cleaned)
    others = [name for name in ordered if name != current_user_name]
    display = others or ordered
    if not display:
        return None
    if len(display) <= 4:
        return " / ".join(display)
    return f'{" / ".join(display[:4])} + {len(display) - 4} more'


def get_target_store(wrapper, db_prefix: str, store_name: str):
    for _, db in iter_wrapper_databases(wrapper):
        db_name = getattr(db, "name", "") or ""
        if not db_name.startswith(db_prefix):
            continue
        try:
            return db[store_name]
        except Exception:
            continue
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


def build_export(root: Path | str | None = None, show_decode_errors: bool = True) -> dict:
    source_info = resolve_teams_source(root)
    ccl = load_ccl()
    leveldb_path = source_info.leveldb_path
    blob_path = source_info.blob_path
    wrapper = ccl.WrappedIndexDB(str(leveldb_path), str(blob_path) if blob_path.exists() else None)

    profile = parse_profile(source_info.local_storage_dir) if source_info.local_storage_dir else None
    current_user_oid = (profile or {}).get("oid")
    current_user_name = (profile or {}).get("display_name")

    people_store = get_target_store(wrapper, "Teams:substrate-suggestions-manager", "people")
    conversations_store = get_target_store(wrapper, "Teams:conversation-manager", "conversations")
    replychains_store = get_target_store(wrapper, "Teams:replychain-manager", "replychains")
    call_history_store = get_target_store(wrapper, "Teams:call-history-manager", "call-history")

    guid_to_name = {}
    if current_user_oid and current_user_name:
        guid_to_name[current_user_oid.lower()] = current_user_name

    if people_store is not None:
        with decode_output_context(show_decode_errors):
            for record in people_store.iterate_records(errors_to_stdout=True):
                value = record.value or {}
                guid = clean_text(value.get("ExternalDirectoryObjectId")) or extract_guid(value.get("MRI")) or extract_guid(value.get("Id"))
                name = clean_text(value.get("DisplayName"))
                if guid and name:
                    guid_to_name[guid.lower()] = name

    threads: dict[str, dict] = {}
    thread_titles: dict[str, str] = {}
    thread_types: dict[str, str] = {}

    if conversations_store is not None:
        with decode_output_context(show_decode_errors):
            for record in conversations_store.iterate_records(errors_to_stdout=True):
                value = record.value or {}
                thread_id = normalize_thread_id(value.get("id"))
                if not thread_id:
                    continue
                thread = threads.setdefault(thread_id, init_thread_record(thread_id))
                thread["metadata_quality"] = "detailed"
                thread_types[thread_id] = safe_text(value.get("type")) or ""

                title = safe_text(value.get("title"))
                if title:
                    thread_titles[thread_id] = title
                    thread["label"] = title

                metadata = thread["metadata"]
                thread_properties = value.get("threadProperties") or {}
                if isinstance(thread_properties, dict):
                    mapping = {
                        "topic": "topic",
                        "topicThreadTopic": "topic_thread_topic",
                        "spaceThreadTopic": "space_thread_topic",
                        "threadType": "thread_type",
                        "creator": "creator",
                        "hasTranscript": "has_transcript",
                        "isLastMessageFromMe": "is_last_message_from_me",
                        "isLiveChatEnabled": "is_live_chat_enabled",
                    }
                    for source, dest in mapping.items():
                        if source in thread_properties:
                            metadata[dest] = thread_properties[source]
                    if isinstance(thread_properties.get("meeting"), dict):
                        meeting = thread_properties["meeting"]
                        thread["meeting"] = {
                            "subject": meeting.get("subject"),
                            "startTime": meeting.get("startTime"),
                            "endTime": meeting.get("endTime"),
                            "isCancelled": meeting.get("isCancelled"),
                        }

                participant_ids, participants = flatten_members(value.get("members") or [], guid_to_name)
                thread["participant_ids"] = sorted(set(thread["participant_ids"]) | set(participant_ids))
                thread["participants"] = sorted(set(thread["participants"]) | set(participants))

    best_messages: dict[tuple[str, str], dict] = {}
    calllog_calls: dict[str, dict] = {}
    if replychains_store is not None:
        with decode_output_context(show_decode_errors):
            for record in replychains_store.iterate_records(errors_to_stdout=True):
                value = record.value or {}
                thread_id = normalize_thread_id(value.get("conversationId"))
                if not thread_id:
                    continue
                thread = threads.setdefault(thread_id, init_thread_record(thread_id))
                message_map = value.get("messageMap") or {}

                if thread_id == "48:calllogs":
                    for _, raw in message_map.items():
                        if not isinstance(raw, dict):
                            continue
                        content = raw.get("content")
                        content_string = None
                        if isinstance(content, str):
                            content_string = content
                        elif isinstance(content, (bytes, bytearray)):
                            content_string = bytes(content).decode("utf-8", "ignore")

                        content_text = safe_text(content_string if content_string is not None else content)
                        message_id = safe_text(raw.get("id") or raw.get("clientMessageId") or raw.get("searchKey"))
                        properties = raw.get("properties") or {}
                        call_log = parse_embedded_json(properties.get("call-log")) if isinstance(properties, dict) else None

                        originator = (call_log or {}).get("originatorParticipant") or {}
                        target = (call_log or {}).get("targetParticipant") or {}
                        originator_id = extract_guid(originator.get("id")) or extract_guid((call_log or {}).get("originator"))
                        target_id = extract_guid(target.get("id")) or extract_guid((call_log or {}).get("target"))
                        originator_name = safe_text(originator.get("displayName")) or (guid_to_name.get(originator_id) if originator_id else None)
                        target_name = safe_text(target.get("displayName")) or (guid_to_name.get(target_id) if target_id else None)
                        if originator_id and originator_name:
                            guid_to_name.setdefault(originator_id, originator_name)
                        if target_id and target_name:
                            guid_to_name.setdefault(target_id, target_name)

                        call_id = safe_text((call_log or {}).get("callId")) or extract_guid(content_text) or extract_guid(message_id)
                        timestamp = to_iso_from_millis(raw.get("originalArrivalTime") or raw.get("clientArrivalTime") or value.get("latestDeliveryTime"))
                        shared_correlation_id = safe_text((call_log or {}).get("sharedCorrelationId"))
                        key = call_id or shared_correlation_id or message_id
                        if not key:
                            continue

                        candidate = {
                            "call_id": call_id,
                            "shared_correlation_id": shared_correlation_id,
                            "start_time": safe_text((call_log or {}).get("startTime")) or timestamp,
                            "connect_time": safe_text((call_log or {}).get("connectTime")),
                            "end_time": safe_text((call_log or {}).get("endTime")),
                            "direction": safe_text((call_log or {}).get("callDirection")),
                            "call_type": safe_text((call_log or {}).get("callType")),
                            "call_state": safe_text((call_log or {}).get("callState")) or safe_text(raw.get("messageType")) or "calllog",
                            "originator_display_name": originator_name,
                            "originator_id": originator_id,
                            "target_display_name": target_name,
                            "target_id": target_id,
                            "conversation_id": safe_text((call_log or {}).get("threadId")) or thread_id,
                            "quality": "calllog_enriched" if call_log else "calllog",
                            "source": "ccl:replychains:48:calllogs",
                            "summary_text": content_text,
                        }
                        candidate = {name: field for name, field in candidate.items() if field not in {None, ""}}

                        existing = calllog_calls.get(key)
                        score = sum(
                            1
                            for field in [
                                "start_time",
                                "connect_time",
                                "end_time",
                                "direction",
                                "call_type",
                                "call_state",
                                "originator_display_name",
                                "target_display_name",
                                "shared_correlation_id",
                            ]
                            if candidate.get(field)
                        )
                        existing_score = sum(
                            1
                            for field in [
                                "start_time",
                                "connect_time",
                                "end_time",
                                "direction",
                                "call_type",
                                "call_state",
                                "originator_display_name",
                                "target_display_name",
                                "shared_correlation_id",
                            ]
                            if (existing or {}).get(field)
                        )
                        if existing is None or score > existing_score or len(candidate.get("summary_text", "")) > len(existing.get("summary_text", "")):
                            calllog_calls[key] = candidate
                    continue

                for _, raw in message_map.items():
                    if not isinstance(raw, dict):
                        continue
                    message_id = safe_text(raw.get("id") or raw.get("clientMessageId") or raw.get("searchKey"))
                    if not message_id:
                        continue

                    content = raw.get("content")
                    content_type = safe_text(raw.get("contentType"))
                    message_type = safe_text(raw.get("messageType"))
                    content_string = None
                    if isinstance(content, str):
                        content_string = content
                    elif isinstance(content, (bytes, bytearray)):
                        content_string = bytes(content).decode("utf-8", "ignore")

                    if message_type == "RichText/Html" or (isinstance(content_string, str) and content_string.startswith("<")):
                        content_html = content_string
                        content_text = html_to_text(content_string)
                    else:
                        content_html = None
                        content_text = safe_text(content_string if content_string is not None else content)

                    sender_name = safe_text(raw.get("imDisplayName") or raw.get("fromDisplayNameInToken"))
                    sender_id = extract_guid(raw.get("creator")) or extract_guid(raw.get("fromUserId")) or extract_guid(raw.get("from"))
                    if sender_id and sender_name:
                        guid_to_name.setdefault(sender_id, sender_name)
                    if sender_id or sender_name:
                        merge_thread_participant(thread, sender_id, sender_name or (guid_to_name.get(sender_id) if sender_id else None))

                    timestamp = to_iso_from_millis(raw.get("originalArrivalTime") or raw.get("clientArrivalTime") or value.get("latestDeliveryTime"))
                    message = {
                        "id": message_id,
                        "client_message_id": safe_text(raw.get("clientMessageId")),
                        "timestamp": timestamp,
                        "sender_display_name": sender_name or (guid_to_name.get(sender_id) if sender_id else None),
                        "sender_id": sender_id,
                        "message_type": message_type,
                        "content_type": content_type,
                        "content_html": content_html,
                        "content_text": content_text,
                        "quality": classify_quality(message_type, content_text),
                        "source": "ccl:replychains",
                    }
                    message = {key: value for key, value in message.items() if value not in {None, ""}}

                    key = (thread_id, message_id)
                    existing = best_messages.get(key)
                    if existing is None or (
                        (message.get("quality") == "curated") > (existing.get("quality") == "curated")
                        or len(message.get("content_text", "")) > len(existing.get("content_text", ""))
                    ):
                        best_messages[key] = message

    for (thread_id, _), message in sorted(best_messages.items(), key=lambda item: (item[0][0], item[1].get("timestamp") or "", item[0][1])):
        thread = threads.setdefault(thread_id, init_thread_record(thread_id))
        thread["messages"].append(message)

    calls = []
    if call_history_store is not None:
        with decode_output_context(show_decode_errors):
            for record in call_history_store.iterate_records(errors_to_stdout=True):
                value = record.value or {}
                originator = value.get("originatorParticipant") or {}
                target = value.get("targetParticipant") or {}
                originator_id = extract_guid(originator.get("id"))
                target_id = extract_guid(target.get("id"))
                originator_name = safe_text(originator.get("displayName")) or (guid_to_name.get(originator_id) if originator_id else None)
                target_name = safe_text(target.get("displayName")) or (guid_to_name.get(target_id) if target_id else None)
                if originator_id and originator_name:
                    guid_to_name.setdefault(originator_id, originator_name)
                if target_id and target_name:
                    guid_to_name.setdefault(target_id, target_name)

                calls.append(
                    {
                        "call_id": safe_text(value.get("callId")),
                        "shared_correlation_id": safe_text(value.get("sharedCorrelationId")),
                        "start_time": safe_text(value.get("startTime")),
                        "connect_time": safe_text(value.get("connectTime")),
                        "end_time": safe_text(value.get("endTime")),
                        "direction": safe_text(value.get("callDirection")),
                        "call_type": safe_text(value.get("callType")),
                        "call_state": safe_text(value.get("callState")),
                        "originator_display_name": originator_name,
                        "originator_id": originator_id,
                        "target_display_name": target_name,
                        "target_id": target_id,
                        "conversation_id": safe_text(value.get("conversationId")),
                        "quality": "structured",
                        "source": "ccl:call-history",
                    }
                )

    structured_by_call_id = {(call["call_id"] or "").lower(): call for call in calls if call.get("call_id")}
    structured_by_shared_correlation_id = {
        (call["shared_correlation_id"] or "").lower(): call
        for call in calls
        if call.get("shared_correlation_id")
    }
    for call in sorted(calllog_calls.values(), key=lambda item: (item.get("start_time") or "", item.get("call_id") or "")):
        call_id = (call.get("call_id") or "").lower()
        shared_correlation_id = (call.get("shared_correlation_id") or "").lower()
        existing = None
        if call_id:
            existing = structured_by_call_id.get(call_id)
        if existing is None and shared_correlation_id:
            existing = structured_by_shared_correlation_id.get(shared_correlation_id)
        if existing is not None:
            for field, value in call.items():
                if field in {"quality", "source"}:
                    continue
                if value and not existing.get(field):
                    existing[field] = value
            continue
        calls.append(call)

    for call in calls:
        originator_id = (call.get("originator_id") or "").lower()
        target_id = (call.get("target_id") or "").lower()
        if originator_id and not call.get("originator_display_name") and guid_to_name.get(originator_id):
            call["originator_display_name"] = guid_to_name[originator_id]
        if target_id and not call.get("target_display_name") and guid_to_name.get(target_id):
            call["target_display_name"] = guid_to_name[target_id]

        conversation_id = normalize_thread_id(call.get("conversation_id"))
        if conversation_id and conversation_id in threads:
            thread = threads[conversation_id]
            merge_thread_participant(thread, originator_id, call.get("originator_display_name"))
            merge_thread_participant(thread, target_id, call.get("target_display_name"))

    for thread_id, thread in threads.items():
        for participant_id in list(thread.get("participant_ids", [])):
            if participant_id and guid_to_name.get(participant_id):
                merge_thread_participant(thread, participant_id, guid_to_name[participant_id])

        for message in thread.get("messages", []):
            merge_thread_participant(
                thread,
                message.get("sender_id"),
                message.get("sender_display_name") or (guid_to_name.get(message.get("sender_id")) if message.get("sender_id") else None),
            )

        if is_placeholder_label(thread.get("label"), thread_id):
            explicit_title = thread_titles.get(thread_id)
            if explicit_title:
                thread["label"] = explicit_title

        if is_placeholder_label(thread.get("label"), thread_id) and thread["category"] == "chat_space":
            others = [name for name in thread.get("participants", []) if name != current_user_name]
            if len(others) == 1:
                thread["label"] = others[0]

        if is_placeholder_label(thread.get("label"), thread_id):
            metadata = thread.get("metadata") or {}
            for field in ["space_thread_topic", "topic_thread_topic", "topic"]:
                if clean_text(metadata.get(field)):
                    thread["label"] = clean_text(metadata[field])
                    break

        if is_placeholder_label(thread.get("label"), thread_id):
            participant_label = build_participant_label(thread.get("participants", []), current_user_name)
            if participant_label:
                thread["label"] = participant_label

        if is_placeholder_label(thread.get("label"), thread_id):
            thread["label"] = thread_id

        thread["message_count"] = len(thread["messages"])
        thread["participants"] = sorted(set(thread.get("participants", [])))
        thread["participant_ids"] = sorted(set(thread.get("participant_ids", [])))

    ordered_threads = sorted(
        threads.values(),
        key=lambda item: (
            item["category"] != "meeting",
            item["metadata_quality"] != "detailed",
            -item.get("message_count", 0),
            item["label"],
        ),
    )

    summary = {
        "threads_total": len(ordered_threads),
        "threads_with_detail": sum(thread["metadata_quality"] == "detailed" for thread in ordered_threads),
        "threads_with_messages": sum(bool(thread["messages"]) for thread in ordered_threads),
        "messages_total": sum(thread["message_count"] for thread in ordered_threads),
        "calls_total": len(calls),
        "people_total": len(guid_to_name),
    }

    return {
        "export_format": "teams-ccl-canonical-v1",
        "generated_at": datetime.now(tz=timezone.utc).isoformat(),
        "source_root": str(source_info.profile_root),
        "search_root": str(source_info.search_root),
        "platform": source_info.platform_name,
        "leveldb_path": str(leveldb_path),
        "blob_path": str(blob_path) if blob_path.exists() else None,
        "local_storage_dir": str(source_info.local_storage_dir) if source_info.local_storage_dir else None,
        "discovery_method": source_info.discovery_method,
        "backend": "ccl_chromium_reader",
        "profile": profile,
        "summary": summary,
        "threads": ordered_threads,
        "calls": sorted(calls, key=lambda item: item.get("start_time") or "", reverse=True),
        "guid_directory": dict(sorted(guid_to_name.items())),
        "limitations": [
            "Built from structured IndexedDB object stores using ccl_chromium_reader and the matching blob directory.",
            "A small number of keys reported decode errors during CCL iteration and may be missing from this export.",
            "This canonical export does not yet merge all auxiliary stores such as reactions, attachments, or notification surfaces.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Export Teams data from structured IndexedDB stores via ccl_chromium_reader.")
    parser.add_argument("--root", help="Teams source root, profile directory, copied evidence root, or IndexedDB directory.")
    parser.add_argument("--output", default="teams_ccl_canonical_v1.json", help="Path to write JSON output.")
    args = parser.parse_args()

    payload = build_export(Path(args.root).resolve() if args.root else None)
    Path(args.output).resolve().write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


if __name__ == "__main__":
    main()
