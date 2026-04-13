#!/usr/bin/env python3

import argparse
import json
import re
from pathlib import Path


SAFE_RE = re.compile(r"[^A-Za-z0-9._-]+")


def safe_name(value: str) -> str:
    value = SAFE_RE.sub("_", value or "")
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "untitled"


def thread_csv_path(thread: dict) -> str:
    label = thread.get("label") or thread.get("id") or "thread"
    filename = f"{safe_name(thread.get('category') or 'thread')}__{safe_name(label)}__{safe_name(thread['id'])[:12]}.csv"
    return f"teams_ccl_csv_v1/conversations/{filename}"


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>MS Teams Viewer</title>
  <style>
    :root {
      --bg: #f4f1ea;
      --panel: #fffaf2;
      --ink: #1f2a2b;
      --muted: #667371;
      --line: #d7d0c4;
      --accent: #1d6f6d;
      --accent-soft: #dcefed;
      --warn: #b55834;
      --shadow: 0 8px 24px rgba(31,42,43,.08);
      --mono: "SFMono-Regular", Consolas, "Liberation Mono", monospace;
      --sans: "Iowan Old Style", "Palatino Linotype", "Book Antiqua", Palatino, Georgia, serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: var(--sans);
      color: var(--ink);
      background-color: var(--bg);
      background:
        radial-gradient(circle at top left, rgba(29,111,109,.08), transparent 28%),
        radial-gradient(circle at bottom right, rgba(181,88,52,.08), transparent 24%),
        linear-gradient(180deg, #f7f4ee 0%, #efe9df 100%);
      min-height: 100vh;
    }
    .app {
      display: grid;
      grid-template-columns: 360px 1fr;
      min-height: 100vh;
    }
    .sidebar {
      border-right: 1px solid var(--line);
      background: rgba(255,250,242,.88);
      backdrop-filter: blur(8px);
      padding: 20px;
      position: sticky;
      top: 0;
      height: 100vh;
      overflow: auto;
    }
    .main {
      padding: 20px 24px 32px;
      overflow: auto;
    }
    h1, h2, h3 { margin: 0; font-weight: 600; }
    .title {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
      margin-bottom: 12px;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
      margin: 16px 0 20px;
    }
    .card, .panel {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 16px;
      box-shadow: var(--shadow);
    }
    .card { padding: 12px 14px; }
    .card .label {
      font-size: 12px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: .08em;
      margin-bottom: 6px;
    }
    .card .value {
      font-size: 28px;
      line-height: 1;
      color: var(--accent);
    }
    .view-tabs {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
      margin-bottom: 12px;
    }
    .tab {
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 10px 12px;
      background: rgba(255,255,255,.72);
      color: var(--ink);
      font: inherit;
      cursor: pointer;
      transition: background .15s ease, border-color .15s ease, transform .15s ease;
    }
    .tab:hover { transform: translateY(-1px); }
    .tab.active {
      background: var(--accent-soft);
      border-color: var(--accent);
      color: var(--accent);
    }
    .toolbar {
      display: grid;
      gap: 10px;
      margin-bottom: 12px;
    }
    .hidden { display: none !important; }
    input, select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 10px 12px;
      background: #fff;
      color: var(--ink);
      font: inherit;
    }
    .chip-row {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 12px;
    }
    .chip {
      padding: 6px 10px;
      border-radius: 999px;
      background: #fff;
      border: 1px solid var(--line);
      font-size: 12px;
      color: var(--muted);
    }
    .list {
      display: grid;
      gap: 8px;
    }
    .list-row {
      padding: 12px 14px;
      border-radius: 14px;
      border: 1px solid transparent;
      background: rgba(255,255,255,.7);
      cursor: pointer;
      transition: transform .15s ease, border-color .15s ease, background .15s ease;
    }
    .list-row:hover { transform: translateY(-1px); border-color: var(--line); }
    .list-row.active { border-color: var(--accent); background: var(--accent-soft); }
    .list-row .meta {
      display: flex;
      justify-content: space-between;
      gap: 8px;
      color: var(--muted);
      font-size: 12px;
      margin-top: 6px;
    }
    .preview {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.35;
      margin-top: 8px;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
    }
    .panel {
      padding: 18px 20px;
      margin-bottom: 16px;
    }
    .detail-toolbar {
      display: grid;
      gap: 10px;
      margin-top: 14px;
    }
    .toolbar-row {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }
    .range-controls {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      align-items: end;
    }
    .toolbar-divider {
      height: 1px;
      background: var(--line);
      width: 100%;
    }
    .range-field {
      display: grid;
      gap: 4px;
      min-width: 150px;
    }
    .range-field span {
      color: var(--muted);
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .06em;
    }
    .segmented {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .seg-btn {
      border: 1px solid var(--line);
      border-radius: 999px;
      background: rgba(255,255,255,.75);
      color: var(--muted);
      font: inherit;
      padding: 6px 10px;
      cursor: pointer;
    }
    .seg-btn.active {
      border-color: var(--accent);
      color: var(--accent);
      background: var(--accent-soft);
    }
    .stat-strip {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 10px;
      margin-top: 14px;
    }
    .stat-card {
      background: rgba(255,255,255,.82);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 10px 12px;
    }
    .stat-card .k {
      color: var(--muted);
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .06em;
      margin-bottom: 4px;
    }
    .stat-card .v {
      color: var(--ink);
      font-size: 14px;
      line-height: 1.35;
    }
    .section-label {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .08em;
      margin: 18px 0 10px;
    }
    .section-divider {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-top: 18px;
    }
    .section-divider::before,
    .section-divider::after {
      content: "";
      height: 1px;
      background: var(--line);
      flex: 1;
    }
    .section-divider .section-label {
      margin: 0;
      white-space: nowrap;
    }
    .jump-list {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 12px;
    }
    .subtle {
      color: var(--muted);
      font-size: 13px;
    }
    .chips {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-top: 10px;
    }
    .messages {
      display: grid;
      gap: 10px;
      margin-top: 14px;
    }
    .search-results {
      display: grid;
      gap: 14px;
      margin-top: 14px;
    }
    .search-group {
      border: 1px solid var(--line);
      border-radius: 14px;
      background: rgba(255,255,255,.78);
      padding: 14px;
      box-shadow: var(--shadow);
    }
    .search-group-header {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }
    .search-group-meta {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.4;
      margin-top: 6px;
    }
    .search-group-divider {
      height: 1px;
      background: var(--line);
      margin-top: 2px;
    }
    .msg.search-match {
      border-color: var(--accent);
      box-shadow: 0 0 0 2px rgba(29,111,109,.12);
      background: #f5fffd;
    }
    .search-term {
      font-weight: 700;
      background: rgba(255, 229, 120, .7);
      border-radius: 3px;
      padding: 0 1px;
    }
    .msg {
      padding: 12px 14px;
      border-radius: 14px;
      border: 1px solid var(--line);
      border-left: 6px solid var(--msg-accent, var(--line));
      background: rgba(255,255,255,.85);
      scroll-margin-top: 24px;
    }
    .msg.event {
      background: rgba(255,248,244,.92);
    }
    .msg.focused {
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(29,111,109,.14);
      background: #f7fffd;
    }
    .msg .head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      font-size: 13px;
      color: var(--muted);
      margin-bottom: 6px;
    }
    .msg .sender {
      color: var(--msg-accent, var(--accent));
      font-weight: 600;
    }
    .msg .body {
      white-space: pre-wrap;
      line-height: 1.4;
      word-break: break-word;
    }
    .call-event {
      border: 2px solid var(--call-outline, var(--line));
      border-radius: 12px;
      background: var(--call-bg, rgba(255,255,255,.72));
      padding: 12px;
      display: grid;
      gap: 10px;
      box-shadow: 0 6px 16px rgba(31,42,43,.10);
    }
    .call-event-header {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }
    .call-event-title {
      color: var(--call-title, var(--accent));
      font-weight: 600;
    }
    .call-event-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .call-event-block {
      background: var(--call-block-bg, rgba(255,255,255,.88));
      border: 1px solid var(--call-outline, var(--line));
      border-radius: 10px;
      padding: 8px 10px;
    }
    .call-event-block .k {
      color: var(--muted);
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .06em;
      margin-bottom: 4px;
    }
    .call-event-actions {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .call-link {
      border: 1px solid var(--accent);
      border-radius: 999px;
      background: var(--accent-soft);
      color: var(--accent);
      font: inherit;
      padding: 6px 10px;
      cursor: pointer;
    }
    .call-link:hover {
      background: #cfe7e4;
    }
    .toast {
      position: fixed;
      right: 20px;
      bottom: 20px;
      z-index: 20;
      background: rgba(31,42,43,.92);
      color: #fff;
      padding: 10px 14px;
      border-radius: 12px;
      box-shadow: var(--shadow);
      opacity: 0;
      transform: translateY(10px);
      transition: opacity .18s ease, transform .18s ease;
      pointer-events: none;
    }
    .toast.show {
      opacity: 1;
      transform: translateY(0);
    }
    details.raw-event {
      border-top: 1px dashed var(--line);
      padding-top: 8px;
    }
    details.raw-event summary {
      cursor: pointer;
      color: var(--muted);
    }
    .meta-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
      margin-top: 14px;
    }
    .meta-block {
      background: rgba(255,255,255,.82);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 10px 12px;
    }
    .meta-block .k {
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .06em;
      margin-bottom: 4px;
    }
    .mono {
      font-family: var(--mono);
      font-size: 12px;
      word-break: break-word;
    }
    .empty {
      color: var(--muted);
      padding: 24px 8px;
    }
    @media (max-width: 980px) {
      .app { grid-template-columns: 1fr; }
      .sidebar { position: static; height: auto; border-right: 0; border-bottom: 1px solid var(--line); }
      .summary { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .meta-grid { grid-template-columns: 1fr; }
      .call-event-grid { grid-template-columns: 1fr; }
      .stat-strip { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside class="sidebar">
      <div class="title">
        <h1>MS Teams Viewer</h1>
      </div>
      <div class="view-tabs">
        <button id="viewMessages" class="tab active" type="button">Messages</button>
        <button id="viewCalls" class="tab" type="button">Calls</button>
      </div>
      <div id="messagesTools" class="toolbar">
        <input id="messageSearch" type="search" placeholder="Search conversations and messages">
        <select id="categoryFilter">
          <option value="">All message categories</option>
          <option value="chat_space">Chats</option>
          <option value="thread">Group chats</option>
        </select>
      </div>
      <div id="callsTools" class="toolbar hidden">
        <input id="callSearch" type="search" placeholder="Search calls, participants, direction, state">
        <select id="callDirectionFilter">
          <option value="">All calls</option>
          <option value="incoming">Incoming</option>
          <option value="outgoing">Outgoing</option>
          <option value="missed">Missed</option>
        </select>
      </div>
      <div class="chip-row">
        <div id="sidebarCount" class="chip"></div>
      </div>
      <div id="threadList" class="list"></div>
      <div id="callList" class="list hidden"></div>
    </aside>
    <main class="main">
      <div class="summary">
        <div class="card"><div class="label">Threads</div><div class="value" id="sumThreads"></div></div>
        <div class="card"><div class="label">Messages</div><div class="value" id="sumMessages"></div></div>
        <div class="card"><div class="label">Calls</div><div class="value" id="sumCalls"></div></div>
      </div>
      <section id="contentPanel" class="panel"></section>
    </main>
  </div>
  <div id="toast" class="toast"></div>
  <script>
    const DATA = __DATA__;

    const state = {
      view: "messages",
      messageSearch: "",
      callSearch: "",
      category: "",
      callDirection: "",
      messageViewFilter: "all",
      threadDateFrom: "",
      threadDateTo: "",
      expandRawEvents: false,
      threadId: null,
      callKey: null,
      focusCallKey: null,
      focusTimestamp: null,
    };

    const contentPanel = document.getElementById("contentPanel");
    const mainPanel = document.querySelector(".main");
    const threadList = document.getElementById("threadList");
    const callList = document.getElementById("callList");
    const messagesTools = document.getElementById("messagesTools");
    const callsTools = document.getElementById("callsTools");
    const sidebarCount = document.getElementById("sidebarCount");
    const viewMessages = document.getElementById("viewMessages");
    const viewCalls = document.getElementById("viewCalls");
    const messageSearch = document.getElementById("messageSearch");
    const categoryFilter = document.getElementById("categoryFilter");
    const callSearch = document.getElementById("callSearch");
    const callDirectionFilter = document.getElementById("callDirectionFilter");
    const toast = document.getElementById("toast");

    const CALLS_BY_ID = new Map();
    const CALLS_BY_SHARED = new Map();
    const CALLS_BY_KEY = new Map();
    const CALLS_BY_THREAD_ID = new Map();
    const CALLS_BY_GUID_PAIR = new Map();
    const CALLS_BY_NAME_PAIR = new Map();
    const SYNTHETIC_CALLS_CACHE = new Map();
    for (const call of DATA.calls || []) {
      if (call.call_id) CALLS_BY_ID.set(String(call.call_id).toLowerCase(), call);
      if (call.shared_correlation_id) CALLS_BY_SHARED.set(String(call.shared_correlation_id).toLowerCase(), call);
      CALLS_BY_KEY.set(callKey(call), call);
      if (call.conversation_id && call.conversation_id !== "48:calllogs") {
        addToMapList(CALLS_BY_THREAD_ID, String(call.conversation_id), call);
      }
      const guidPair = participantGuidPairKey(callParticipantIds(call));
      if (guidPair) addToMapList(CALLS_BY_GUID_PAIR, guidPair, call);
      const namePair = participantNamePairKey(callParticipantNames(call));
      if (namePair) addToMapList(CALLS_BY_NAME_PAIR, namePair, call);
    }

    document.getElementById("sumThreads").textContent = DATA.summary.threads_total;
    document.getElementById("sumMessages").textContent = DATA.summary.messages_total;
    document.getElementById("sumCalls").textContent = DATA.summary.calls_total;

    function fmt(value) {
      if (!value) return "[no_time]";
      const date = new Date(value);
      if (Number.isNaN(date.getTime())) return value;
      return date.toLocaleString();
    }

    function timeValue(value) {
      if (!value) return Number.NEGATIVE_INFINITY;
      const parsed = Date.parse(value);
      if (Number.isNaN(parsed)) return Number.NEGATIVE_INFINITY;
      return parsed;
    }

    function escapeHtml(value) {
      return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function stripFcs(value) {
      return String(value ?? "")
        .replace(/\\s*\\(FCS\\)/g, "")
        .replace(/\\s{2,}/g, " ")
        .trim();
    }

    function addToMapList(map, key, value) {
      if (!key) return;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(value);
    }

    function dedupe(values) {
      return [...new Set((values || []).filter(Boolean))];
    }

    function normalizeGuid(value) {
      return String(value || "").trim().toLowerCase();
    }

    function normalizeName(value) {
      return stripFcs(value).toLowerCase();
    }

    const CURRENT_USER_ID = normalizeGuid(DATA.profile && DATA.profile.oid);
    const CURRENT_USER_NAME = normalizeName(DATA.profile && DATA.profile.display_name);
    const MESSAGE_COLOR_SELF = "#355CDE";
    const MESSAGE_COLOR_PALETTE = [
      "#0F766E",
      "#C05621",
      "#B7791F",
      "#B83280",
      "#2F855A",
      "#9C4221",
      "#BE123C",
      "#4D7C0F",
    ];
    const CALL_VISUALS = {
      "#355CDE": { bg: "#88B5FF", outline: "#18357A", block: "#EAF3FF", title: "#16306F" },
      "#0F766E": { bg: "#72F2E7", outline: "#104641", block: "#EDFFFD", title: "#0E4C46" },
      "#C05621": { bg: "#FFB887", outline: "#6F2C08", block: "#FFF1E8", title: "#6F2C08" },
      "#B7791F": { bg: "#FFE27B", outline: "#6C4608", block: "#FFF9DE", title: "#654107" },
      "#B83280": { bg: "#FF9DDB", outline: "#6F1A4A", block: "#FFF2FA", title: "#6C1948" },
      "#2F855A": { bg: "#96F5C1", outline: "#18482F", block: "#EEFFF5", title: "#18482F" },
      "#9C4221": { bg: "#FFAA84", outline: "#5A1F0B", block: "#FFF0E9", title: "#5A1F0B" },
      "#BE123C": { bg: "#FF96AE", outline: "#6C1026", block: "#FFF1F5", title: "#6C1026" },
      "#4D7C0F": { bg: "#C6FA72", outline: "#2E4A09", block: "#F7FFE9", title: "#2E4A09" },
      "#8A6F5A": { bg: "#E9D7C6", outline: "#4D3B2C", block: "#FBF5F0", title: "#4D3B2C" },
    };

    function participantGuidPairKey(values) {
      const items = dedupe((values || []).map(normalizeGuid).filter(Boolean)).sort();
      return items.length === 2 ? items.join("|") : "";
    }

    function participantNamePairKey(values) {
      const items = dedupe((values || []).map(normalizeName).filter(Boolean)).sort();
      return items.length === 2 ? items.join("|") : "";
    }

    function callParticipantIds(call) {
      return dedupe([call.originator_id, call.target_id].map(normalizeGuid).filter(Boolean));
    }

    function callParticipantNames(call) {
      return dedupe([call.originator_display_name, call.target_display_name].map(normalizeName).filter(Boolean));
    }

    function callTimelineTimestamp(call) {
      return call.end_time || call.start_time || call.connect_time || "";
    }

    function callDurationSeconds(call) {
      const start = Date.parse(call.start_time || "");
      const end = Date.parse(call.end_time || "");
      if (Number.isNaN(start) || Number.isNaN(end) || end < start) return null;
      return Math.round((end - start) / 1000);
    }

    function displayCallDirection(call) {
      return String(call.direction || "").toLowerCase();
    }

    function shouldMarkMissed(call) {
      const seconds = callDurationSeconds(call);
      return (
        displayCallDirection(call) === "outgoing" &&
        seconds !== null &&
        seconds < 45
      );
    }

    function displayCallState(call) {
      if (shouldMarkMissed(call)) return "missed";
      return String(call.call_state || "").toLowerCase();
    }

    function displayCallStatus(call) {
      const direction = displayCallDirection(call);
      const state = displayCallState(call);
      if (direction && state) return `${direction} | ${state}`;
      return direction || state || "";
    }

    function hashString(value) {
      let hash = 0;
      const text = String(value || "");
      for (let index = 0; index < text.length; index += 1) {
        hash = ((hash << 5) - hash) + text.charCodeAt(index);
        hash |= 0;
      }
      return Math.abs(hash);
    }

    function messageSenderKey(message) {
      return normalizeGuid(message.sender_id || "") || normalizeName(message.sender_display_name || "");
    }

    function isOwnMessage(message) {
      const senderId = normalizeGuid(message.sender_id || "");
      const senderName = normalizeName(message.sender_display_name || "");
      return Boolean(
        (CURRENT_USER_ID && senderId && senderId === CURRENT_USER_ID) ||
        (CURRENT_USER_NAME && senderName && senderName === CURRENT_USER_NAME)
      );
    }

    function messageAccent(message) {
      if (isOwnMessage(message)) return MESSAGE_COLOR_SELF;
      const key = messageSenderKey(message);
      if (!key) return "#8A6F5A";
      return MESSAGE_COLOR_PALETTE[hashString(key) % MESSAGE_COLOR_PALETTE.length];
    }

    function callVisualsForColor(color) {
      return CALL_VISUALS[color] || CALL_VISUALS["#8A6F5A"];
    }

    function inboundCallerAccent(linkedCall, message) {
      if (linkedCall && (linkedCall.originator_display_name || linkedCall.originator_id)) {
        return messageAccent({
          sender_display_name: linkedCall.originator_display_name || "",
          sender_id: linkedCall.originator_id || "",
        });
      }
      return messageAccent(message);
    }

    function callCardVisuals(message, linkedCall) {
      if (!linkedCall) {
        return callVisualsForColor(messageAccent(message));
      }
      if (displayCallDirection(linkedCall) === "outgoing") {
        return callVisualsForColor(MESSAGE_COLOR_SELF);
      }
      if (displayCallDirection(linkedCall) === "incoming") {
        return callVisualsForColor(inboundCallerAccent(linkedCall, message));
      }
      return callVisualsForColor(messageAccent(message));
    }

    function truncate(value, maxLength = 110) {
      const text = String(value ?? "").trim();
      if (text.length <= maxLength) return text;
      return `${text.slice(0, maxLength - 1)}…`;
    }

    function showToast(message) {
      toast.textContent = message;
      toast.classList.add("show");
      window.clearTimeout(showToast.timer);
      showToast.timer = window.setTimeout(() => toast.classList.remove("show"), 1600);
    }

    function scrollMainToTop() {
      window.requestAnimationFrame(() => {
        if (mainPanel) {
          mainPanel.scrollTop = 0;
          mainPanel.scrollTo({ top: 0, behavior: "auto" });
        }
        if (contentPanel) {
          contentPanel.scrollTop = 0;
        }
        document.documentElement.scrollTop = 0;
        document.body.scrollTop = 0;
        window.scrollTo({ top: 0, left: 0, behavior: "auto" });
      });
    }

    async function copyText(value, label) {
      const text = String(value ?? "");
      if (!text) return;
      try {
        await navigator.clipboard.writeText(text);
      } catch (error) {
        const area = document.createElement("textarea");
        area.value = text;
        area.setAttribute("readonly", "");
        area.style.position = "absolute";
        area.style.left = "-9999px";
        document.body.appendChild(area);
        area.select();
        document.execCommand("copy");
        document.body.removeChild(area);
      }
      showToast(label);
    }

    function prettyCategory(category) {
      if (category === "chat_space") return "Chat";
      if (category === "meeting") return "Meeting";
      if (category === "thread") return "Group Chat";
      return category || "";
    }

    function isChatBasedThread(thread) {
      return ["chat_space", "thread", "meeting"].includes(thread.category);
    }

    function threadSearchText(thread) {
      const parts = [
        thread.label,
        thread.id,
        thread.category,
        ...(thread.participants || []),
        ...thread.messages.slice(0, 100).map(message => message.content_text || ""),
        ...syntheticCallsForThread(thread).slice(0, 40).map(call => call.summary_text || call.call_state || call.call_type || ""),
      ];
      return parts.filter(Boolean).join(" ").toLowerCase();
    }

    function searchTerms() {
      return [...new Set(String(state.messageSearch || "")
        .trim()
        .toLowerCase()
        .split(" ")
        .map(term => term.trim())
        .filter(Boolean))];
    }

    function textMatchesSearch(text) {
      const haystack = String(text || "").toLowerCase();
      const query = String(state.messageSearch || "").trim().toLowerCase();
      if (!query) return true;
      if (!haystack) return false;
      if (haystack.includes(query)) return true;
      const terms = searchTerms();
      return terms.length > 1 && terms.every(term => haystack.includes(term));
    }

    function highlightSearchHtml(value) {
      const text = String(value ?? "");
      const terms = searchTerms().sort((left, right) => right.length - left.length);
      if (!text || !terms.length) return escapeHtml(text);

      const lower = text.toLowerCase();
      let html = "";
      let index = 0;

      while (index < text.length) {
        let matchedTerm = "";
        for (const term of terms) {
          if (term && lower.startsWith(term, index)) {
            matchedTerm = text.slice(index, index + term.length);
            break;
          }
        }
        if (matchedTerm) {
          html += `<strong class="search-term">${escapeHtml(matchedTerm)}</strong>`;
          index += matchedTerm.length;
          continue;
        }
        html += escapeHtml(text[index]);
        index += 1;
      }

      return html;
    }

    function latestThreadMessage(thread) {
      let bestMessage = null;
      let bestTime = Number.NEGATIVE_INFINITY;
      for (const message of thread.messages || []) {
        const current = timeValue(message.timestamp);
        if (current >= bestTime) {
          bestTime = current;
          bestMessage = message;
        }
      }
      return bestMessage;
    }

    function latestThreadPreview(thread) {
      const message = latestThreadMessage(thread);
      if (!message) return "No recoverable message preview";
      if (message.message_type === "Event/Call") {
        return truncate(`Call event: ${prettyCallEventName(callEventName(message.content_text || ""))}`);
      }
      if (message.message_type === "ThreadActivity/AddMember") {
        const parsed = parseAddMemberEvent(message);
        return truncate(`Member added: ${parsed.addedNames.join(" / ") || "Unknown"}`);
      }
      return truncate(stripFcs(message.content_text || message.message_type || ""));
    }

    function threadLastTimestamp(thread) {
      let latest = Number.NEGATIVE_INFINITY;
      for (const message of thread.messages || []) {
        latest = Math.max(latest, timeValue(message.timestamp));
      }
      return latest;
    }

    function threadLastTimestampRaw(thread) {
      let best = null;
      let latest = Number.NEGATIVE_INFINITY;
      for (const message of thread.messages || []) {
        const current = timeValue(message.timestamp);
        if (current >= latest) {
          latest = current;
          best = message.timestamp || best;
        }
      }
      return best;
    }

    function filteredThreads() {
      return DATA.threads
        .filter(thread => {
          if (thread.message_count <= 0 && thread.category !== "meeting") return false;
          if (state.category === "chat_space") {
            if (!["chat_space", "meeting"].includes(thread.category)) return false;
          } else if (state.category && thread.category !== state.category) {
            return false;
          }
          if (state.messageSearch && !textMatchesSearch(threadSearchText(thread))) return false;
          return true;
        })
        .sort((left, right) => {
          const timeDiff = threadLastTimestamp(right) - threadLastTimestamp(left);
          if (timeDiff !== 0) return timeDiff;
          return (left.label || left.id || "").localeCompare(right.label || right.id || "");
        });
    }

    function callKey(call) {
      return [
        call.call_id || "",
        call.start_time || "",
        call.shared_correlation_id || "",
        call.originator_id || "",
        call.target_id || "",
      ].join("|");
    }

    function callLabel(call) {
      const origin = call.originator_display_name || call.originator_id || "";
      const target = call.target_display_name || call.target_id || "";
      if (origin && target) return `${origin} -> ${target}`;
      return origin || target || call.summary_text || call.call_id || "[unknown call]";
    }

    function extractGuids(text) {
      return Array.from(String(text || "").matchAll(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g))
        .map(match => match[0].toLowerCase());
    }

    function findLinkedCall(message) {
      if (message.synthetic_call_key && CALLS_BY_KEY.has(message.synthetic_call_key)) {
        return CALLS_BY_KEY.get(message.synthetic_call_key);
      }
      const ids = extractGuids(message.content_text || "");
      for (const id of ids) {
        if (CALLS_BY_ID.has(id)) return CALLS_BY_ID.get(id);
      }
      for (const id of ids) {
        if (CALLS_BY_SHARED.has(id)) return CALLS_BY_SHARED.get(id);
      }
      return null;
    }

    function callEventName(text) {
      const match = String(text || "").match(/\\b(callStarted|callEnded|callMissed|callAccepted|callRejected|callCancelled)\\b/i);
      return match ? match[1] : "callEvent";
    }

    function prettyCallEventName(value) {
      return String(value || "")
        .replace(/^call/, "Call ")
        .replace(/([a-z])([A-Z])/g, "$1 $2")
        .trim();
    }

    function collectEventParticipants(message) {
      const seen = [];
      for (const guid of extractGuids(message.content_text || "")) {
        const name = DATA.guid_directory && DATA.guid_directory[guid];
        if (name && !seen.includes(name)) {
          seen.push(name);
        }
      }
      if (message.sender_display_name && !seen.includes(message.sender_display_name)) {
        seen.push(message.sender_display_name);
      }
      return seen;
    }

    function representedCallKeys(thread) {
      const keys = new Set();
      for (const message of thread.messages || []) {
        const linked = findLinkedCall(message);
        if (linked) keys.add(callKey(linked));
      }
      return keys;
    }

    function syntheticCallsForThread(thread) {
      if (!thread || !thread.id) return [];
      if (SYNTHETIC_CALLS_CACHE.has(thread.id)) return SYNTHETIC_CALLS_CACHE.get(thread.id);

      const represented = representedCallKeys(thread);
      const matches = [];
      const seen = new Set();

      for (const call of CALLS_BY_THREAD_ID.get(String(thread.id)) || []) {
        const key = callKey(call);
        if (!represented.has(key) && !seen.has(key)) {
          seen.add(key);
          matches.push(call);
        }
      }

      const isTwoPersonConversation = (thread.participants || []).length === 2;
      if (isTwoPersonConversation) {
        const guidPair = participantGuidPairKey(thread.participant_ids || []);
        const namePair = participantNamePairKey(thread.participants || []);
        const pairMatches = [
          ...(guidPair ? (CALLS_BY_GUID_PAIR.get(guidPair) || []) : []),
          ...(namePair ? (CALLS_BY_NAME_PAIR.get(namePair) || []) : []),
        ];
        for (const call of pairMatches) {
          const key = callKey(call);
          if (!represented.has(key) && !seen.has(key)) {
            seen.add(key);
            matches.push(call);
          }
        }
      }

      matches.sort((left, right) => timeValue(callTimelineTimestamp(right)) - timeValue(callTimelineTimestamp(left)));
      SYNTHETIC_CALLS_CACHE.set(thread.id, matches);
      return matches;
    }

    function syntheticCallMessages(thread) {
      return syntheticCallsForThread(thread).map(call => ({
        id: `synthetic-call:${callKey(call)}`,
        timestamp: callTimelineTimestamp(call),
        sender_display_name: call.originator_display_name || call.target_display_name || "Call Record",
        sender_id: call.originator_id || call.target_id || "",
        message_type: "Call/History",
        quality: "event",
        content_text: call.summary_text || "",
        synthetic_call: true,
        synthetic_call_key: callKey(call),
      }));
    }

    function renderCallEventBody(message) {
      const raw = message.content_text || "";
      const linkedCall = findLinkedCall(message);
      const eventName = message.synthetic_call ? "Call Record" : prettyCallEventName(callEventName(raw));
      const participants = collectEventParticipants(message);
      const linkedKey = linkedCall ? callKey(linkedCall) : "";
      const visuals = callCardVisuals(message, linkedCall);
      const syntheticParticipants = linkedCall
        ? [linkedCall.originator_display_name, linkedCall.target_display_name].filter(Boolean).map(name => stripFcs(name)).join(" / ")
        : "";
      const participantLabel = message.synthetic_call
        ? syntheticParticipants
        : (participants.length ? participants.map(name => stripFcs(name)).join(" / ") : "");

      return `
        <div class="call-event" style="--call-bg:${escapeHtml(visuals.bg)};--call-outline:${escapeHtml(visuals.outline)};--call-block-bg:${escapeHtml(visuals.block)};--call-title:${escapeHtml(visuals.title)};">
          <div class="call-event-header">
            <div class="call-event-title">${highlightSearchHtml(eventName)}</div>
            <div class="call-event-actions">
              ${linkedCall ? `<button class="call-link open-call-record" type="button" data-call-key="${escapeHtml(linkedKey)}">Open In Calls</button>` : ``}
            </div>
          </div>
          <div class="call-event-grid">
            <div class="call-event-block"><div class="k">Participants</div><div>${highlightSearchHtml(participantLabel || stripFcs(callLabel(linkedCall || {})))}</div></div>
            <div class="call-event-block"><div class="k">Call Type</div><div>${highlightSearchHtml(linkedCall ? (linkedCall.call_type || "") : "")}</div></div>
            <div class="call-event-block"><div class="k">Direction</div><div>${highlightSearchHtml(linkedCall ? displayCallDirection(linkedCall) : "")}</div></div>
            <div class="call-event-block"><div class="k">State</div><div>${highlightSearchHtml(linkedCall ? displayCallState(linkedCall) : eventName)}</div></div>
            <div class="call-event-block"><div class="k">Start Time</div><div>${escapeHtml(fmt(linkedCall ? linkedCall.start_time : message.timestamp))}</div></div>
            <div class="call-event-block"><div class="k">End Time</div><div>${escapeHtml(fmt(linkedCall ? linkedCall.end_time : null))}</div></div>
          </div>
          ${message.synthetic_call
            ? `<div class="subtle">Merged from the Calls dataset for this conversation.</div>`
            : `<details class="raw-event" ${state.expandRawEvents ? "open" : ""}>
                 <summary>Show raw event payload</summary>
                 <div class="body mono">${escapeHtml(raw)}</div>
               </details>`}
        </div>
      `;
    }

    function messageLinkedCallKey(message) {
      if (message.synthetic_call_key) return message.synthetic_call_key;
      const linked = findLinkedCall(message);
      return linked ? callKey(linked) : "";
    }

    function threadTimelineItems(thread) {
      return [...(thread.messages || []), ...syntheticCallMessages(thread)];
    }

    function hasNearbyCuratedMessages(thread, call) {
      const start = timeValue(call.start_time || call.end_time || "");
      const end = timeValue(call.end_time || call.start_time || "");
      if (!Number.isFinite(start) && !Number.isFinite(end)) return false;
      const windowStart = (Number.isFinite(start) ? start : end) - (15 * 60 * 1000);
      const windowEnd = (Number.isFinite(end) ? end : start) + (15 * 60 * 1000);
      return (thread.messages || []).some(message => {
        if (message.quality === "event") return false;
        const current = timeValue(message.timestamp);
        return Number.isFinite(current) && current >= windowStart && current <= windowEnd;
      });
    }

    function resolveGuidLabel(guid) {
      const clean = String(guid || "").toLowerCase();
      return stripFcs((DATA.guid_directory && DATA.guid_directory[clean]) || clean || "Unknown");
    }

    function parseAddMemberEvent(message) {
      const unique = [];
      for (const guid of extractGuids(message.content_text || "")) {
        if (!unique.includes(guid)) unique.push(guid);
      }

      const senderId = String(message.sender_id || "").toLowerCase();
      const actorId = senderId || unique[0] || "";
      const actorName = stripFcs(message.sender_display_name || resolveGuidLabel(actorId) || "Unknown");
      const addedIds = unique.filter(guid => guid !== actorId);
      const addedNames = addedIds.map(resolveGuidLabel);

      return {
        actorId,
        actorName,
        addedIds,
        addedNames,
      };
    }

    function renderAddMemberEventBody(message, thread) {
      const parsed = parseAddMemberEvent(message);
      const addedLabel = parsed.addedNames.length ? parsed.addedNames.join(" / ") : "Unknown";
      return `
        <div class="call-event" style="--call-bg:#e4e8ed;--call-outline:#66727f;--call-block-bg:#f4f6f8;--call-title:#3e4852;">
          <div class="call-event-header">
            <div class="call-event-title">${escapeHtml(parsed.addedNames.length > 1 ? "Members Added" : "Member Added")}</div>
          </div>
          <div class="call-event-grid">
            <div class="call-event-block"><div class="k">Actor</div><div>${escapeHtml(parsed.actorName)}</div></div>
            <div class="call-event-block"><div class="k">Added</div><div>${escapeHtml(addedLabel)}</div></div>
            <div class="call-event-block"><div class="k">Added Count</div><div>${escapeHtml(String(parsed.addedNames.length || 0))}</div></div>
            <div class="call-event-block"><div class="k">Current Conversation</div><div>${escapeHtml(stripFcs(thread.label || thread.id || ""))}</div></div>
          </div>
          <details class="raw-event" ${state.expandRawEvents ? "open" : ""}>
            <summary>Show raw event payload</summary>
            <div class="body mono">${escapeHtml(message.content_text || "")}</div>
          </details>
        </div>
      `;
    }

    function linkedThreadsForCall(call) {
      const ids = [call.call_id, call.shared_correlation_id]
        .filter(Boolean)
        .map(value => String(value).toLowerCase());
      const targetKey = callKey(call);
      const matches = [];
      for (const thread of DATA.threads || []) {
        if (!thread || thread.category === "call_logs" || String(thread.id || "").toLowerCase() === "48:calllogs") continue;
        const directConversationMatch =
          call.conversation_id &&
          call.conversation_id !== "48:calllogs" &&
          String(thread.id || "").toLowerCase() === String(call.conversation_id || "").toLowerCase();
        const messageMatch = (thread.messages || []).some(message => {
          const linkedKey = messageLinkedCallKey(message);
          if (linkedKey && linkedKey === targetKey) return true;
          const text = String(message.content_text || "").toLowerCase();
          return ids.some(id => id && text.includes(id));
        });
        const syntheticMatch = syntheticCallsForThread(thread).some(candidate => callKey(candidate) === targetKey);
        if ((directConversationMatch || messageMatch || syntheticMatch) && !matches.some(existing => existing.id === thread.id)) {
          matches.push(thread);
        }
      }
      return matches.sort((left, right) => {
        const leftDirect = String(left.id || "").toLowerCase() === String(call.conversation_id || "").toLowerCase() ? 1 : 0;
        const rightDirect = String(right.id || "").toLowerCase() === String(call.conversation_id || "").toLowerCase() ? 1 : 0;
        if (rightDirect !== leftDirect) return rightDirect - leftDirect;
        return threadLastTimestamp(right) - threadLastTimestamp(left);
      });
    }

    function preferredThreadsForCall(call) {
      const matches = linkedThreadsForCall(call);
      const directPairKeys = new Set(
        matches
          .filter(thread => thread.category === "chat_space" && (thread.participants || []).length === 2)
          .map(thread => participantNamePairKey(thread.participants || []))
          .filter(Boolean)
      );
      return matches.filter(thread => {
        const pairKey = participantNamePairKey(thread.participants || []);
        if (!pairKey || !directPairKeys.has(pairKey)) return true;
        return thread.category === "chat_space";
      });
    }

    function callJumpTargets(call) {
      const targets = [];
      const targetKey = callKey(call);
      for (const thread of preferredThreadsForCall(call)) {
        const hasEventTarget = threadTimelineItems(thread).some(message => messageLinkedCallKey(message) === targetKey);
        const hasMessageTarget = hasNearbyCuratedMessages(thread, call);
        const focusTimestamp = call.start_time || call.end_time || "";
        if (hasMessageTarget) {
          targets.push({
            thread,
            focusTimestamp,
            focusCallKey: "",
            viewFilter: "all",
          });
        } else {
          targets.push({
            thread,
            focusTimestamp,
            focusCallKey: hasEventTarget ? targetKey : "",
            viewFilter: "all",
          });
        }
        if (hasEventTarget) {
          targets.push({
            thread,
            focusTimestamp,
            focusCallKey: targetKey,
            viewFilter: "events",
          });
        }
      }
      const seen = new Set();
      return targets.filter(target => {
        const key = `${target.thread.id}|${target.viewFilter}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }

    function callJumpLabel(target) {
      const name = stripFcs(target.thread.label || target.thread.id || "Conversation");
      const when = fmt(target.focusTimestamp);
      const parts = [name];
      if (when && when !== "[no_time]") parts.push(`at ${when}`);
      parts.push(`[${prettyCategory(target.thread.category)} / ${target.viewFilter === "all" ? "All" : target.viewFilter === "messages" ? "Messages" : "Events"}]`);
      return parts.join(" ");
    }

    function messageDisplaySender(message) {
      if (message.message_type === "ThreadActivity/AddMember") {
        return parseAddMemberEvent(message).actorName || "Unknown";
      }
      return stripFcs(message.sender_display_name || message.sender_id || "Unknown");
    }

    function messagePassesFilter(message) {
      if (state.messageViewFilter === "messages") return message.quality !== "event";
      if (state.messageViewFilter === "events") return message.quality === "event";
      return true;
    }

    function hasThreadDateRange() {
      return Boolean(state.threadDateFrom || state.threadDateTo);
    }

    function dateBoundaryValue(value, isEnd) {
      if (!value) return null;
      const parsed = new Date(`${value}T${isEnd ? "23:59:59.999" : "00:00:00.000"}`);
      const time = parsed.getTime();
      return Number.isNaN(time) ? null : time;
    }

    function normalizedThreadDateRange() {
      const start = dateBoundaryValue(state.threadDateFrom, false);
      const end = dateBoundaryValue(state.threadDateTo, true);
      if (start === null && end === null) return { start: null, end: null };
      if (start !== null && end !== null && start > end) {
        return { start: end, end: start };
      }
      return { start, end };
    }

    function messagePassesDateRange(message) {
      const { start, end } = normalizedThreadDateRange();
      if (start === null && end === null) return true;
      const current = timeValue(message.timestamp);
      if (!Number.isFinite(current)) return false;
      if (start !== null && current < start) return false;
      if (end !== null && current > end) return false;
      return true;
    }

    function callSearchText(call) {
      const parts = [
        callLabel(call),
        call.call_id,
        call.shared_correlation_id,
        displayCallDirection(call),
        displayCallState(call),
        displayCallStatus(call),
        call.call_type,
        call.conversation_id,
        call.summary_text,
        call.originator_display_name,
        call.originator_id,
        call.target_display_name,
        call.target_id,
      ];
      return parts.filter(Boolean).join(" ").toLowerCase();
    }

    function callDuration(call) {
      let seconds = callDurationSeconds(call);
      if (seconds === null) return "";
      const hours = Math.floor(seconds / 3600);
      seconds -= hours * 3600;
      const minutes = Math.floor(seconds / 60);
      seconds -= minutes * 60;
      const parts = [];
      if (hours) parts.push(`${hours}h`);
      if (minutes || hours) parts.push(`${minutes}m`);
      parts.push(`${seconds}s`);
      return parts.join(" ");
    }

    function filteredCalls() {
      return DATA.calls
        .filter(call => {
          if (state.callDirection === "missed") {
            if (displayCallState(call) !== "missed") return false;
          } else if (state.callDirection && displayCallDirection(call) !== state.callDirection) {
            return false;
          }
          if (state.callSearch && !callSearchText(call).includes(state.callSearch)) return false;
          return true;
        })
        .sort((left, right) => (right.start_time || "").localeCompare(left.start_time || ""));
    }

    function searchMessageText(thread, message) {
      const linkedCall = findLinkedCall(message);
      const parts = [
        thread.label,
        thread.id,
        ...(thread.participants || []),
        messageDisplaySender(message),
        message.message_type,
        message.content_text,
      ];
      if (linkedCall) {
        parts.push(
          callLabel(linkedCall),
          linkedCall.call_type,
          linkedCall.call_state,
          linkedCall.direction,
          linkedCall.summary_text,
          linkedCall.originator_display_name,
          linkedCall.target_display_name,
        );
      }
      if (message.message_type === "ThreadActivity/AddMember") {
        const parsed = parseAddMemberEvent(message);
        parts.push(parsed.actorName, ...(parsed.addedNames || []));
      }
      return parts.filter(Boolean).join(" ").toLowerCase();
    }

    function includeMessageInSearchResults(message) {
      if (!message) return false;
      if (message.message_type === "ThreadActivity/AddMember") return false;
      if (message.message_type === "Call/History") return false;
      if (message.synthetic_call) return false;
      return true;
    }

    function searchResultGroups() {
      const groups = [];
      if (!state.messageSearch) return groups;

      for (const thread of filteredThreads()) {
        const timeline = threadTimelineItems(thread)
          .slice()
          .sort((left, right) => {
            const timeDiff = timeValue(right.timestamp) - timeValue(left.timestamp);
            if (timeDiff !== 0) return timeDiff;
            return (right.id || "").localeCompare(left.id || "");
          })
          .filter(messagePassesFilter)
          .filter(includeMessageInSearchResults);

        const hits = [];
        for (let index = 0; index < timeline.length; index += 1) {
          if (textMatchesSearch(searchMessageText(thread, timeline[index]))) {
            hits.push(index);
          }
        }
        if (!hits.length) continue;

        const clusters = [];
        for (const hitIndex of hits) {
          const start = Math.max(0, hitIndex - 1);
          const end = Math.min(timeline.length - 1, hitIndex + 1);
          const previous = clusters[clusters.length - 1];
          if (previous && start <= previous.end + 1) {
            previous.end = Math.max(previous.end, end);
            previous.hitCount += 1;
          } else {
            clusters.push({ start, end, hitCount: 1, hitIndex });
          }
        }

        for (const cluster of clusters) {
          const focusMessage = timeline[cluster.hitIndex];
          groups.push({
            thread,
            hitCount: cluster.hitCount,
            focusTimestamp: focusMessage.timestamp || "",
            focusCallKey: messageLinkedCallKey(focusMessage),
            viewFilter: "all",
            messages: timeline.slice(cluster.start, cluster.end + 1),
          });
        }
      }

      return groups.sort((left, right) => {
        const timeDiff = timeValue(right.focusTimestamp) - timeValue(left.focusTimestamp);
        if (timeDiff !== 0) return timeDiff;
        return (left.thread.label || left.thread.id || "").localeCompare(right.thread.label || right.thread.id || "");
      });
    }

    function renderMessageList() {
      const rows = filteredThreads();
      sidebarCount.textContent = `${rows.length} conversations`;
      if (!state.threadId || !rows.some(thread => thread.id === state.threadId)) {
        state.threadId = rows[0] ? rows[0].id : null;
      }
      threadList.innerHTML = rows.map(thread => `
        <div class="list-row ${thread.id === state.threadId ? "active" : ""}" data-id="${escapeHtml(thread.id)}">
          <div><strong>${escapeHtml(stripFcs(thread.label || thread.id))}</strong></div>
          <div class="meta">
            <span>${escapeHtml(prettyCategory(thread.category))}</span>
            <span>${escapeHtml(fmt(threadLastTimestampRaw(thread)))}</span>
          </div>
          <div class="preview">${escapeHtml(latestThreadPreview(thread))}</div>
        </div>
      `).join("") || `<div class="empty">No conversations match the current filters.</div>`;
      [...threadList.querySelectorAll(".list-row")].forEach(element => {
        element.addEventListener("click", () => {
          state.threadId = element.dataset.id;
          state.threadDateFrom = "";
          state.threadDateTo = "";
          state.focusCallKey = null;
          state.focusTimestamp = null;
          renderMessageList();
          renderContent();
          scrollMainToTop();
        });
      });
    }

    function renderCallList() {
      const rows = filteredCalls();
      sidebarCount.textContent = `${rows.length} calls`;
      if (!state.callKey || !rows.some(call => callKey(call) === state.callKey)) {
        state.callKey = rows[0] ? callKey(rows[0]) : null;
      }
      callList.innerHTML = rows.map(call => `
        <div class="list-row ${callKey(call) === state.callKey ? "active" : ""}" data-key="${escapeHtml(callKey(call))}">
          <div><strong>${escapeHtml(stripFcs(callLabel(call)))}</strong></div>
          <div class="meta">
            <span>${escapeHtml(displayCallStatus(call))}</span>
            <span>${escapeHtml(fmt(call.start_time))}</span>
          </div>
        </div>
      `).join("") || `<div class="empty">No calls match the current filters.</div>`;
      [...callList.querySelectorAll(".list-row")].forEach(element => {
        element.addEventListener("click", () => {
          state.callKey = element.dataset.key;
          renderCallList();
          renderContent();
        });
      });
    }

    function renderSearchResultsPanel() {
        const groups = searchResultGroups();
        const totalMatches = groups.reduce((sum, group) => sum + group.hitCount, 0);
        contentPanel.innerHTML = `
        <div class="title">
          <h2>Search Results</h2>
          <div class="chip">${escapeHtml(String(totalMatches))} matches</div>
        </div>
        <div class="subtle">Showing grouped timeline matches for <strong>${escapeHtml(state.messageSearch)}</strong> with +/- 1 surrounding timeline item.</div>
        <div class="search-results">
          ${groups.map((group, index) => `
            <div class="search-group">
              <div class="search-group-header">
                <div>
                  <div><strong>${escapeHtml(stripFcs(group.thread.label || group.thread.id))}</strong></div>
                  <div class="search-group-meta">${escapeHtml(prettyCategory(group.thread.category))} | ${escapeHtml(fmt(group.focusTimestamp))} | ${escapeHtml(String(group.hitCount))} match${group.hitCount === 1 ? "" : "es"}</div>
                </div>
                <button type="button" class="call-link open-search-thread" data-thread-id="${escapeHtml(group.thread.id)}" data-focus-time="${escapeHtml(group.focusTimestamp || "")}" data-focus-call-key="${escapeHtml(group.focusCallKey || "")}" data-view-filter="${escapeHtml(group.viewFilter || "all")}">Open In Chats</button>
              </div>
              <div class="messages">
                ${group.messages.map(message => `
                  <div class="msg ${message.quality === "event" ? "event" : ""} ${textMatchesSearch(searchMessageText(group.thread, message)) ? "search-match" : ""}" data-call-key="${escapeHtml(messageLinkedCallKey(message))}" data-timestamp="${escapeHtml(String(timeValue(message.timestamp)))}" style="--msg-accent:${escapeHtml(messageAccent(message))}">
                    <div class="head">
                      <span class="sender">${highlightSearchHtml(messageDisplaySender(message))}</span>
                      <span>${escapeHtml(fmt(message.timestamp))}</span>
                    </div>
                    <div class="head">
                      <span>${highlightSearchHtml(message.message_type || "")}</span>
                      <span>${escapeHtml(message.quality || "")}</span>
                    </div>
                    ${(message.synthetic_call || message.message_type === "Event/Call")
                      ? renderCallEventBody(message)
                      : message.message_type === "ThreadActivity/AddMember"
                        ? renderAddMemberEventBody(message, group.thread)
                        : `<div class="body">${highlightSearchHtml(message.content_text || "")}</div>`}
                  </div>
                `).join("")}
              </div>
            </div>
            ${index < groups.length - 1 ? `<div class="search-group-divider"></div>` : ``}
          `).join("") || `<div class="empty">No timeline items match the current search.</div>`}
        </div>
      `;

      [...contentPanel.querySelectorAll(".open-search-thread")].forEach(element => {
        element.addEventListener("click", () => {
          state.view = "messages";
          state.threadId = element.dataset.threadId;
          state.messageViewFilter = element.dataset.viewFilter || "all";
          state.threadDateFrom = "";
          state.threadDateTo = "";
          state.focusTimestamp = element.dataset.focusTime || null;
          state.focusCallKey = element.dataset.focusCallKey || null;
          state.messageSearch = "";
          messageSearch.value = "";
          renderView();
        });
      });

      [...contentPanel.querySelectorAll(".open-call-record")].forEach(element => {
        element.addEventListener("click", () => {
          state.view = "calls";
          state.callKey = element.dataset.callKey;
          renderView();
        });
      });
    }

    function renderThreadPanel() {
      const thread = DATA.threads.find(item => item.id === state.threadId);
      if (!thread) {
        contentPanel.innerHTML = `<div class="empty">Select a conversation to view its data.</div>`;
        return;
      }
      const meta = thread.metadata || {};
      const timelineMessages = threadTimelineItems(thread);
      const allMessages = timelineMessages.slice().sort((left, right) => {
        const timeDiff = timeValue(right.timestamp) - timeValue(left.timestamp);
        if (timeDiff !== 0) return timeDiff;
        return (right.id || "").localeCompare(left.id || "");
      });
      const rangeMessages = allMessages.filter(messagePassesDateRange);
      const syntheticCallCount = allMessages.filter(message => message.synthetic_call).length;
      const messages = rangeMessages.filter(messagePassesFilter);
      const firstMessage = messages.length ? messages[messages.length - 1] : null;
      const lastMessage = messages[0] || null;
      const eventCount = messages.filter(message => message.quality === "event").length;
      const curatedCount = messages.filter(message => message.quality !== "event").length;
      const meeting = thread.meeting || {};
      const showCompactChatMeta = isChatBasedThread(thread);
      contentPanel.innerHTML = `
        <div class="title">
          <h2>${escapeHtml(stripFcs(thread.label || thread.id))}</h2>
          <div class="chip">${escapeHtml(prettyCategory(thread.category))} | ${allMessages.length} timeline items</div>
        </div>
        <div class="chips">
          ${(thread.participants || []).map(name => `<div class="chip">${escapeHtml(stripFcs(name))}</div>`).join("")}
        </div>
        <div class="detail-toolbar">
          <div class="toolbar-row">
            <div class="segmented">
              <button type="button" class="seg-btn message-filter ${state.messageViewFilter === "all" ? "active" : ""}" data-filter="all">All</button>
              <button type="button" class="seg-btn message-filter ${state.messageViewFilter === "messages" ? "active" : ""}" data-filter="messages">Messages</button>
              <button type="button" class="seg-btn message-filter ${state.messageViewFilter === "events" ? "active" : ""}" data-filter="events">Events</button>
            </div>
            <div class="call-event-actions">
              <button type="button" class="call-link jump-newest">Newest</button>
              <button type="button" class="call-link jump-oldest">Oldest</button>
              ${thread.csv_path ? `<a class="call-link" href="${escapeHtml(thread.csv_path)}" target="_blank" rel="noopener">Open CSV</a>` : ``}
              <button type="button" class="call-link copy-thread-id">Copy Thread Id</button>
              <button type="button" class="call-link toggle-raw-events">${state.expandRawEvents ? "Hide Raw Payloads" : "Show Raw Payloads"}</button>
            </div>
          </div>
          <div class="range-controls">
            <label class="range-field">
              <span>From</span>
              <input type="date" class="thread-date-from" value="${escapeHtml(state.threadDateFrom)}">
            </label>
            <label class="range-field">
              <span>To</span>
              <input type="date" class="thread-date-to" value="${escapeHtml(state.threadDateTo)}">
            </label>
            <button type="button" class="call-link clear-date-range"${hasThreadDateRange() ? "" : " disabled"}>Clear Range</button>
          </div>
          <div class="toolbar-divider"></div>
        </div>
        <div class="stat-strip">
          <div class="stat-card"><div class="k">Visible</div><div class="v">${escapeHtml(String(messages.length))}</div></div>
          <div class="stat-card"><div class="k">Curated</div><div class="v">${escapeHtml(String(curatedCount))}</div></div>
          <div class="stat-card"><div class="k">Events</div><div class="v">${escapeHtml(String(eventCount))}</div></div>
          <div class="stat-card"><div class="k">Last Activity</div><div class="v">${escapeHtml(fmt(lastMessage ? lastMessage.timestamp : null))}</div></div>
        </div>
        <div class="meta-grid">
          <div class="meta-block"><div class="k">Thread Id</div><div class="mono">${escapeHtml(thread.id)}</div></div>
          <div class="meta-block"><div class="k">Metadata Quality</div><div>${escapeHtml(thread.metadata_quality || "")}</div></div>
          <div class="meta-block"><div class="k">Participants</div><div>${(thread.participants || []).length}</div></div>
          <div class="meta-block"><div class="k">Known Metadata</div><div>${Object.keys(meta).length}</div></div>
          <div class="meta-block"><div class="k">Merged Calls</div><div>${escapeHtml(String(syntheticCallCount))}</div></div>
          ${showCompactChatMeta ? "" : `<div class="meta-block"><div class="k">Meeting Subject</div><div>${escapeHtml(meeting.subject || "")}</div></div>`}
          ${showCompactChatMeta ? "" : `<div class="meta-block"><div class="k">Meeting Window</div><div>${escapeHtml(meeting.startTime ? `${fmt(meeting.startTime)} to ${fmt(meeting.endTime)}` : "")}</div></div>`}
          <div class="meta-block"><div class="k">First Activity</div><div>${escapeHtml(fmt(firstMessage ? firstMessage.timestamp : null))}</div></div>
          ${showCompactChatMeta ? "" : `<div class="meta-block"><div class="k">Latest Preview</div><div>${escapeHtml(latestThreadPreview(thread))}</div></div>`}
        </div>
        <div class="section-divider">
          <div class="section-label">Timeline</div>
        </div>
        <div class="messages">
          ${messages.map(message => `
            <div class="msg ${message.quality === "event" ? "event" : ""} ${state.focusCallKey && messageLinkedCallKey(message) === state.focusCallKey ? "focused" : ""}" data-call-key="${escapeHtml(messageLinkedCallKey(message))}" data-timestamp="${escapeHtml(String(timeValue(message.timestamp)))}" style="--msg-accent:${escapeHtml(messageAccent(message))}">
              <div class="head">
                <span class="sender">${escapeHtml(messageDisplaySender(message))}</span>
                <span>${escapeHtml(fmt(message.timestamp))}</span>
              </div>
              <div class="head">
                <span>${escapeHtml(message.message_type || "")}</span>
                <span>${escapeHtml(message.quality || "")}</span>
              </div>
              ${(message.synthetic_call || message.message_type === "Event/Call")
                ? renderCallEventBody(message)
                : message.message_type === "ThreadActivity/AddMember"
                  ? renderAddMemberEventBody(message, thread)
                  : `<div class="body">${escapeHtml(message.content_text || "")}</div>`}
            </div>
          `).join("") || `<div class="empty">${hasThreadDateRange() ? "No timeline items match the current date range and filters." : "No messages in this conversation."}</div>`}
        </div>
      `;
      [...contentPanel.querySelectorAll(".message-filter")].forEach(element => {
        element.addEventListener("click", () => {
          state.messageViewFilter = element.dataset.filter;
          renderThreadPanel();
        });
      });
      const threadDateFrom = contentPanel.querySelector(".thread-date-from");
      if (threadDateFrom) {
        threadDateFrom.addEventListener("change", event => {
          state.threadDateFrom = event.target.value || "";
          renderThreadPanel();
        });
      }
      const threadDateTo = contentPanel.querySelector(".thread-date-to");
      if (threadDateTo) {
        threadDateTo.addEventListener("change", event => {
          state.threadDateTo = event.target.value || "";
          renderThreadPanel();
        });
      }
      const clearDateRange = contentPanel.querySelector(".clear-date-range");
      if (clearDateRange) {
        clearDateRange.addEventListener("click", () => {
          state.threadDateFrom = "";
          state.threadDateTo = "";
          renderThreadPanel();
        });
      }
      const copyThreadId = contentPanel.querySelector(".copy-thread-id");
      if (copyThreadId) {
        copyThreadId.addEventListener("click", () => copyText(thread.id, "Thread id copied"));
      }
      const toggleRaw = contentPanel.querySelector(".toggle-raw-events");
      if (toggleRaw) {
        toggleRaw.addEventListener("click", () => {
          state.expandRawEvents = !state.expandRawEvents;
          renderThreadPanel();
        });
      }
      const jumpNewest = contentPanel.querySelector(".jump-newest");
      if (jumpNewest) {
        jumpNewest.addEventListener("click", () => {
          const firstMessageCard = contentPanel.querySelector(".messages .msg");
          if (firstMessageCard) {
            firstMessageCard.scrollIntoView({ block: "center", behavior: "smooth" });
          }
        });
      }
      const jumpOldest = contentPanel.querySelector(".jump-oldest");
      if (jumpOldest) {
        jumpOldest.addEventListener("click", () => {
          const cards = contentPanel.querySelectorAll(".messages .msg");
          const lastMessageCard = cards[cards.length - 1];
          if (lastMessageCard) {
            lastMessageCard.scrollIntoView({ block: "center", behavior: "smooth" });
          }
        });
      }
      [...contentPanel.querySelectorAll(".open-call-record")].forEach(element => {
        element.addEventListener("click", () => {
          state.view = "calls";
          state.callKey = element.dataset.callKey;
          renderView();
        });
      });
      focusThreadPosition();
    }

    function focusThreadPosition() {
      if (!state.focusCallKey && !state.focusTimestamp) return;
      let target = null;
      if (state.focusCallKey) {
        target = contentPanel.querySelector(`.msg[data-call-key="${CSS.escape(state.focusCallKey)}"]`);
      }
      if (!target && state.focusTimestamp) {
        const focusTime = timeValue(state.focusTimestamp);
        let bestDiff = Number.POSITIVE_INFINITY;
        for (const element of contentPanel.querySelectorAll(".msg[data-timestamp]")) {
          const current = Number(element.dataset.timestamp || Number.NEGATIVE_INFINITY);
          const diff = Math.abs(current - focusTime);
          if (diff < bestDiff) {
            bestDiff = diff;
            target = element;
          }
        }
        if (target) target.classList.add("focused");
      }
      if (target) {
        target.scrollIntoView({ block: "center", behavior: "smooth" });
      }
    }

    function renderCallPanel() {
      const call = DATA.calls.find(item => callKey(item) === state.callKey);
        if (!call) {
          contentPanel.innerHTML = `<div class="empty">Select a call to view its data.</div>`;
          return;
        }
      const relatedThreads = preferredThreadsForCall(call);
      const jumpTargets = callJumpTargets(call);
      contentPanel.innerHTML = `
        <div class="title">
          <h2>${escapeHtml(stripFcs(callLabel(call)))}</h2>
          <div class="chip">${escapeHtml(displayCallStatus(call))}</div>
        </div>
        <div class="chips">
          <div class="chip">${escapeHtml(call.call_type || "")}</div>
          <div class="chip">${escapeHtml(call.quality || "")}</div>
        </div>
        <div class="detail-toolbar">
          <div class="segmented">
            <div class="chip">${escapeHtml(callDuration(call) || "Duration unavailable")}</div>
          </div>
          <div class="call-event-actions">
            <a class="call-link" href="teams_ccl_csv_v1/call_history.csv" target="_blank" rel="noopener">Open Call CSV</a>
            <button type="button" class="call-link copy-call-id">Copy Call Id</button>
          </div>
        </div>
        <div class="stat-strip">
          <div class="stat-card"><div class="k">Start</div><div class="v">${escapeHtml(fmt(call.start_time))}</div></div>
          <div class="stat-card"><div class="k">End</div><div class="v">${escapeHtml(fmt(call.end_time))}</div></div>
          <div class="stat-card"><div class="k">Duration</div><div class="v">${escapeHtml(callDuration(call) || "Unavailable")}</div></div>
          <div class="stat-card"><div class="k">Linked Conversations</div><div class="v">${escapeHtml(String(relatedThreads.length))}</div></div>
        </div>
        <div class="meta-grid">
          <div class="meta-block"><div class="k">Start Time</div><div>${escapeHtml(fmt(call.start_time))}</div></div>
          <div class="meta-block"><div class="k">Connect Time</div><div>${escapeHtml(fmt(call.connect_time))}</div></div>
          <div class="meta-block"><div class="k">End Time</div><div>${escapeHtml(fmt(call.end_time))}</div></div>
          <div class="meta-block"><div class="k">Direction</div><div>${escapeHtml(displayCallDirection(call))}</div></div>
          <div class="meta-block"><div class="k">State</div><div>${escapeHtml(displayCallState(call))}</div></div>
          <div class="meta-block"><div class="k">Type</div><div>${escapeHtml(call.call_type || "")}</div></div>
          <div class="meta-block"><div class="k">Originator</div><div>${escapeHtml(stripFcs(call.originator_display_name || call.originator_id || ""))}</div></div>
          <div class="meta-block"><div class="k">Target</div><div>${escapeHtml(stripFcs(call.target_display_name || call.target_id || ""))}</div></div>
          <div class="meta-block"><div class="k">Call Id</div><div class="mono">${escapeHtml(call.call_id || "")}</div></div>
          <div class="meta-block"><div class="k">Shared Correlation Id</div><div class="mono">${escapeHtml(call.shared_correlation_id || "")}</div></div>
          <div class="meta-block"><div class="k">Participants</div><div>${escapeHtml(stripFcs(callLabel(call)))}</div></div>
          <div class="meta-block"><div class="k">Summary</div><div>${escapeHtml(call.summary_text || "")}</div></div>
        </div>
        <div class="section-label">Open Conversation At Call Time</div>
        ${jumpTargets.length
          ? `<div class="jump-list">
              ${jumpTargets.map(target => `<button type="button" class="call-link open-thread-record" data-thread-id="${escapeHtml(target.thread.id)}" data-focus-time="${escapeHtml(target.focusTimestamp || "")}" data-focus-call-key="${escapeHtml(target.focusCallKey || "")}" data-view-filter="${escapeHtml(target.viewFilter || "all")}">${escapeHtml(callJumpLabel(target))}</button>`).join("")}
            </div>`
          : `<div class="subtle">No linked conversation messages were found for this call.</div>`}
      `;
      const copyCallId = contentPanel.querySelector(".copy-call-id");
      if (copyCallId) {
        copyCallId.addEventListener("click", () => copyText(call.call_id || "", "Call id copied"));
      }
      [...contentPanel.querySelectorAll(".open-thread-record")].forEach(element => {
        element.addEventListener("click", () => {
          state.view = "messages";
          state.threadId = element.dataset.threadId;
          state.messageViewFilter = element.dataset.viewFilter || "all";
          state.threadDateFrom = "";
          state.threadDateTo = "";
          state.focusTimestamp = element.dataset.focusTime || null;
          state.focusCallKey = element.dataset.focusCallKey || null;
          renderView();
        });
      });
    }

    function renderContent() {
      if (state.view === "calls") {
        renderCallPanel();
        return;
      }
      if (state.messageSearch) {
        renderSearchResultsPanel();
        return;
      }
      renderThreadPanel();
    }

    function renderView() {
      const showingMessages = state.view === "messages";
      viewMessages.classList.toggle("active", showingMessages);
      viewCalls.classList.toggle("active", !showingMessages);
      messagesTools.classList.toggle("hidden", !showingMessages);
      callsTools.classList.toggle("hidden", showingMessages);
      threadList.classList.toggle("hidden", !showingMessages);
      callList.classList.toggle("hidden", showingMessages);
      if (showingMessages) {
        renderMessageList();
      } else {
        renderCallList();
      }
      renderContent();
    }

    viewMessages.addEventListener("click", () => {
      state.view = "messages";
      state.focusCallKey = null;
      state.focusTimestamp = null;
      renderView();
    });

    viewCalls.addEventListener("click", () => {
      state.view = "calls";
      renderView();
    });

    messageSearch.addEventListener("input", event => {
      state.messageSearch = event.target.value.trim().toLowerCase();
      renderMessageList();
      renderContent();
      scrollMainToTop();
    });

    categoryFilter.addEventListener("change", event => {
      state.category = event.target.value;
      renderMessageList();
      renderContent();
    });

    callSearch.addEventListener("input", event => {
      state.callSearch = event.target.value.trim().toLowerCase();
      renderCallList();
      renderContent();
    });

    callDirectionFilter.addEventListener("change", event => {
      state.callDirection = event.target.value;
      renderCallList();
      renderContent();
    });

    renderView();
  </script>
</body>
</html>
"""


def load_export(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def build_browser(export_data: dict, output_path: Path) -> None:
    threads = []
    for thread in export_data.get("threads") or []:
        row = dict(thread)
        row["csv_path"] = thread_csv_path(thread)
        threads.append(row)

    payload = {
        "summary": export_data.get("summary") or {},
        "profile": export_data.get("profile") or {},
        "threads": threads,
        "calls": export_data.get("calls") or [],
        "guid_directory": export_data.get("guid_directory") or {},
    }
    html = HTML_TEMPLATE.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
    output_path.write_text(html, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build a single-file browser viewer for the Teams CCL export.")
    parser.add_argument("--input", default="teams_ccl_canonical_v1.json", help="Path to the canonical CCL JSON.")
    parser.add_argument("--output", default="teams_ccl_browser_v1.html", help="Path to write the browser HTML.")
    args = parser.parse_args()

    export_data = load_export(Path(args.input).resolve())
    build_browser(export_data, Path(args.output).resolve())


if __name__ == "__main__":
    main()
