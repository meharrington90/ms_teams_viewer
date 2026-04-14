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
  <title>Michaelsoft Teams</title>
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
    *,
    *::before,
    *::after { box-sizing: border-box; }
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
      height: 100vh;
      overflow: hidden;
    }
    .app {
      display: grid;
      grid-template-columns: 360px minmax(0, 1fr);
      min-height: 0;
      height: 100vh;
      max-width: 100vw;
      overflow: hidden;
    }
    .sidebar {
      border-right: 1px solid var(--line);
      background: rgba(255,250,242,.88);
      backdrop-filter: blur(8px);
      padding: 20px;
      display: flex;
      flex-direction: column;
      gap: 0;
      height: 100vh;
      min-height: 0;
      overflow: hidden;
      min-width: 0;
    }
    .main {
      padding: 20px 24px 32px;
      overflow: auto;
      min-height: 0;
      min-width: 0;
    }
    h1, h2, h3 { margin: 0; font-weight: 600; }
    .title {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 12px;
      min-width: 0;
    }
    .title > * {
      min-width: 0;
    }
    .title h1,
    .title h2,
    .title h3 {
      overflow-wrap: anywhere;
      word-break: break-word;
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
      gap: 6px;
      margin-bottom: 10px;
    }
    .tab {
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 5px 12px;
      background: rgba(255,255,255,.72);
      color: var(--ink);
      font: inherit;
      cursor: pointer;
      min-height: 40px;
      display: flex;
      align-items: center;
      justify-content: center;
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
      gap: 8px;
      margin-bottom: 10px;
      align-items: stretch;
    }
    .hidden { display: none !important; }
    input, select {
      width: 100%;
      min-width: 0;
      max-width: 100%;
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 8px 12px;
      background: #fff;
      color: var(--ink);
      font: inherit;
      min-height: 46px;
      line-height: 1.25;
      box-shadow: inset 0 1px 0 rgba(255,255,255,.55);
    }
    input:focus, select:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px rgba(29,111,109,.12);
    }
    .toolbar > * {
      min-width: 0;
    }
    .chip-row {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 0;
    }
    .sidebar-top {
      display: grid;
      gap: 0;
      min-width: 0;
      padding-bottom: 12px;
    }
    .sidebar-divider {
      height: 1px;
      background: var(--line);
      margin-right: 2px;
      flex: 0 0 auto;
    }
    .sidebar-list-wrap {
      flex: 1 1 auto;
      min-height: 0;
      overflow: auto;
      padding-top: 12px;
      scrollbar-gutter: stable;
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
      gap: 6px;
    }
    .list-row {
      padding: 9px 12px;
      border-radius: 14px;
      border: 1px solid transparent;
      background: rgba(255,255,255,.7);
      cursor: pointer;
      transition: transform .15s ease, border-color .15s ease, background .15s ease;
      min-width: 0;
      overflow: hidden;
    }
    .list-row:hover { transform: translateY(-1px); border-color: var(--line); }
    .list-row.active { border-color: var(--accent); background: var(--accent-soft); }
    .list-row > div {
      min-width: 0;
    }
    .list-row strong {
      display: block;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    .list-row .meta {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 6px;
      align-items: baseline;
      color: var(--muted);
      font-size: 12px;
      margin-top: 4px;
      min-width: 0;
    }
    .list-row .meta span {
      min-width: 0;
    }
    .list-row .meta span:last-child {
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
      justify-self: end;
    }
    .preview {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.3;
      margin-top: 6px;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    .panel {
      padding: 18px 20px;
      margin-bottom: 16px;
      min-width: 0;
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
    .sidebar-range-controls {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      align-items: end;
      width: 100%;
    }
    .sidebar-range-controls .range-field {
      flex: 0 1 calc((100% - 6px) / 2);
      width: calc((100% - 6px) / 2);
      min-width: 0;
    }
    .sidebar-range-controls .range-action {
      flex: 0 0 100%;
      width: 100%;
      min-height: 40px;
      border-radius: 12px;
    }
    .toolbar-divider {
      height: 1px;
      background: var(--line);
      width: 100%;
    }
    .range-field {
      display: grid;
      gap: 3px;
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
    .participant-cloud {
      margin-top: 10px;
    }
    .expandable-shell {
      display: grid;
      gap: 8px;
    }
    .expandable-list {
      margin-top: 0;
      align-items: flex-start;
      min-width: 0;
      position: relative;
    }
    .expandable-list.collapsed {
      max-height: 76px;
      overflow: hidden;
    }
    .expandable-list.expanded {
      max-height: none;
      overflow: visible;
    }
    .expandable-list.collapsed::after {
      content: "";
      position: absolute;
      left: 0;
      right: 0;
      bottom: 0;
      height: 18px;
      background: linear-gradient(180deg, rgba(255,250,242,0), rgba(255,250,242,.96));
      pointer-events: none;
    }
    .call-event-block .expandable-list.collapsed::after {
      background: linear-gradient(180deg, rgba(255,255,255,0), rgba(255,255,255,.96));
    }
    .toggle-expandable-list {
      justify-self: start;
    }
    .participant-chip {
      font: inherit;
      line-height: 1.2;
    }
    .participant-chip-link {
      background: #eaf1ff;
      border-color: #355cde;
      color: #1d3f9c;
      cursor: pointer;
    }
    .participant-chip-link:hover {
      background: #dfe9ff;
    }
    .participant-chip-static {
      background: #f1f3f5;
      border-color: #d5dbe3;
      color: #6a7582;
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
    .msg .body.attachment-body {
      display: grid;
      gap: 10px;
      white-space: normal;
    }
    .msg .body-text {
      white-space: pre-wrap;
      line-height: 1.4;
      word-break: break-word;
    }
    .attachment-stack {
      display: grid;
      gap: 10px;
    }
    .attachment-card {
      border: 1px solid var(--line);
      border-radius: 12px;
      background: rgba(255,255,255,.72);
      padding: 10px 12px;
      display: grid;
      gap: 6px;
    }
    .attachment-meta {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 10px;
      flex-wrap: wrap;
    }
    .attachment-kind {
      color: var(--muted);
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: .06em;
    }
    .attachment-name {
      font-weight: 600;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    .attachment-actions {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .attachment-note {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.35;
      margin-top: 4px;
    }
    details.msg-collapsible {
      padding: 0;
      overflow: hidden;
    }
    details.msg-collapsible > summary {
      cursor: pointer;
      list-style-position: outside;
      padding: 12px 14px;
    }
    details.msg-collapsible > summary::-webkit-details-marker {
      margin-right: 8px;
    }
    details.msg-collapsible[open] > summary {
      padding-bottom: 10px;
    }
    details.msg-collapsible .msg-body-wrap {
      padding: 0 14px 12px;
      border-top: 1px solid rgba(31,42,43,.08);
    }
    .msg-summary-note {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.35;
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
    .call-event-block.wide {
      grid-column: 1 / -1;
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
      display: inline-flex;
      align-items: center;
      justify-content: center;
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
      min-width: 0;
    }
    .meta-block {
      background: rgba(255,255,255,.82);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 10px 12px;
      min-width: 0;
    }
    .meta-block.wide {
      grid-column: 1 / -1;
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
      body {
        height: auto;
        overflow: auto;
      }
      .app {
        grid-template-columns: 1fr;
        height: auto;
        min-height: 100vh;
        overflow: visible;
      }
      .sidebar {
        height: auto;
        min-height: auto;
        border-right: 0;
        border-bottom: 1px solid var(--line);
      }
      .sidebar-list-wrap {
        overflow: visible;
        padding-top: 12px;
      }
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
      <div class="sidebar-top">
        <div class="title">
          <h1>Michaelsoft Teams</h1>
        </div>
        <div class="view-tabs">
          <button id="viewMessages" class="tab active" type="button">Messages</button>
          <button id="viewCalls" class="tab" type="button">Calls</button>
        </div>
        <div id="messagesTools" class="toolbar">
          <input id="messageSearch" type="search" placeholder="Search Conversations and Messages...">
          <select id="categoryFilter">
            <option value="">All Message Types</option>
            <option value="chat_space">Chats</option>
            <option value="team_chat">Meetings</option>
            <option value="thread">Groups</option>
          </select>
          <div class="sidebar-range-controls">
            <label class="range-field">
              <span>From</span>
              <input id="messageDateFrom" type="date" value="">
            </label>
            <label class="range-field">
              <span>To</span>
              <input id="messageDateTo" type="date" value="">
            </label>
            <button id="clearMessageDateRange" type="button" class="call-link range-action">Clear Range</button>
          </div>
        </div>
        <div id="callsTools" class="toolbar hidden">
          <input id="callSearch" type="search" placeholder="Search Calls and Details...">
          <select id="callGroupFilter">
            <option value="">All Call Types</option>
            <option value="phone">Phone</option>
            <option value="meeting">Meeting</option>
            <option value="group">Group</option>
            <option value="call">Call</option>
          </select>
          <select id="callDirectionFilter">
            <option value="">All Directions</option>
            <option value="inbound">Inbound</option>
            <option value="outbound">Outbound</option>
            <option value="declined">Declined</option>
            <option value="missed">Missed</option>
          </select>
          <div class="sidebar-range-controls">
            <label class="range-field">
              <span>From</span>
              <input id="callDateFrom" type="date" value="">
            </label>
            <label class="range-field">
              <span>To</span>
              <input id="callDateTo" type="date" value="">
            </label>
            <button id="clearCallDateRange" type="button" class="call-link range-action">Clear Range</button>
          </div>
        </div>
        <div class="chip-row">
          <div id="sidebarCount" class="chip"></div>
        </div>
      </div>
      <div class="sidebar-divider"></div>
      <div class="sidebar-list-wrap">
        <div id="threadList" class="list"></div>
        <div id="callList" class="list hidden"></div>
      </div>
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
      messageDateFrom: "",
      messageDateTo: "",
      callSearch: "",
      category: "",
      callGroup: "",
      callDirection: "",
      callDateFrom: "",
      callDateTo: "",
      messageViewFilter: "all",
      threadDateFrom: "",
      threadDateTo: "",
      expandHiddenData: false,
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
    const messageDateFrom = document.getElementById("messageDateFrom");
    const messageDateTo = document.getElementById("messageDateTo");
    const clearMessageDateRange = document.getElementById("clearMessageDateRange");
    const callSearch = document.getElementById("callSearch");
    const callGroupFilter = document.getElementById("callGroupFilter");
    const callDirectionFilter = document.getElementById("callDirectionFilter");
    const callDateFrom = document.getElementById("callDateFrom");
    const callDateTo = document.getElementById("callDateTo");
    const clearCallDateRange = document.getElementById("clearCallDateRange");
    const toast = document.getElementById("toast");
    const MESSAGE_ATTACHMENT_CACHE = new WeakMap();

    const PERSON_NAME_BY_GUID = new Map(
      Object.entries(DATA.guid_directory || {})
        .map(([guid, name]) => [normalizeGuid(guid), stripFcs(name)])
        .filter(([guid, name]) => guid && name)
    );
    const PERSON_GUID_BY_NAME = new Map(
      [...PERSON_NAME_BY_GUID.entries()]
        .map(([guid, name]) => [normalizeName(name), guid])
        .filter(([name, guid]) => name && guid)
    );
    const DIRECT_CHAT_BY_GUID = new Map();
    const DIRECT_CHAT_BY_NAME = new Map();
    const CALLS_BY_ID = new Map();
    const CALLS_BY_SHARED = new Map();
    const CALLS_BY_KEY = new Map();
    const THREAD_BY_ID = new Map();
    const PREFERRED_THREADS_BY_CALL_KEY = new Map();
    const CALL_JUMP_TARGETS_BY_CALL_KEY = new Map();
    const PHONE_CALL_BY_KEY = new Map();
    const CALL_EXTERNAL_PARTICIPANT_COUNT_BY_KEY = new Map();
    const CALL_GROUP_KEY_BY_KEY = new Map();
    const CALLS_BY_THREAD_ID = new Map();
    const CALLS_BY_GUID_PAIR = new Map();
    const CALLS_BY_NAME_PAIR = new Map();
    const SYNTHETIC_CALLS_CACHE = new Map();
    const THREAD_TIMELINE_CACHE = new Map();
    const SORTED_THREAD_TIMELINE_CACHE = new Map();
    const LINKED_CALLS_BY_THREAD_CACHE = new Map();
    const THREAD_SEARCH_TEXT_CACHE = new Map();
    const THREAD_SIDEBAR_SUMMARY_CACHE = new Map();
    const CALL_LABEL_CACHE = new Map();
    const CALL_SEARCH_TEXT_CACHE = new Map();
    const CALL_PRIMARY_TIME_VALUE_CACHE = new Map();
    const CALL_KEY_CACHE = new WeakMap();
    const LINKED_CALL_BY_MESSAGE_CACHE = new WeakMap();
    const MESSAGE_HIDDEN_META_CACHE = new WeakMap();
    const MESSAGE_SEARCH_TEXT_CACHE = new WeakMap();
    for (const call of DATA.calls || []) {
      rememberPerson(call.originator_id, call.originator_display_name);
      rememberPerson(call.target_id, call.target_display_name);
      for (const session of call.participant_sessions || []) {
        rememberPerson(session.id, session.display_name);
      }
      const participantIds = (call.participant_ids || []).map(normalizeGuid).filter(Boolean);
      const participantNames = (call.participant_display_names || []).map(value => stripFcs(value)).filter(Boolean);
      if (participantIds.length === participantNames.length) {
        for (let index = 0; index < participantIds.length; index += 1) {
          rememberPerson(participantIds[index], participantNames[index]);
        }
      }
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

    function cleanCallParticipantName(value) {
      let cleaned = stripFcs(value || "").trim();
      if (!cleaned) return "";
      cleaned = cleaned.replace(/\\s+Conference Virtual Assistant$/i, "").trim();
      cleaned = cleaned.replace(
        /^(.*?)(?:\\s+\\d+\\s+(?:28:[0-9a-fA-F-]{36}|8:orgid:[0-9a-fA-F-]{36}|4:\\+?[0-9(). -]{7,})(?:\\s+\\+?1?[0-9(). -]{7,})?.*)$/i,
        "$1"
      ).trim();
      return stripFcs(cleaned);
    }

    function normalizePhoneNumber(value) {
      let text = String(value || "").trim();
      if (!text) return "";
      text = text.replace(/^tel:/i, "");
      if (/^4:/i.test(text)) {
        text = text.replace(/^4:/i, "");
      }
      if (/[A-Za-z]/.test(text)) return "";
      if (!/^[+()0-9 .-]{7,}$/.test(text)) return "";
      text = text.replace(/[() .-]+/g, "");
      const hasPlus = text.startsWith("+");
      const digits = text.replace(/^\\+/, "");
      if (!digits || !/^\\d+$/.test(digits) || digits.length < 7) return "";
      return hasPlus ? `+${digits}` : digits;
    }

    function formatPhoneNumber(value) {
      const normalized = normalizePhoneNumber(value);
      if (!normalized) return "";
      const digits = normalized.replace(/^\\+/, "");
      if (digits.length === 10) {
        return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
      }
      if (digits.length === 11 && digits.startsWith("1")) {
        const local = digits.slice(1);
        return `(${local.slice(0, 3)}) ${local.slice(3, 6)}-${local.slice(6)}`;
      }
      return normalized.startsWith("+") ? normalized : digits;
    }

    function isGenericPhoneLabel(value) {
      return /^(wireless caller|state[_ ]of[_ ]alaska|unknown caller|anonymous|private caller|external caller)$/i.test(stripFcs(value || ""));
    }

    function callPhoneNumber(call, side) {
      return formatPhoneNumber(call[`${side}_phone_number`] || call[`${side}_endpoint`]);
    }

    function callSideLabel(call, side) {
      const rawName = cleanCallParticipantName(call[`${side}_display_name`] || "");
      const phone = callPhoneNumber(call, side);
      if (phone) {
        if (!rawName || isGenericPhoneLabel(rawName)) return phone;
        if (normalizePhoneNumber(rawName) === normalizePhoneNumber(phone)) return phone;
        return rawName;
      }
      if (normalizePhoneNumber(rawName)) return formatPhoneNumber(rawName);
      return rawName || phone || stripFcs(call[`${side}_id`] || call[`${side}_endpoint`] || "");
    }

    function callPanelSideLabel(call, side) {
      const label = callSideLabel(call, side);
      return isGenericPhoneLabel(label) ? "" : label;
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
      const text = String(value || "").trim();
      if (!text) return "";
      const match = text.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
      return match ? match[0].toLowerCase() : text.toLowerCase();
    }

    function normalizeName(value) {
      return stripFcs(value).toLowerCase();
    }

    function rememberPerson(guid, displayName) {
      const normalizedGuid = normalizeGuid(guid);
      const cleanedName = cleanCallParticipantName(displayName);
      if (normalizedGuid && cleanedName && !PERSON_NAME_BY_GUID.has(normalizedGuid)) {
        PERSON_NAME_BY_GUID.set(normalizedGuid, cleanedName);
      }
      if (normalizedGuid && cleanedName && !PERSON_GUID_BY_NAME.has(normalizeName(cleanedName))) {
        PERSON_GUID_BY_NAME.set(normalizeName(cleanedName), normalizedGuid);
      }
    }

    function personNameForGuid(guid) {
      return stripFcs(PERSON_NAME_BY_GUID.get(normalizeGuid(guid)) || "");
    }

    function participantBaseName(value) {
      const cleaned = cleanCallParticipantName(value || "").trim();
      if (!cleaned) return "";
      const bracketMatch = cleaned.match(/^(.*?)\\s*\\[([^\\]]+)\\]\\s*$/);
      if (bracketMatch && normalizePhoneNumber(bracketMatch[2])) {
        return cleanCallParticipantName(bracketMatch[1] || "");
      }
      return cleaned;
    }

    function personGuidForName(name) {
      return PERSON_GUID_BY_NAME.get(normalizeName(participantBaseName(name))) || "";
    }

    function looksLikeGuid(value) {
      return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(stripFcs(value || ""));
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

    function rememberDirectChatThread(thread) {
      if (!thread || thread.category !== "chat_space") return;
      if (!Array.isArray(thread.messages) || thread.messages.length === 0) return;
      const ids = dedupe((thread.participant_ids || []).map(normalizeGuid).filter(Boolean));
      const names = dedupe((thread.participants || []).map(value => stripFcs(value)).filter(Boolean));
      const hasCurrentId = Boolean(CURRENT_USER_ID && ids.includes(CURRENT_USER_ID));
      const hasCurrentName = Boolean(CURRENT_USER_NAME && names.some(name => normalizeName(name) === CURRENT_USER_NAME));
      if (!hasCurrentId && !hasCurrentName) return;

      if (ids.length === 2 && names.length === 2) {
        const otherId = ids.find(guid => guid !== CURRENT_USER_ID) || "";
        const otherName = names.find(name => normalizeName(name) !== CURRENT_USER_NAME) || "";
        rememberPerson(otherId, otherName);
      }

      for (const guid of ids) {
        if (guid && guid !== CURRENT_USER_ID && !DIRECT_CHAT_BY_GUID.has(guid)) {
          DIRECT_CHAT_BY_GUID.set(guid, thread);
        }
      }
      for (const name of names) {
        const normalizedName = normalizeName(name);
        if (normalizedName && normalizedName !== CURRENT_USER_NAME && !DIRECT_CHAT_BY_NAME.has(normalizedName)) {
          DIRECT_CHAT_BY_NAME.set(normalizedName, thread);
        }
      }
    }

    for (const thread of DATA.threads || []) {
      if (thread && thread.id) {
        THREAD_BY_ID.set(thread.id, thread);
      }
      for (const message of thread.messages || []) {
        rememberPerson(message.sender_id, message.sender_display_name);
      }
      rememberDirectChatThread(thread);
    }

    function participantGuidPairKey(values) {
      const items = dedupe((values || []).map(normalizeGuid).filter(Boolean)).sort();
      return items.length === 2 ? items.join("|") : "";
    }

    function participantNamePairKey(values) {
      const items = dedupe((values || []).map(normalizeName).filter(Boolean)).sort();
      return items.length === 2 ? items.join("|") : "";
    }

    function summarizeNames(values, maxVisible = 3) {
      const items = dedupe((values || []).map(value => cleanCallParticipantName(value)).filter(Boolean));
      if (!items.length) return "";
      if (items.length <= maxVisible) return items.join(" / ");
      return `${items.slice(0, maxVisible).join(" / ")} + ${items.length - maxVisible} more`;
    }

    function actualGuid(value) {
      const match = String(value || "").match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
      return match ? match[0].toLowerCase() : "";
    }

    function participantNameIdentity(name) {
      const cleaned = participantBaseName(name);
      if (!cleaned) return "";
      const phone = normalizePhoneNumber(cleaned);
      if (phone) return `phone:${phone}`;
      const normalized = normalizeName(cleaned);
      return normalized ? `name:${normalized}` : "";
    }

    function isPhoneEnrichedLabel(name) {
      const cleaned = stripFcs(name || "").trim();
      const bracketMatch = cleaned.match(/^(.*?)\\s*\\[([^\\]]+)\\]\\s*$/);
      return Boolean(bracketMatch && normalizePhoneNumber(bracketMatch[2]));
    }

    function chooseParticipantLabel(existingName, candidateName) {
      const existing = cleanCallParticipantName(existingName || "");
      const candidate = cleanCallParticipantName(candidateName || "");
      if (!existing) return candidate;
      if (!candidate) return existing;
      const existingEnriched = isPhoneEnrichedLabel(existing) || Boolean(normalizePhoneNumber(existing));
      const candidateEnriched = isPhoneEnrichedLabel(candidate) || Boolean(normalizePhoneNumber(candidate));
      if (candidateEnriched && !existingEnriched) return candidate;
      if (candidate.length > existing.length) return candidate;
      return existing;
    }

    function callParticipantDisplayNames(call) {
      return normalizeParticipantList(callParticipantEntries(call))
        .filter(entry => !entry.hidden)
        .map(entry => cleanCallParticipantName(entry.label || entry.name || ""))
        .filter(Boolean);
    }

    function callParticipantIds(call) {
      return dedupe([
        ...(call.participant_ids || []),
        call.originator_id,
        call.target_id,
      ].map(normalizeGuid).filter(Boolean));
    }

    function callParticipantNames(call) {
      return callParticipantDisplayNames(call).map(normalizeName).filter(Boolean);
    }

    function callParticipantEntries(call) {
      const combined = [];
      const mergedByIdentity = new Map();

      function mergeEntry(entry) {
        const name = stripFcs(entry && entry.name || "");
        const cleanName = cleanCallParticipantName(name);
        if (!cleanName || isGenericPhoneLabel(cleanName)) return;
        const guid = actualGuid(entry && entry.id) || personGuidForName(cleanName);
        const identityKey = guid || participantNameIdentity(cleanName);
        if (!identityKey) return;
        const existing = mergedByIdentity.get(identityKey);
        mergedByIdentity.set(identityKey, {
          id: guid || (existing && existing.id) || "",
          name: chooseParticipantLabel(existing && existing.name, cleanName),
        });
      }

      const originatorLabel = callPanelSideLabel(call, "originator");
      if (originatorLabel) {
        mergeEntry({
          id: call.originator_id || call.originator_endpoint || "",
          name: originatorLabel,
        });
      }
      const targetLabel = callPanelSideLabel(call, "target");
      if (targetLabel) {
        mergeEntry({
          id: call.target_id || call.target_endpoint || "",
          name: targetLabel,
        });
      }

      for (const session of call.participant_sessions || []) {
        mergeEntry({
          id: session.id || "",
          name: session.display_name || "",
        });
      }

      const baseEntries = buildParticipantEntries(
        call.participant_ids || [],
        call.participant_display_names || []
      );
      for (const entry of baseEntries) {
        mergeEntry(entry);
      }
      for (const entry of mergedByIdentity.values()) {
        combined.push(entry);
      }
      return combined;
    }

    function threadParticipantEntries(thread) {
      const entries = [];
      const resolvedIds = (thread.participant_ids || []).map(normalizeGuid).filter(Boolean);
      for (const guid of resolvedIds) {
        entries.push({ id: guid, name: personNameForGuid(guid) });
      }
      for (const name of thread.participants || []) {
        entries.push({ id: "", name: cleanCallParticipantName(name) });
      }
      return entries;
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

    function isMeetingThreadId(value) {
      return /^e?19:meeting_/i.test(String(value || "").trim());
    }

    function hasRealMeetingWindow(call) {
      return [call.meeting_start_time, call.meeting_end_time]
        .some(value => value && !String(value).startsWith("0001-01-01"));
    }

    function isMeetingCall(call) {
      if (isMeetingThreadId(call.conversation_id) || isMeetingThreadId(call.group_chat_thread_id)) {
        return true;
      }
      if (hasRealMeetingWindow(call)) {
        return true;
      }
      return Boolean(String(call.meeting_series_kind || "").trim());
    }

    function normalizedCallType(call) {
      return String(call.call_type || "").trim().toLowerCase();
    }

    function normalizedCallQuality(call) {
      return String(call.quality || "").trim().toLowerCase();
    }

    function displayCallTypeLabel(call) {
      const value = normalizedCallType(call);
      const labels = {
        meeting: "Meeting",
        group: "Group",
        twoparty: "Two Party",
        multiparty: "Multi Party",
      };
      if (labels[value]) return labels[value];
      return value ? prettyCallEventName(value) : "";
    }

    function displayCallQualityLabel(call) {
      const value = normalizedCallQuality(call);
      const labels = {
        thread_event_inferred: "Inferred Call",
        calllog_enriched: "Call Log",
        structured: "Structured Call",
      };
      if (labels[value]) return labels[value];
      return value ? prettyCallEventName(value) : "";
    }

    function looksLikePhoneEndpoint(name) {
      const cleaned = stripFcs(name || "").trim();
      if (!cleaned) return false;
      if (normalizePhoneNumber(cleaned)) return true;
      if (isGenericPhoneLabel(cleaned)) return true;
      if (cleaned.includes(",")) return false;
      return cleaned.length >= 5 && cleaned === cleaned.toUpperCase();
    }

    function isPhoneCall(call) {
      const key = callKey(call);
      if (PHONE_CALL_BY_KEY.has(key)) return PHONE_CALL_BY_KEY.get(key);
      let value = false;
      if (!isMeetingCall(call)) {
        const explicitPhoneEndpoint = Boolean(callPhoneNumber(call, "originator") || callPhoneNumber(call, "target"));
        const participants = dedupe([
          ...(call.participant_display_names || []),
          call.originator_display_name,
          call.target_display_name,
        ].map(value => stripFcs(value)).filter(Boolean));
        const hasPhoneEndpoint = explicitPhoneEndpoint || participants.some(looksLikePhoneEndpoint);
        const inCallLogBucket = String(call.conversation_id || "").toLowerCase() === "48:calllogs";
        if (explicitPhoneEndpoint) {
          value = Boolean(inCallLogBucket);
        } else {
          const hasLinkedConversation = preferredThreadsForCall(call).length > 0;
          value = Boolean(hasPhoneEndpoint && inCallLogBucket && !hasLinkedConversation);
        }
      }
      PHONE_CALL_BY_KEY.set(key, value);
      return value;
    }

    function isCurrentUserParticipantEntry(entry) {
      const guid = normalizeGuid(entry && entry.id);
      const name = normalizeName(entry && (entry.name || entry.label || ""));
      return Boolean((guid && CURRENT_USER_ID && guid === CURRENT_USER_ID) || (name && CURRENT_USER_NAME && name === CURRENT_USER_NAME));
    }

    function callExternalParticipantCount(call) {
      const key = callKey(call);
      if (CALL_EXTERNAL_PARTICIPANT_COUNT_BY_KEY.has(key)) return CALL_EXTERNAL_PARTICIPANT_COUNT_BY_KEY.get(key);
      const entries = normalizeParticipantList(callParticipantEntries(call))
        .filter(entry => !entry.hidden)
        .filter(entry => !isCurrentUserParticipantEntry(entry));
      const count = entries.length;
      CALL_EXTERNAL_PARTICIPANT_COUNT_BY_KEY.set(key, count);
      return count;
    }

    function hasLinkedGroupThread(call) {
      return preferredThreadsForCall(call).some(thread => thread.category === "thread" && (thread.participants || []).length > 2);
    }

    function isGroupCall(call) {
      if (isMeetingCall(call) || isPhoneCall(call)) return false;
      if (call.group_chat_thread_id) return true;
      if (callExternalParticipantCount(call) > 1) return true;
      return hasLinkedGroupThread(call);
    }

    function callGroupKey(call) {
      const key = callKey(call);
      if (CALL_GROUP_KEY_BY_KEY.has(key)) return CALL_GROUP_KEY_BY_KEY.get(key);
      let value = "call";
      if (isPhoneCall(call)) value = "phone";
      else if (isMeetingCall(call)) value = "meeting";
      else if (isGroupCall(call)) value = "group";
      CALL_GROUP_KEY_BY_KEY.set(key, value);
      return value;
    }

    function displayCallGroupLabel(call) {
      const labels = {
        phone: "Phone",
        meeting: "Meeting",
        group: "Group",
        call: "Call",
      };
      return labels[callGroupKey(call)] || "Call";
    }

    function meetingParticipation(call) {
      const value = String(call.user_participation || "").toLowerCase();
      return value === "joined" || value === "missed" ? value : "";
    }

    function meetingCallTitle(call) {
      const subject = stripFcs(call.meeting_subject || "").trim() || "Meeting";
      const when = fmt(call.start_time || call.meeting_start_time || call.connect_time);
      return when && when !== "[no_time]" ? `${subject} [${when}]` : subject;
    }

    function displayCallDirection(call) {
      if (isMeetingCall(call)) return "meeting";
      return String(call.direction || "").toLowerCase();
    }

    function displayCallDirectionLabel(call) {
      const direction = displayCallDirection(call);
      const labels = {
        incoming: "Inbound",
        outgoing: "Outbound",
        meeting: "Meeting",
      };
      return labels[direction] || (direction ? `${direction[0].toUpperCase()}${direction.slice(1)}` : "");
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
      if (isMeetingCall(call)) return meetingParticipation(call) || "meeting";
      if (shouldMarkMissed(call)) return "missed";
      const raw = String(call.call_state || "").toLowerCase();
      if (raw === "rejected" || raw === "declined" || raw === "cancelled") return "declined";
      return raw;
    }

    function displayCallStateLabel(call) {
      const state = displayCallState(call);
      const labels = {
        started: "Started",
        ended: "Ended",
        joined: "Joined",
        meeting: "Meeting",
        missed: "Missed",
        accepted: "Accepted",
        declined: "Declined",
        rejected: "Rejected",
        cancelled: "Cancelled",
      };
      if (labels[state]) return labels[state];
      return state ? prettyCallEventName(state) : "";
    }

    function displayCallStatus(call) {
      if (isMeetingCall(call)) {
        const participation = displayCallStateLabel(call);
        return participation && participation !== "Meeting" ? `Meeting | ${participation}` : "Meeting";
      }
      const direction = displayCallDirectionLabel(call);
      const state = displayCallStateLabel(call);
      if (direction && state) return `${direction} | ${state}`;
      return direction || state || "";
    }

    function includeRawCallTypeChip(call) {
      const rawType = normalizedCallType(call);
      const group = callGroupKey(call);
      if (!rawType) return false;
      if (group === "meeting" && rawType === "meeting") return false;
      if (group === "call" && (rawType === "twoparty" || rawType === "multiparty")) return false;
      if (group === "phone" && (rawType === "twoparty" || rawType === "multiparty")) return false;
      if (group === "group" && (rawType === "multiparty" || rawType === "group")) return false;
      return true;
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
      if (category === "team_chat" || category === "meeting") return "Meeting Chat";
      if (category === "thread") return "Group Chat";
      return category || "";
    }

    function isChatBasedThread(thread) {
      return ["chat_space", "thread", "team_chat", "meeting"].includes(thread.category);
    }

    function matchesThreadCategory(thread, category) {
      if (!category) return true;
      if (category === "team_chat") {
        return ["team_chat", "meeting"].includes(thread.category);
      }
      return thread.category === category;
    }

    function threadSearchText(thread) {
      const cacheKey = String(thread && thread.id || "");
      if (!hasSidebarMessageDateRange() && cacheKey && THREAD_SEARCH_TEXT_CACHE.has(cacheKey)) {
        return THREAD_SEARCH_TEXT_CACHE.get(cacheKey);
      }
      const messageTerms = visibleSidebarThreadMessages(thread).slice(0, 100).flatMap(message => [
        message.content_text || "",
        message.content_html || "",
        displayMessageType(message),
        ...messageAttachments(message).map(attachment => attachment.name || ""),
      ]);
      const parts = [
        thread.label,
        thread.id,
        thread.category,
        ...(thread.participants || []),
        ...messageTerms,
        ...syntheticCallsForThread(thread).slice(0, 40).map(call => call.summary_text || call.call_state || call.call_type || ""),
      ];
      const value = parts.filter(Boolean).join(" ").toLowerCase();
      if (!hasSidebarMessageDateRange() && cacheKey) {
        THREAD_SEARCH_TEXT_CACHE.set(cacheKey, value);
      }
      return value;
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

    function threadNameSearchPriority(thread) {
      const query = String(state.messageSearch || "").trim().toLowerCase();
      const terms = searchTerms();
      if (!query) return 2;
      const candidates = dedupe([
        stripFcs(thread.label || ""),
        ...(thread.participants || []).map(value => stripFcs(value)),
      ].filter(Boolean));
      let best = 2;
      for (const candidate of candidates) {
        const lowered = candidate.toLowerCase();
        const words = lowered.split(/[^a-z0-9]+/).filter(Boolean);
        if (!lowered) continue;
        if (lowered === query) return 0;
        if (words.some(word => word.startsWith(query))) {
          best = Math.min(best, 0);
          continue;
        }
        if (terms.length > 1 && terms.every(term => words.some(word => word.startsWith(term)))) {
          best = Math.min(best, 0);
          continue;
        }
        if (lowered.includes(query)) {
          best = Math.min(best, 1);
          continue;
        }
        if (terms.length > 1 && terms.every(term => lowered.includes(term))) {
          best = Math.min(best, 1);
        }
      }
      return best;
    }

    function threadSearchCategoryPriority(thread) {
      if (thread.category === "chat_space") return 0;
      if (thread.category === "thread") return 1;
      if (thread.category === "team_chat" || thread.category === "meeting") return 2;
      return 3;
    }

    function hasSidebarMessageDateRange() {
      return Boolean(state.messageDateFrom || state.messageDateTo);
    }

    function normalizedSidebarMessageDateRange() {
      const start = dateBoundaryValue(state.messageDateFrom, false);
      const end = dateBoundaryValue(state.messageDateTo, true);
      if (start === null && end === null) return { start: null, end: null };
      if (start !== null && end !== null && start > end) {
        return { start: end, end: start };
      }
      return { start, end };
    }

    function messagePassesSidebarDateRange(message) {
      const { start, end } = normalizedSidebarMessageDateRange();
      if (start === null && end === null) return true;
      const current = timeValue(message.timestamp);
      if (!Number.isFinite(current)) return false;
      if (start !== null && current < start) return false;
      if (end !== null && current > end) return false;
      return true;
    }

    function visibleSidebarThreadMessages(thread) {
      if (!hasSidebarMessageDateRange()) return thread.messages || [];
      return (thread.messages || []).filter(messagePassesSidebarDateRange);
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

    function threadSidebarSummary(thread) {
      const key = String(thread && thread.id || "");
      if (!key || hasSidebarMessageDateRange()) return null;
      if (THREAD_SIDEBAR_SUMMARY_CACHE.has(key)) {
        return THREAD_SIDEBAR_SUMMARY_CACHE.get(key);
      }

      let latestMessage = null;
      let latestTime = Number.NEGATIVE_INFINITY;
      for (const message of thread.messages || []) {
        const current = timeValue(message.timestamp);
        if (current >= latestTime) {
          latestTime = current;
          latestMessage = message;
        }
      }

      const summary = {
        latestMessage,
        latestTimestamp: latestTime,
        latestTimestampRaw: latestMessage ? (latestMessage.timestamp || null) : null,
      };
      THREAD_SIDEBAR_SUMMARY_CACHE.set(key, summary);
      return summary;
    }

    function threadPreviewForMessage(message) {
      if (!message) return "No recoverable message preview";
      const attachmentPreview = messageAttachmentPreview(message);
      if (attachmentPreview) {
        if (message.content_text) {
          return truncate(`${stripFcs(message.content_text)} [${attachmentPreview}]`);
        }
        return truncate(attachmentPreview);
      }
      if (message.message_type === "Event/Call") {
        return truncate(`Call event: ${prettyCallEventName(callEventName(message.content_text || ""))}`);
      }
      if (message.message_type === "ThreadActivity/AddMember") {
        const meta = messageHiddenMeta(message);
        return truncate(meta && meta.previewLabel ? meta.previewLabel : "Member added");
      }
      if (message.message_type === "ThreadActivity/MemberJoined") {
        const meta = messageHiddenMeta(message);
        return truncate(meta && meta.previewLabel ? meta.previewLabel : "Member joined");
      }
      if (message.message_type === "ThreadActivity/DeleteMember") {
        const meta = messageHiddenMeta(message);
        return truncate(meta && meta.previewLabel ? meta.previewLabel : "Member removed");
      }
      return truncate(stripFcs(message.content_text || message.message_type || ""));
    }

    function latestThreadMessage(thread) {
      const summary = threadSidebarSummary(thread);
      if (summary) return summary.latestMessage;
      let bestMessage = null;
      let bestTime = Number.NEGATIVE_INFINITY;
      for (const message of visibleSidebarThreadMessages(thread)) {
        const current = timeValue(message.timestamp);
        if (current >= bestTime) {
          bestTime = current;
          bestMessage = message;
        }
      }
      return bestMessage;
    }

    function latestThreadPreview(thread) {
      const summary = threadSidebarSummary(thread);
      if (summary) return threadPreviewForMessage(summary.latestMessage);
      return threadPreviewForMessage(latestThreadMessage(thread));
    }

    function callPrimaryTimestamp(call) {
      return call.start_time || call.connect_time || call.end_time || "";
    }

    function threadLastTimestamp(thread) {
      const summary = threadSidebarSummary(thread);
      if (summary) return summary.latestTimestamp;
      let latest = Number.NEGATIVE_INFINITY;
      for (const message of visibleSidebarThreadMessages(thread)) {
        latest = Math.max(latest, timeValue(message.timestamp));
      }
      return latest;
    }

    function threadLastTimestampRaw(thread) {
      const summary = threadSidebarSummary(thread);
      if (summary) return summary.latestTimestampRaw;
      let best = null;
      let latest = Number.NEGATIVE_INFINITY;
      for (const message of visibleSidebarThreadMessages(thread)) {
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
          if (thread.message_count <= 0 && !["team_chat", "meeting"].includes(thread.category)) return false;
          if (!matchesThreadCategory(thread, state.category)) {
            return false;
          }
          if (visibleSidebarThreadMessages(thread).length === 0) return false;
          if (state.messageSearch && !textMatchesSearch(threadSearchText(thread))) return false;
          return true;
        })
        .sort((left, right) => {
          if (state.messageSearch) {
            const leftPriority = threadNameSearchPriority(left);
            const rightPriority = threadNameSearchPriority(right);
            const priorityDiff = leftPriority - rightPriority;
            if (priorityDiff !== 0) return priorityDiff;
            if (leftPriority < 2 && rightPriority < 2) {
              const categoryDiff = threadSearchCategoryPriority(left) - threadSearchCategoryPriority(right);
              if (categoryDiff !== 0) return categoryDiff;
            }
          }
          const timeDiff = threadLastTimestamp(right) - threadLastTimestamp(left);
          if (timeDiff !== 0) return timeDiff;
          return (left.label || left.id || "").localeCompare(right.label || right.id || "");
        });
    }

    function callKey(call) {
      if (!call || typeof call !== "object") return "||||";
      if (CALL_KEY_CACHE.has(call)) return CALL_KEY_CACHE.get(call);
      const value = [
        call.call_id || "",
        call.start_time || "",
        call.shared_correlation_id || "",
        call.originator_id || "",
        call.target_id || "",
      ].join("|");
      CALL_KEY_CACHE.set(call, value);
      return value;
    }

    function callLabel(call) {
      const key = callKey(call);
      if (CALL_LABEL_CACHE.has(key)) return CALL_LABEL_CACHE.get(key);
      let value = "[unknown call]";
      if (isMeetingCall(call)) {
        value = meetingCallTitle(call);
      } else {
        const participants = callParticipantDisplayNames(call);
        if (participants.length) {
          value = summarizeNames(participants);
        } else {
          const origin = callSideLabel(call, "originator");
          const target = callSideLabel(call, "target");
          value = origin && target ? `${origin} -> ${target}` : (origin || target || call.summary_text || call.call_id || "[unknown call]");
        }
      }
      CALL_LABEL_CACHE.set(key, value);
      return value;
    }

    function extractGuids(text) {
      return Array.from(String(text || "").matchAll(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g))
        .map(match => match[0].toLowerCase());
    }

    function eventCallGuid(text) {
      const match = String(text || "").match(/\\b([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\\b(?:\\s+False)?\\s*$/i);
      return match ? match[1].toLowerCase() : "";
    }

    function findLinkedCall(message) {
      if (LINKED_CALL_BY_MESSAGE_CACHE.has(message)) {
        return LINKED_CALL_BY_MESSAGE_CACHE.get(message);
      }
      let linked = null;
      if (message.synthetic_call_key && CALLS_BY_KEY.has(message.synthetic_call_key)) {
        linked = CALLS_BY_KEY.get(message.synthetic_call_key);
      } else {
        const eventGuid = eventCallGuid(message.content_text || "");
        if (eventGuid && CALLS_BY_ID.has(eventGuid)) {
          linked = CALLS_BY_ID.get(eventGuid);
        } else if (eventGuid && CALLS_BY_SHARED.has(eventGuid)) {
          linked = CALLS_BY_SHARED.get(eventGuid);
        } else {
          const ids = extractGuids(message.content_text || "");
          for (const id of ids) {
            if (CALLS_BY_ID.has(id)) {
              linked = CALLS_BY_ID.get(id);
              break;
            }
          }
          if (!linked) {
            for (const id of ids) {
              if (CALLS_BY_SHARED.has(id)) {
                linked = CALLS_BY_SHARED.get(id);
                break;
              }
            }
          }
        }
      }
      LINKED_CALL_BY_MESSAGE_CACHE.set(message, linked);
      return linked;
    }

    function callEventName(text) {
      const match = String(text || "").match(/\\b(callStarted|callEnded|callMissed|callAccepted|callRejected|callCancelled)\\b/i);
      return match ? match[1] : "callEvent";
    }

    function prettyCallEventName(value) {
      return String(value || "")
        .replace(/[_-]+/g, " ")
        .replace(/^call/, "Call ")
        .replace(/([a-z])([A-Z])/g, "$1 $2")
        .trim();
    }

    function collectEventParticipants(message) {
      const seen = [];
      for (const guid of extractGuids(message.content_text || "")) {
        const name = personNameForGuid(guid);
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

    function linkedCallsForThread(thread) {
      const cacheKey = String(thread && thread.id || "");
      if (cacheKey && LINKED_CALLS_BY_THREAD_CACHE.has(cacheKey)) {
        return LINKED_CALLS_BY_THREAD_CACHE.get(cacheKey);
      }
      const linked = new Map();
      for (const message of thread.messages || []) {
        const call = findLinkedCall(message);
        if (call) linked.set(callKey(call), call);
      }
      for (const call of syntheticCallsForThread(thread)) {
        linked.set(callKey(call), call);
      }
      const value = [...linked.values()].sort((left, right) => timeValue(callTimelineTimestamp(right)) - timeValue(callTimelineTimestamp(left)));
      if (cacheKey) {
        LINKED_CALLS_BY_THREAD_CACHE.set(cacheKey, value);
      }
      return value;
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
      const linkedKey = linkedCall ? callKey(linkedCall) : "";
      const visuals = callCardVisuals(message, linkedCall);
      const participantEntries = callEventParticipantEntries(message, linkedCall);

      const cardHtml = `
        <div class="call-event" style="--call-bg:${escapeHtml(visuals.bg)};--call-outline:${escapeHtml(visuals.outline)};--call-block-bg:${escapeHtml(visuals.block)};--call-title:${escapeHtml(visuals.title)};">
          <div class="call-event-header">
            <div class="call-event-title">${highlightSearchHtml(eventName)}</div>
            <div class="call-event-actions">
              ${linkedCall ? `<button class="call-link open-call-record" type="button" data-call-key="${escapeHtml(linkedKey)}">Open In Calls</button>` : ``}
            </div>
          </div>
          <div class="call-event-grid">
            <div class="call-event-block wide"><div class="k">Participants</div>${renderExpandableChipGroup(participantEntries, { emptyLabel: stripFcs(callLabel(linkedCall || {})) || "No participants were resolved." })}</div>
            <div class="call-event-block"><div class="k">Direction</div><div>${highlightSearchHtml(linkedCall ? displayCallDirectionLabel(linkedCall) : "")}</div></div>
            <div class="call-event-block"><div class="k">State</div><div>${highlightSearchHtml(linkedCall ? displayCallStateLabel(linkedCall) : eventName)}</div></div>
            <div class="call-event-block"><div class="k">Start Time</div><div>${escapeHtml(fmt(linkedCall ? linkedCall.start_time : message.timestamp))}</div></div>
            <div class="call-event-block"><div class="k">End Time</div><div>${escapeHtml(fmt(linkedCall ? linkedCall.end_time : null))}</div></div>
          </div>
          ${message.synthetic_call
            ? `<div class="subtle">Merged from the Calls dataset for this conversation.</div>`
            : ``}
        </div>
      `;
      return cardHtml;
    }

    function messageLinkedCallKey(message) {
      if (message.synthetic_call_key) return message.synthetic_call_key;
      const linked = findLinkedCall(message);
      return linked ? callKey(linked) : "";
    }

    function threadTimelineItems(thread) {
      const cacheKey = String(thread && thread.id || "");
      if (cacheKey && THREAD_TIMELINE_CACHE.has(cacheKey)) {
        return THREAD_TIMELINE_CACHE.get(cacheKey);
      }
      const value = [...(thread.messages || []), ...syntheticCallMessages(thread)];
      if (cacheKey) {
        THREAD_TIMELINE_CACHE.set(cacheKey, value);
      }
      return value;
    }

    function sortedThreadTimelineItems(thread) {
      const cacheKey = String(thread && thread.id || "");
      if (cacheKey && SORTED_THREAD_TIMELINE_CACHE.has(cacheKey)) {
        return SORTED_THREAD_TIMELINE_CACHE.get(cacheKey);
      }
      const value = threadTimelineItems(thread)
        .slice()
        .sort((left, right) => {
          const timeDiff = timeValue(right.timestamp) - timeValue(left.timestamp);
          if (timeDiff !== 0) return timeDiff;
          return (right.id || "").localeCompare(left.id || "");
        });
      if (cacheKey) {
        SORTED_THREAD_TIMELINE_CACHE.set(cacheKey, value);
      }
      return value;
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
      const cleaned = personNameForGuid(guid);
      return cleaned || stripFcs(guid || "Unknown");
    }

    function systemEventActorLabel() {
      return "System";
    }

    function parseJsonObject(text) {
      try {
        const parsed = JSON.parse(String(text || ""));
        return parsed && typeof parsed === "object" ? parsed : null;
      } catch {
        return null;
      }
    }

    function resolveMemberEntries(entries) {
      const memberIds = [];
      const memberNames = [];
      for (const entry of Array.isArray(entries) ? entries : []) {
        const memberId = normalizeGuid((entry && (entry.id || entry.mri || entry.userId)) || "");
        const friendlyName = stripFcs((entry && (entry.friendlyname || entry.displayName || entry.name)) || "");
        const resolvedName = friendlyName || (memberId ? personNameForGuid(memberId) : "");
        if (memberId && !memberIds.includes(memberId)) memberIds.push(memberId);
        if (resolvedName && !memberNames.includes(resolvedName)) memberNames.push(resolvedName);
      }
      return { memberIds, memberNames };
    }

    function participantEntryKey(entry) {
      const guid = normalizeGuid(entry && entry.id);
      const name = normalizeName(entry && entry.name);
      return guid || name || "";
    }

    function buildParticipantEntries(ids = [], names = []) {
      const entries = [];
      if (ids.length && ids.length === names.length) {
        for (let index = 0; index < ids.length; index += 1) {
          const guid = normalizeGuid(ids[index]);
          const name = cleanCallParticipantName(names[index] || (guid ? resolveGuidLabel(guid) : ""));
          entries.push({ id: guid, name: name || (guid ? resolveGuidLabel(guid) : "") });
        }
      } else {
        for (const guid of ids || []) {
          const normalizedGuid = normalizeGuid(guid);
          if (normalizedGuid) {
            entries.push({ id: normalizedGuid, name: resolveGuidLabel(normalizedGuid) });
          }
        }
        for (const name of names || []) {
          const cleanedName = cleanCallParticipantName(name);
          if (cleanedName) entries.push({ id: "", name: cleanedName });
        }
      }

      const unique = [];
      const seen = new Set();
      for (const entry of entries) {
        const key = participantEntryKey(entry);
        if (!key || seen.has(key)) continue;
        seen.add(key);
        unique.push({
          id: normalizeGuid(entry.id || ""),
          name: cleanCallParticipantName(entry.name || (entry.id ? personNameForGuid(entry.id) : "")),
        });
      }
      return unique;
    }

    function hiddenEntryMeta(entries, noun = "members") {
      const normalized = normalizeParticipantList(entries);
      const visibleEntries = sortParticipantEntries(normalized.filter(entry => !entry.hidden));
      const hiddenEntries = sortParticipantEntries(normalized.filter(entry => entry.hidden));
      const count = hiddenEntries.length;
      const singular = noun.replace(/s$/, "");
      return {
        count,
        label: count ? `${count} hidden ${count === 1 ? singular : noun}` : "",
        visibleEntries,
        hiddenEntries,
      };
    }

    function previewLabelForEntries(prefix, entries, noun = "members") {
      const meta = hiddenEntryMeta(entries, noun);
      const visibleLabels = meta.visibleEntries.map(entry => stripFcs(entry.label || entry.name || ""));
      if (visibleLabels.length === 1 && meta.count === 0) return `${prefix}: ${visibleLabels[0]}`;
      if (visibleLabels.length > 1 && meta.count === 0) return `${prefix}: ${visibleLabels.length} ${noun}`;
      if (!visibleLabels.length && meta.count === 1) return `${prefix}: Hidden`;
      if (!visibleLabels.length && meta.count > 1) return `${prefix}: ${meta.label}`;
      return `${prefix}: ${visibleLabels.length + meta.count} ${noun}`;
    }

    function callEventParticipantEntries(message, linkedCall = findLinkedCall(message)) {
      const raw = message.content_text || "";
      const participants = collectEventParticipants(message);
      return linkedCall
        ? callParticipantEntries(linkedCall)
        : buildParticipantEntries(extractGuids(raw), participants.map(name => stripFcs(name)));
    }

    function messageHiddenMeta(message) {
      if (MESSAGE_HIDDEN_META_CACHE.has(message)) {
        return MESSAGE_HIDDEN_META_CACHE.get(message);
      }
      let value = null;
      if (message.synthetic_call || message.message_type === "Event/Call") {
        const entries = callEventParticipantEntries(message);
        const meta = hiddenEntryMeta(entries, "participants");
        value = meta.count ? {
          count: meta.count,
          label: meta.label,
          noun: "participants",
          previewLabel: previewLabelForEntries("Call event", entries, "participants"),
        } : null;
      } else if (message.message_type === "ThreadActivity/AddMember") {
        const parsed = parseAddMemberEvent(message);
        const entries = buildParticipantEntries(parsed.addedIds, parsed.addedNames);
        const meta = hiddenEntryMeta(entries, "members");
        value = {
          count: meta.count,
          label: meta.label,
          noun: "members",
          previewLabel: previewLabelForEntries("Member added", entries, "members"),
        };
      } else if (message.message_type === "ThreadActivity/MemberJoined") {
        const parsed = parseMemberJoinedEvent(message);
        const entries = buildParticipantEntries(parsed.joinedIds, parsed.joinedNames);
        const meta = hiddenEntryMeta(entries, "members");
        value = {
          count: meta.count,
          label: meta.label,
          noun: "members",
          previewLabel: previewLabelForEntries("Member joined", entries, "members"),
        };
      } else if (message.message_type === "ThreadActivity/DeleteMember") {
        const parsed = parseDeleteMemberEvent(message);
        const entries = buildParticipantEntries(parsed.removedIds, parsed.removedNames);
        const meta = hiddenEntryMeta(entries, "members");
        value = {
          count: meta.count,
          label: meta.label,
          noun: "members",
          previewLabel: previewLabelForEntries("Member removed", entries, "members"),
        };
      }
      MESSAGE_HIDDEN_META_CACHE.set(message, value);
      return value;
    }

    function displayMessageType(message) {
      if (message.message_type === "ThreadActivity/AddMember") return "Member Added";
      if (message.message_type === "ThreadActivity/MemberJoined") return "Member Joined";
      if (message.message_type === "ThreadActivity/DeleteMember") return "Member Removed";
      if (messageHasAttachments(message)) {
        if (message.content_text) return "Message + Attachment";
        return messageAttachmentPreview(message) || "Attachment";
      }
      return message.message_type || "";
    }

    function directChatThreadForPerson(entry) {
      const guid = normalizeGuid(entry && entry.id);
      const name = normalizeName(entry && entry.name);
      if ((guid && guid === CURRENT_USER_ID) || (name && name === CURRENT_USER_NAME)) return null;
      if (guid && DIRECT_CHAT_BY_GUID.has(guid)) {
        const thread = DIRECT_CHAT_BY_GUID.get(guid);
        if (!state.threadId || thread.id !== state.threadId) return thread;
      }
      if (name && DIRECT_CHAT_BY_NAME.has(name)) {
        const thread = DIRECT_CHAT_BY_NAME.get(name);
        if (!state.threadId || thread.id !== state.threadId) return thread;
      }
      return null;
    }

    function sortParticipantEntries(entries) {
      return entries.slice().sort((left, right) => {
        const leftLabel = stripFcs(left.label || left.name || left.id || "").toLowerCase();
        const rightLabel = stripFcs(right.label || right.name || right.id || "").toLowerCase();
        return leftLabel.localeCompare(rightLabel);
      });
    }

    function normalizeParticipantList(values) {
      const items = [];
      const seenIds = new Set();
      const seenNames = new Set();
      for (const value of values || []) {
        const entry = value && typeof value === "object" && !Array.isArray(value)
          ? {
              id: normalizeGuid(value.id || value.guid || value.userId || ""),
              name: cleanCallParticipantName(value.name || value.display_name || value.displayName || ""),
            }
          : {
              id: "",
              name: cleanCallParticipantName(value),
            };
      const guid = normalizeGuid(entry.id);
      const resolvedName = cleanCallParticipantName(entry.name || (guid ? personNameForGuid(guid) : ""));
        const visibleLabel = resolvedName && !looksLikeGuid(resolvedName) ? resolvedName : "";
        const hiddenLabel = stripFcs(guid || entry.name || "");
        const entryIdKey = guid || "";
        const visibleKey = normalizeName(visibleLabel);
        const hiddenKey = normalizeName(hiddenLabel);
        if (visibleLabel) {
          if (entryIdKey && seenIds.has(entryIdKey)) continue;
          if (visibleKey && seenNames.has(visibleKey)) continue;
          if (entryIdKey) seenIds.add(entryIdKey);
          if (visibleKey) seenNames.add(visibleKey);
          items.push({
            id: guid,
            name: visibleLabel,
            label: visibleLabel,
            hidden: false,
          });
          continue;
        }
        if (!hiddenLabel) continue;
        if (entryIdKey && seenIds.has(entryIdKey)) continue;
        if (!entryIdKey && hiddenKey && seenNames.has(hiddenKey)) continue;
        if (entryIdKey) seenIds.add(entryIdKey);
        if (hiddenKey) seenNames.add(hiddenKey);
        items.push({
          id: guid,
          name: "",
          label: hiddenLabel,
          hidden: true,
        });
      }
      return items;
    }

    function renderHiddenEntries(entries, summaryLabel = "Show Hidden Data") {
      const hiddenEntries = sortParticipantEntries(entries.filter(entry => entry.hidden));
      if (!hiddenEntries.length) return "";
      return `
        <details class="raw-event hidden-data" ${state.expandHiddenData ? "open" : ""}>
          <summary>${escapeHtml(`${summaryLabel} (${hiddenEntries.length})`)}</summary>
          <div class="chips">
            ${hiddenEntries.map(entry => `<span class="chip participant-chip participant-chip-static">${escapeHtml(entry.label)}</span>`).join("")}
          </div>
        </details>
      `;
    }

    function renderExpandableChipGroup(values, options = {}) {
      const entries = normalizeParticipantList(values);
      const visibleEntries = sortParticipantEntries(entries.filter(entry => !entry.hidden));
      const hiddenDetails = renderHiddenEntries(entries, options.hiddenSummaryLabel || "Show Hidden Data");
      if (!visibleEntries.length) {
        return `
          <div class="expandable-shell">
            <div class="subtle">${escapeHtml(options.emptyLabel || "No visible people listed.")}</div>
            ${hiddenDetails}
          </div>
        `;
      }
      return `
        <div class="expandable-shell">
          <div class="chips expandable-list collapsed" data-expandable-list>
            ${visibleEntries.map(entry => {
              const chatThread = options.enableChatLinks === false ? null : directChatThreadForPerson(entry);
              const label = stripFcs(entry.label || entry.name || "");
              if (chatThread) {
                return `<button type="button" class="chip participant-chip participant-chip-link open-participant-chat" data-thread-id="${escapeHtml(chatThread.id)}" data-label="${escapeHtml(label)}">${escapeHtml(label)}</button>`;
              }
              return `<span class="chip participant-chip participant-chip-static">${escapeHtml(label)}</span>`;
            }).join("")}
          </div>
          <button
            type="button"
            class="call-link toggle-expandable-list hidden"
            data-expand-label="${escapeHtml(options.expandLabel || "Expand")}"
            data-collapse-label="${escapeHtml(options.collapseLabel || "Collapse")}"
          >${escapeHtml(options.expandLabel || "Expand")}</button>
          ${hiddenDetails}
        </div>
      `;
    }

    function initExpandableLists(scope) {
      for (const shell of scope.querySelectorAll(".expandable-shell")) {
        const list = shell.querySelector("[data-expandable-list]");
        const button = shell.querySelector(".toggle-expandable-list");
        if (!list || !button) continue;

        list.classList.remove("expanded");
        list.classList.add("collapsed");
        const needsToggle = list.scrollHeight > list.clientHeight + 4;
        button.classList.toggle("hidden", !needsToggle);
        if (!needsToggle) {
          list.classList.remove("collapsed");
          continue;
        }

        const updateButtonText = () => {
          button.textContent = list.classList.contains("expanded")
            ? (button.dataset.collapseLabel || "Collapse")
            : (button.dataset.expandLabel || "Expand");
        };
        updateButtonText();
        button.addEventListener("click", () => {
          const expanded = list.classList.toggle("expanded");
          list.classList.toggle("collapsed", !expanded);
          updateButtonText();
        });
      }
    }

    function initParticipantChatLinks(scope) {
      for (const element of scope.querySelectorAll(".open-participant-chat")) {
        element.addEventListener("click", () => {
          state.view = "messages";
          state.threadId = element.dataset.threadId;
          state.category = "";
          state.messageViewFilter = "all";
          state.threadDateFrom = "";
          state.threadDateTo = "";
          state.focusTimestamp = null;
          state.focusCallKey = null;
          state.messageSearch = "";
          messageSearch.value = "";
          categoryFilter.value = "";
          renderView();
          scrollMainToTop();
        });
      }
    }

    function initAttachmentActions(scope) {
      for (const element of scope.querySelectorAll(".copy-attachment-link")) {
        element.addEventListener("click", () => copyText(element.dataset.url || "", "Attachment link copied"));
      }
    }

    function parseAddMemberEvent(message) {
      const unique = [];
      for (const guid of extractGuids(message.content_text || "")) {
        if (!unique.includes(guid)) unique.push(guid);
      }

      const senderId = String(message.sender_id || "").toLowerCase();
      const actorId = senderId || unique[0] || "";
      const actorName = stripFcs(message.sender_display_name || personNameForGuid(actorId) || "");
      const addedIds = unique.filter(guid => guid !== actorId);
      const addedNames = addedIds.map(personNameForGuid);

      return {
        actorId,
        actorName,
        addedIds,
        addedNames,
      };
    }

    function parseMemberJoinedEvent(message) {
      const payload = parseJsonObject(message.content_text || "");
      const actorId = normalizeGuid(payload && payload.initiator);
      const actorName = stripFcs(
        message.sender_display_name ||
        personNameForGuid(actorId) ||
        "System"
      );
      const members = resolveMemberEntries(payload && payload.members);
      return {
        actorId,
        actorName,
        joinedIds: members.memberIds,
        joinedNames: members.memberNames,
      };
    }

    function parseDeleteMemberEvent(message) {
      const actorId = normalizeGuid(message.sender_id || "");
      const actorName = stripFcs(
        message.sender_display_name ||
        personNameForGuid(actorId) ||
        ""
      );
      let removedIds = dedupe(extractGuids(message.content_text || "").map(normalizeGuid).filter(Boolean));
      if (actorId) {
        const filtered = removedIds.filter(id => id !== actorId);
        if (filtered.length) removedIds = filtered;
      }
      const removedNames = dedupe(removedIds.map(personNameForGuid).filter(Boolean));
      return {
        actorId,
        actorName,
        removedIds,
        removedNames,
      };
    }

    function isMembershipEvent(message) {
      return [
        "ThreadActivity/AddMember",
        "ThreadActivity/MemberJoined",
        "ThreadActivity/DeleteMember",
      ].includes(message.message_type);
    }

    function renderMembershipNamesBlock(label, values, emptyLabel = "No people listed.") {
      return `
        <div class="call-event-block wide">
          <div class="k">${escapeHtml(label)}</div>
          ${renderExpandableChipGroup(values, { emptyLabel })}
        </div>
      `;
    }

    function renderAddMemberEventBody(message, thread) {
      const parsed = parseAddMemberEvent(message);
      const addedEntries = buildParticipantEntries(parsed.addedIds, parsed.addedNames);
      const addedCount = parsed.addedIds.length;
      const cardHtml = `
        <div class="call-event" style="--call-bg:#e4e8ed;--call-outline:#66727f;--call-block-bg:#f4f6f8;--call-title:#3e4852;">
          <div class="call-event-header">
            <div class="call-event-title">${escapeHtml(addedCount > 1 ? "Members Added" : "Member Added")}</div>
          </div>
          <div class="call-event-grid">
            <div class="call-event-block"><div class="k">Actor</div><div>${escapeHtml(systemEventActorLabel())}</div></div>
            <div class="call-event-block"><div class="k">Added Count</div><div>${escapeHtml(String(addedCount || 0))}</div></div>
            <div class="call-event-block"><div class="k">Current Conversation</div><div>${escapeHtml(stripFcs(thread.label || thread.id || ""))}</div></div>
            ${renderMembershipNamesBlock("Added", addedEntries, "No added members were resolved.")}
          </div>
        </div>
      `;
      return cardHtml;
    }

    function renderMemberJoinedEventBody(message, thread) {
      const parsed = parseMemberJoinedEvent(message);
      const joinedEntries = buildParticipantEntries(parsed.joinedIds, parsed.joinedNames);
      const joinedCount = parsed.joinedIds.length;
      const cardHtml = `
        <div class="call-event" style="--call-bg:#e4f0e9;--call-outline:#5b7f69;--call-block-bg:#f5faf7;--call-title:#2f5e40;">
          <div class="call-event-header">
            <div class="call-event-title">${escapeHtml(joinedCount > 1 ? "Members Joined" : "Member Joined")}</div>
          </div>
          <div class="call-event-grid">
            <div class="call-event-block"><div class="k">Actor</div><div>${escapeHtml(systemEventActorLabel())}</div></div>
            <div class="call-event-block"><div class="k">Joined Count</div><div>${escapeHtml(String(joinedCount || 0))}</div></div>
            <div class="call-event-block"><div class="k">Current Conversation</div><div>${escapeHtml(stripFcs(thread.label || thread.id || ""))}</div></div>
            ${renderMembershipNamesBlock("Joined", joinedEntries, "No joined members were resolved.")}
          </div>
        </div>
      `;
      return cardHtml;
    }

    function renderDeleteMemberEventBody(message, thread) {
      const parsed = parseDeleteMemberEvent(message);
      const removedEntries = buildParticipantEntries(parsed.removedIds, parsed.removedNames);
      const removedCount = parsed.removedIds.length;
      const cardHtml = `
        <div class="call-event" style="--call-bg:#f6ece4;--call-outline:#9a6b52;--call-block-bg:#fff7f1;--call-title:#7c523b;">
          <div class="call-event-header">
            <div class="call-event-title">${escapeHtml(removedCount > 1 ? "Members Removed" : "Member Removed")}</div>
          </div>
          <div class="call-event-grid">
            <div class="call-event-block"><div class="k">Actor</div><div>${escapeHtml(systemEventActorLabel())}</div></div>
            <div class="call-event-block"><div class="k">Removed Count</div><div>${escapeHtml(String(removedCount || 0))}</div></div>
            <div class="call-event-block"><div class="k">Current Conversation</div><div>${escapeHtml(stripFcs(thread.label || thread.id || ""))}</div></div>
            ${renderMembershipNamesBlock("Removed", removedEntries, "No removed members were resolved.")}
          </div>
        </div>
      `;
      return cardHtml;
    }

    function renderMembershipEventBody(message, thread) {
      if (message.message_type === "ThreadActivity/AddMember") return renderAddMemberEventBody(message, thread);
      if (message.message_type === "ThreadActivity/MemberJoined") return renderMemberJoinedEventBody(message, thread);
      if (message.message_type === "ThreadActivity/DeleteMember") return renderDeleteMemberEventBody(message, thread);
      return `<div class="body">${escapeHtml(message.content_text || "")}</div>`;
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
      const key = callKey(call);
      if (PREFERRED_THREADS_BY_CALL_KEY.has(key)) return PREFERRED_THREADS_BY_CALL_KEY.get(key);
      const matches = linkedThreadsForCall(call);
      const directPairKeys = new Set(
        matches
          .filter(thread => thread.category === "chat_space" && (thread.participants || []).length === 2)
          .map(thread => participantNamePairKey(thread.participants || []))
          .filter(Boolean)
      );
      const filtered = matches.filter(thread => {
        const pairKey = participantNamePairKey(thread.participants || []);
        if (!pairKey || !directPairKeys.has(pairKey)) return true;
        return thread.category === "chat_space";
      });
      PREFERRED_THREADS_BY_CALL_KEY.set(key, filtered);
      return filtered;
    }

    function callJumpTargets(call) {
      const key = callKey(call);
      if (CALL_JUMP_TARGETS_BY_CALL_KEY.has(key)) return CALL_JUMP_TARGETS_BY_CALL_KEY.get(key);
      const targets = [];
      const targetKey = callKey(call);
      for (const thread of preferredThreadsForCall(call)) {
        const hasEventTarget = threadTimelineItems(thread).some(message => messageLinkedCallKey(message) === targetKey);
        const hasMessageTarget = hasNearbyCuratedMessages(thread, call);
        const focusTimestamp = call.start_time || call.end_time || "";
        if (hasMessageTarget || hasEventTarget) {
          targets.push({
            thread,
            focusTimestamp,
            focusCallKey: hasMessageTarget ? "" : targetKey,
            viewFilter: "all",
          });
        }
      }
      const seen = new Set();
      const filtered = targets.filter(target => {
        const key = `${target.thread.id}|${target.viewFilter}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
      CALL_JUMP_TARGETS_BY_CALL_KEY.set(key, filtered);
      return filtered;
    }

    function callJumpLabel(target) {
      const name = stripFcs(target.thread.label || target.thread.id || "Conversation");
      const when = fmt(target.focusTimestamp);
      return when && when !== "[no_time]" ? `${name} [${when}]` : name;
    }

    function attachmentNameFromUrl(value) {
      const text = String(value || "").trim();
      if (!text) return "";
      try {
        const parsed = new URL(text);
        const parts = parsed.pathname.split("/").filter(Boolean);
        return decodeURIComponent(parts[parts.length - 1] || "").trim();
      } catch {
        return "";
      }
    }

    function attachmentExtension(value) {
      const text = stripFcs(value || "").trim().toLowerCase();
      if (!text) return "";
      const target = text.includes(".") ? text : attachmentNameFromUrl(text).toLowerCase();
      const match = target.match(/\\.([a-z0-9]{2,6})(?:$|[?#])/i);
      return match ? match[1].toLowerCase() : "";
    }

    function normalizeAttachmentKind(value, name = "", url = "") {
      const raw = String(value || "").trim().toLowerCase();
      const ext = attachmentExtension(name) || attachmentExtension(url);
      if (["image", "amsimage", "inlineimage", "jpg", "jpeg", "png", "gif", "bmp", "webp", "heic", "heif", "tif", "tiff"].includes(raw) || ["jpg", "jpeg", "png", "gif", "bmp", "webp", "heic", "heif", "tif", "tiff"].includes(ext)) {
        return "image";
      }
      if (["video", "amsvideo", "mp4", "mov", "avi", "wmv", "m4v", "webm"].includes(raw) || ["mp4", "mov", "avi", "wmv", "m4v", "webm"].includes(ext)) {
        return "video";
      }
      return "file";
    }

    function attachmentKindLabel(kind) {
      const normalized = normalizeAttachmentKind(kind);
      if (normalized === "image") return "Photo";
      if (normalized === "video") return "Video";
      return "File";
    }

    function attachmentHost(url) {
      const value = String(url || "").trim();
      if (!value) return "";
      try {
        return new URL(value).hostname.toLowerCase();
      } catch (error) {
        return "";
      }
    }

    function attachmentSourceLabel(attachment) {
      const host = attachmentHost(attachment.url || attachment.preview_url || "");
      if (!host) return "Remote attachment";
      if (host.includes("api.ams.") || host.includes(".teams.microsoft.com")) return "Teams media";
      if (host.includes("sharepoint.com")) return "SharePoint";
      if (host.includes("onedrive.live.com")) return "OneDrive";
      return "Remote attachment";
    }

    function attachmentActionLabel(attachment) {
      const kindLabel = attachmentKindLabel(attachment.kind);
      const sourceLabel = attachmentSourceLabel(attachment);
      if (sourceLabel === "Teams media") {
        if (kindLabel === "Photo") return "Open Secure Photo";
        if (kindLabel === "Video") return "Open Secure Video";
        return "Open Secure File";
      }
      if (sourceLabel === "SharePoint") return kindLabel === "File" ? "Open SharePoint File" : `Open SharePoint ${kindLabel}`;
      if (sourceLabel === "OneDrive") return kindLabel === "File" ? "Open OneDrive File" : `Open OneDrive ${kindLabel}`;
      if (kindLabel === "Photo") return "Open Remote Photo";
      if (kindLabel === "Video") return "Open Remote Video";
      return "Open Remote File";
    }

    function attachmentAccessNote(attachment) {
      const sourceLabel = attachmentSourceLabel(attachment);
      if (!String(attachment.url || attachment.preview_url || "").trim()) {
        return "No recoverable attachment link was found in the export.";
      }
      if (sourceLabel === "Teams media") {
        return "Teams media link. Microsoft sign-in may be required.";
      }
      if (sourceLabel === "SharePoint") {
        return "SharePoint link. Microsoft sign-in may be required.";
      }
      if (sourceLabel === "OneDrive") {
        return "OneDrive link. Microsoft sign-in may be required.";
      }
      return "Remote attachment link. Authentication may be required.";
    }

    function normalizeAttachmentRecord(value) {
      if (!value || typeof value !== "object") return null;
      const url = String(value.url || "").trim();
      const previewUrl = String(value.preview_url || value.previewUrl || "").trim();
      const name = stripFcs(value.name || "") || attachmentNameFromUrl(url || previewUrl) || attachmentKindLabel(value.kind || "file");
      const kind = normalizeAttachmentKind(value.kind, name, url || previewUrl);
      if (!url && !previewUrl) return null;
      return {
        id: String(value.id || "").trim(),
        name,
        url,
        preview_url: previewUrl,
        kind,
      };
    }

    function extractAttachmentsFromHtml(contentHtml) {
      const html = String(contentHtml || "").trim();
      if (!html) return [];
      const doc = new DOMParser().parseFromString(html, "text/html");
      const attachments = [];
      const seen = new Set();

      function addAttachment(record) {
        const normalized = normalizeAttachmentRecord(record);
        if (!normalized) return;
        const key = [normalized.id, normalized.url, normalized.preview_url, normalized.name].join("|");
        if (seen.has(key)) return;
        seen.add(key);
        attachments.push(normalized);
      }

      for (const node of doc.querySelectorAll("img[src]")) {
        const itemtype = String(node.getAttribute("itemtype") || "").toLowerCase();
        if (itemtype.includes("emoji")) continue;
        if (!(itemtype.includes("amsimage") || itemtype.includes("inlineimage") || itemtype.includes("amsvideo"))) continue;
        const src = String(node.getAttribute("src") || "").trim();
        addAttachment({
          id: node.getAttribute("itemid") || node.getAttribute("id") || "",
          name: node.getAttribute("alt") || node.getAttribute("title") || "",
          url: src,
          preview_url: src,
          kind: itemtype,
        });
      }

      for (const node of doc.querySelectorAll("a[href]")) {
        const itemtype = String(node.getAttribute("itemtype") || "").toLowerCase();
        if (!(itemtype.includes("hyperlink/files") || itemtype.includes("fileshyperlink"))) continue;
        const href = String(node.getAttribute("href") || "").trim();
        addAttachment({
          name: stripFcs(node.textContent || "") || node.getAttribute("title") || "",
          url: href,
          kind: itemtype,
        });
      }
      return attachments;
    }

    function messageAttachments(message) {
      if (!message || typeof message !== "object") return [];
      if (MESSAGE_ATTACHMENT_CACHE.has(message)) return MESSAGE_ATTACHMENT_CACHE.get(message);
      const base = (message.attachments || [])
        .map(normalizeAttachmentRecord)
        .filter(Boolean);
      const combined = [];
      const seen = new Set();
      for (const record of [...base, ...extractAttachmentsFromHtml(message.content_html || "")]) {
        const key = [record.id, record.url, record.preview_url, record.name].join("|");
        if (seen.has(key)) continue;
        seen.add(key);
        combined.push(record);
      }
      MESSAGE_ATTACHMENT_CACHE.set(message, combined);
      return combined;
    }

    function messageHasAttachments(message) {
      return messageAttachments(message).length > 0;
    }

    function messageAttachmentPreview(message) {
      const attachments = messageAttachments(message);
      if (!attachments.length) return "";
      if (attachments.length === 1) {
        const attachment = attachments[0];
        const kind = attachmentKindLabel(attachment.kind);
        if (attachment.name && attachment.name !== kind) {
          return `${kind}: ${attachment.name}`;
        }
        return `${kind} attachment`;
      }
      return `${attachments.length} attachments`;
    }

    function renderAttachmentCards(message) {
      const attachments = messageAttachments(message);
      if (!attachments.length) return "";
      return `
        <div class="attachment-stack">
          ${attachments.map(attachment => {
            const openUrl = attachment.url || attachment.preview_url || "";
            const kindLabel = attachmentKindLabel(attachment.kind);
            const actionLabel = attachmentActionLabel(attachment);
            const accessNote = attachmentAccessNote(attachment);
            return `
              <div class="attachment-card">
                <div class="attachment-meta">
                  <div>
                    <div class="attachment-kind">${escapeHtml(kindLabel)} Attachment</div>
                    <div class="attachment-name">${escapeHtml(attachment.name || `${kindLabel} attachment`)}</div>
                    ${accessNote ? `<div class="attachment-note">${escapeHtml(accessNote)}</div>` : ``}
                  </div>
                  ${openUrl ? `<div class="attachment-actions"><a class="call-link" href="${escapeHtml(openUrl)}" target="_blank" rel="noopener">${escapeHtml(actionLabel)}</a><button type="button" class="call-link copy-attachment-link" data-url="${escapeHtml(openUrl)}">Copy Link</button></div>` : ``}
                </div>
              </div>
            `;
          }).join("")}
        </div>
      `;
    }

    function displayMessageQuality(message) {
      if (messageHasAttachments(message) && message.quality !== "event") return "attachment";
      return message.quality || "";
    }

    function messageDisplaySender(message) {
      if (message.message_type === "ThreadActivity/AddMember") {
        return systemEventActorLabel();
      }
      if (message.message_type === "ThreadActivity/MemberJoined") {
        return systemEventActorLabel();
      }
      if (message.message_type === "ThreadActivity/DeleteMember") {
        return systemEventActorLabel();
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

    function hasCallDateRange() {
      return Boolean(state.callDateFrom || state.callDateTo);
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

    function messagePassesThreadDateRange(message) {
      const { start, end } = normalizedThreadDateRange();
      if (start === null && end === null) return true;
      const current = timeValue(message.timestamp);
      if (!Number.isFinite(current)) return false;
      if (start !== null && current < start) return false;
      if (end !== null && current > end) return false;
      return true;
    }

    function messagePassesCombinedDateRange(message) {
      return messagePassesSidebarDateRange(message) && messagePassesThreadDateRange(message);
    }

    function clearAllMessageDateRanges() {
      state.messageDateFrom = "";
      state.messageDateTo = "";
      state.threadDateFrom = "";
      state.threadDateTo = "";
      if (messageDateFrom) messageDateFrom.value = "";
      if (messageDateTo) messageDateTo.value = "";
    }

    function normalizedCallDateRange() {
      const start = dateBoundaryValue(state.callDateFrom, false);
      const end = dateBoundaryValue(state.callDateTo, true);
      if (start === null && end === null) return { start: null, end: null };
      if (start !== null && end !== null && start > end) {
        return { start: end, end: start };
      }
      return { start, end };
    }

    function callPassesDateRange(call) {
      const { start, end } = normalizedCallDateRange();
      if (start === null && end === null) return true;
      const current = timeValue(callPrimaryTimestamp(call));
      if (!Number.isFinite(current)) return false;
      if (start !== null && current < start) return false;
      if (end !== null && current > end) return false;
      return true;
    }

    function callSearchText(call) {
      const key = callKey(call);
      if (CALL_SEARCH_TEXT_CACHE.has(key)) return CALL_SEARCH_TEXT_CACHE.get(key);
      const parts = [
        callLabel(call),
        call.call_id,
        call.shared_correlation_id,
        displayCallDirection(call),
        displayCallDirectionLabel(call),
        displayCallState(call),
        displayCallStatus(call),
        displayCallGroupLabel(call),
        call.call_type,
        displayCallTypeLabel(call),
        displayCallQualityLabel(call),
        call.conversation_id,
        call.group_chat_thread_id,
        call.summary_text,
        call.meeting_subject,
        call.meeting_start_time,
        call.meeting_end_time,
        call.user_participation,
        callSideLabel(call, "originator"),
        call.originator_id,
        call.originator_endpoint,
        call.originator_phone_number,
        callSideLabel(call, "target"),
        call.target_id,
        call.target_endpoint,
        call.target_phone_number,
      ];
      const value = parts.filter(Boolean).join(" ").toLowerCase();
      CALL_SEARCH_TEXT_CACHE.set(key, value);
      return value;
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
          if (state.callGroup && callGroupKey(call) !== state.callGroup) {
            return false;
          }
          if (state.callDirection === "inbound") {
            if (displayCallDirection(call) !== "incoming") return false;
          } else if (state.callDirection === "outbound") {
            if (displayCallDirection(call) !== "outgoing") return false;
          } else if (state.callDirection === "declined") {
            const stateValue = displayCallState(call);
            if (stateValue !== "declined" && stateValue !== "rejected") return false;
          } else if (state.callDirection === "missed") {
            if (displayCallState(call) !== "missed") return false;
          }
          if (state.callSearch && !callSearchText(call).includes(state.callSearch)) {
            return false;
          }
          if (!callPassesDateRange(call)) {
            return false;
          }
          return true;
        })
        .sort((left, right) => callPrimaryTimestampValue(right) - callPrimaryTimestampValue(left));
    }

    function searchMessageText(thread, message) {
      if (MESSAGE_SEARCH_TEXT_CACHE.has(message)) {
        return MESSAGE_SEARCH_TEXT_CACHE.get(message);
      }
      const linkedCall = findLinkedCall(message);
      const attachments = messageAttachments(message);
      const parts = [
        thread.label,
        thread.id,
        ...(thread.participants || []),
        messageDisplaySender(message),
        message.message_type,
        message.content_text,
        message.content_html,
        displayMessageType(message),
        displayMessageQuality(message),
        ...attachments.map(attachment => attachment.name),
        ...attachments.map(attachment => attachment.url),
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
      if (message.message_type === "ThreadActivity/MemberJoined") {
        const parsed = parseMemberJoinedEvent(message);
        parts.push(parsed.actorName, ...(parsed.joinedNames || []));
      }
      if (message.message_type === "ThreadActivity/DeleteMember") {
        const parsed = parseDeleteMemberEvent(message);
        parts.push(parsed.actorName, ...(parsed.removedNames || []));
      }
      const value = parts.filter(Boolean).join(" ").toLowerCase();
      MESSAGE_SEARCH_TEXT_CACHE.set(message, value);
      return value;
    }

    function includeMessageInSearchResults(message) {
      if (!message) return false;
      if (message.message_type === "ThreadActivity/AddMember") return false;
      if (message.message_type === "Call/History") return false;
      if (message.synthetic_call) return false;
      return true;
    }

    function renderMessageBody(message, thread, useHighlight = false) {
      if (message.synthetic_call || message.message_type === "Event/Call") {
        return renderCallEventBody(message);
      }
      if (isMembershipEvent(message)) {
        return renderMembershipEventBody(message, thread);
      }
      if (messageHasAttachments(message)) {
        const textHtml = message.content_text
          ? `<div class="body-text">${useHighlight ? highlightSearchHtml(message.content_text || "") : escapeHtml(message.content_text || "")}</div>`
          : ``;
        return `<div class="body attachment-body">${textHtml}${renderAttachmentCards(message)}</div>`;
      }
      return `<div class="body">${useHighlight ? highlightSearchHtml(message.content_text || "") : escapeHtml(message.content_text || "")}</div>`;
    }

    function renderMessageRow(message, thread, options = {}) {
      const useHighlight = Boolean(options.highlight);
      const isFocused = Boolean(options.focused);
      const isSearchMatch = Boolean(options.searchMatch);
      const classes = [
        "msg",
        message.quality === "event" ? "event" : "",
        isFocused ? "focused" : "",
        isSearchMatch ? "search-match" : "",
      ].filter(Boolean).join(" ");
      const senderHtml = useHighlight
        ? highlightSearchHtml(messageDisplaySender(message))
        : escapeHtml(messageDisplaySender(message));
      const typeHtml = useHighlight
        ? highlightSearchHtml(displayMessageType(message))
        : escapeHtml(displayMessageType(message));
      const bodyHtml = renderMessageBody(message, thread, useHighlight);
      const hiddenMeta = messageHiddenMeta(message);
      const openHidden = Boolean((hiddenMeta && hiddenMeta.count > 0) && (state.expandHiddenData || isFocused));
      const shellAttrs = `class="${escapeHtml(classes)}${hiddenMeta && hiddenMeta.count > 0 ? " msg-collapsible" : ""}" data-call-key="${escapeHtml(messageLinkedCallKey(message))}" data-timestamp="${escapeHtml(String(timeValue(message.timestamp)))}" style="--msg-accent:${escapeHtml(messageAccent(message))}"`;

      if (hiddenMeta && hiddenMeta.count > 0) {
        return `
          <details ${shellAttrs} ${openHidden ? "open" : ""}>
            <summary>
              <div class="head">
                <span class="sender">${senderHtml}</span>
                <span>${escapeHtml(fmt(message.timestamp))}</span>
              </div>
              <div class="head">
                <span>${typeHtml}</span>
                <span>${escapeHtml(displayMessageQuality(message))}</span>
              </div>
              <div class="msg-summary-note">${escapeHtml(hiddenMeta.label)}</div>
            </summary>
            <div class="msg-body-wrap">
              ${bodyHtml}
            </div>
          </details>
        `;
      }

      return `
        <div ${shellAttrs}>
          <div class="head">
            <span class="sender">${senderHtml}</span>
            <span>${escapeHtml(fmt(message.timestamp))}</span>
          </div>
        <div class="head">
          <span>${typeHtml}</span>
          <span>${escapeHtml(displayMessageQuality(message))}</span>
        </div>
        ${bodyHtml}
      </div>
      `;
    }

    function searchResultGroups() {
      const groups = [];
      if (!state.messageSearch) return groups;

      for (const thread of filteredThreads()) {
        const timeline = sortedThreadTimelineItems(thread)
          .filter(messagePassesSidebarDateRange)
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
            <span>${escapeHtml(fmt(callPrimaryTimestamp(call)))}</span>
          </div>
        </div>
      `).join("") || `<div class="empty">No calls match the current filters.</div>`;
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
                ${group.messages.map(message => renderMessageRow(message, group.thread, {
                  highlight: true,
                  searchMatch: textMatchesSearch(searchMessageText(group.thread, message)),
                })).join("")}
              </div>
            </div>
            ${index < groups.length - 1 ? `<div class="search-group-divider"></div>` : ``}
          `).join("") || `<div class="empty">No timeline items match the current search.</div>`}
        </div>
      `;
      initExpandableLists(contentPanel);
      initParticipantChatLinks(contentPanel);
      initAttachmentActions(contentPanel);

      [...contentPanel.querySelectorAll(".open-search-thread")].forEach(element => {
        element.addEventListener("click", () => {
          state.view = "messages";
          state.threadId = element.dataset.threadId;
          state.messageViewFilter = element.dataset.viewFilter || "all";
          state.threadDateFrom = state.messageDateFrom || "";
          state.threadDateTo = state.messageDateTo || "";
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
      const thread = THREAD_BY_ID.get(state.threadId);
      if (!thread) {
        contentPanel.innerHTML = `<div class="empty">Select a conversation to view its data.</div>`;
        return;
      }
      const meta = thread.metadata || {};
      const allMessages = sortedThreadTimelineItems(thread);
      const rangeMessages = allMessages.filter(messagePassesCombinedDateRange);
      const linkedCalls = linkedCallsForThread(thread);
      const linkedCallCount = linkedCalls.length;
      const meetingCall = linkedCalls.find(call => call.meeting_subject || call.meeting_start_time || call.meeting_end_time) || {};
      const messages = rangeMessages.filter(messagePassesFilter);
      const firstMessage = messages.length ? messages[messages.length - 1] : null;
      const lastMessage = messages[0] || null;
      const eventCount = messages.filter(message => message.quality === "event").length;
      const curatedCount = messages.filter(message => message.quality !== "event").length;
      const meeting = thread.meeting || {};
      const meetingSubject = stripFcs(meeting.subject || meetingCall.meeting_subject || meta.topic || thread.label || "");
      const meetingStart = meeting.startTime || meetingCall.meeting_start_time || "";
      const meetingEnd = meeting.endTime || meetingCall.meeting_end_time || "";
      const showCompactChatMeta = ["chat_space", "thread"].includes(thread.category);
      contentPanel.innerHTML = `
        <div class="title">
          <h2>${escapeHtml(stripFcs(thread.label || thread.id))}</h2>
          <div class="chip">${escapeHtml(prettyCategory(thread.category))} | ${allMessages.length} timeline items</div>
        </div>
        <div class="participant-cloud">
          ${renderExpandableChipGroup(threadParticipantEntries(thread), {
            emptyLabel: "No participants were resolved.",
            enableChatLinks: ["team_chat", "thread"].includes(thread.category),
          })}
        </div>
        <div class="section-divider">
          <div class="section-label">Details</div>
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
              <button type="button" class="call-link toggle-raw-events">${state.expandHiddenData ? "Hide Hidden Data" : "Show Hidden Data"}</button>
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
            <button type="button" class="call-link clear-date-range"${(hasThreadDateRange() || hasSidebarMessageDateRange()) ? "" : " disabled"}>Clear Range</button>
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
          <div class="meta-block"><div class="k">Participants</div><div>${(thread.participants || []).length}</div></div>
          <div class="meta-block"><div class="k">Merged Calls</div><div>${escapeHtml(String(linkedCallCount))}</div></div>
          ${showCompactChatMeta ? "" : `<div class="meta-block"><div class="k">Meeting Subject</div><div>${escapeHtml(meetingSubject)}</div></div>`}
          ${showCompactChatMeta ? "" : `<div class="meta-block"><div class="k">Meeting Window</div><div>${escapeHtml(meetingStart ? `${fmt(meetingStart)} to ${fmt(meetingEnd)}` : "")}</div></div>`}
          <div class="meta-block"><div class="k">First Activity</div><div>${escapeHtml(fmt(firstMessage ? firstMessage.timestamp : null))}</div></div>
        </div>
        <div class="section-divider">
          <div class="section-label">Timeline</div>
        </div>
        <div class="messages">
          ${messages.map(message => renderMessageRow(message, thread, {
            focused: state.focusCallKey && messageLinkedCallKey(message) === state.focusCallKey,
          })).join("") || `<div class="empty">${(hasThreadDateRange() || hasSidebarMessageDateRange()) ? "No timeline items match the current date range and filters." : "No messages in this conversation."}</div>`}
        </div>
      `;
      initExpandableLists(contentPanel);
      initParticipantChatLinks(contentPanel);
      initAttachmentActions(contentPanel);
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
          clearAllMessageDateRanges();
          renderMessageList();
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
          state.expandHiddenData = !state.expandHiddenData;
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
      const call = CALLS_BY_KEY.get(state.callKey);
        if (!call) {
          contentPanel.innerHTML = `<div class="empty">Select a call to view its data.</div>`;
          return;
        }
      const callParticipantEntriesList = callParticipantEntries(call);
      const callMeetingWindow = call.meeting_start_time ? `${fmt(call.meeting_start_time)} to ${fmt(call.meeting_end_time)}` : "";
      const relatedThreads = preferredThreadsForCall(call);
      const jumpTargets = callJumpTargets(call);
      const callGroupLabel = displayCallGroupLabel(call);
      const callTypeLabel = displayCallTypeLabel(call);
      const callQualityLabel = displayCallQualityLabel(call);
      const showRawType = includeRawCallTypeChip(call);
      const topDetailChips = dedupe([
        callGroupLabel,
        showRawType ? callTypeLabel : "",
        callQualityLabel,
      ].filter(Boolean));
      contentPanel.innerHTML = `
        <div class="title">
          <h2>${escapeHtml(stripFcs(callLabel(call)))}</h2>
          <div class="chip">${escapeHtml(displayCallStatus(call))}</div>
        </div>
        <div class="chips">
          ${topDetailChips.map(label => `<div class="chip">${escapeHtml(label)}</div>`).join("")}
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
          <div class="meta-block"><div class="k">Direction</div><div>${escapeHtml(displayCallDirectionLabel(call))}</div></div>
          <div class="meta-block"><div class="k">State</div><div>${escapeHtml(displayCallStateLabel(call))}</div></div>
          <div class="meta-block"><div class="k">Type</div><div>${escapeHtml(callGroupLabel)}</div></div>
          ${showRawType ? `<div class="meta-block"><div class="k">Raw Type</div><div>${escapeHtml(callTypeLabel)}</div></div>` : ``}
          <div class="meta-block"><div class="k">Originator</div><div>${escapeHtml(callSideLabel(call, "originator"))}</div></div>
          ${(callSideLabel(call, "target")) ? `<div class="meta-block"><div class="k">Target</div><div>${escapeHtml(callSideLabel(call, "target"))}</div></div>` : ``}
          <div class="meta-block"><div class="k">Call Id</div><div class="mono">${escapeHtml(call.call_id || "")}</div></div>
          ${call.shared_correlation_id ? `<div class="meta-block"><div class="k">Shared Correlation Id</div><div class="mono">${escapeHtml(call.shared_correlation_id || "")}</div></div>` : ``}
          <div class="meta-block wide"><div class="k">Participants</div>${renderExpandableChipGroup(callParticipantEntriesList, { emptyLabel: stripFcs(callLabel(call)) || "No participants were resolved." })}</div>
          ${call.meeting_subject ? `<div class="meta-block"><div class="k">Meeting Subject</div><div>${escapeHtml(stripFcs(call.meeting_subject || ""))}</div></div>` : ``}
          ${callMeetingWindow ? `<div class="meta-block"><div class="k">Meeting Window</div><div>${escapeHtml(callMeetingWindow)}</div></div>` : ``}
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
      initExpandableLists(contentPanel);
      initParticipantChatLinks(contentPanel);
      [...contentPanel.querySelectorAll(".open-thread-record")].forEach(element => {
        element.addEventListener("click", () => {
          state.view = "messages";
          state.threadId = element.dataset.threadId;
          state.messageViewFilter = element.dataset.viewFilter || "all";
          state.threadDateFrom = state.messageDateFrom || "";
          state.threadDateTo = state.messageDateTo || "";
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

    threadList.addEventListener("click", event => {
      const row = event.target.closest(".list-row[data-id]");
      if (!row || !threadList.contains(row)) return;
      state.threadId = row.dataset.id;
      state.threadDateFrom = state.messageDateFrom || "";
      state.threadDateTo = state.messageDateTo || "";
      state.focusCallKey = null;
      state.focusTimestamp = null;
      renderMessageList();
      renderContent();
      scrollMainToTop();
    });

    callList.addEventListener("click", event => {
      const row = event.target.closest(".list-row[data-key]");
      if (!row || !callList.contains(row)) return;
      state.callKey = row.dataset.key;
      renderCallList();
      renderContent();
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

    messageDateFrom.addEventListener("change", event => {
      state.messageDateFrom = event.target.value || "";
      state.threadDateFrom = state.messageDateFrom;
      state.threadDateTo = state.messageDateTo;
      renderMessageList();
      renderContent();
    });

    messageDateTo.addEventListener("change", event => {
      state.messageDateTo = event.target.value || "";
      state.threadDateFrom = state.messageDateFrom;
      state.threadDateTo = state.messageDateTo;
      renderMessageList();
      renderContent();
    });

    clearMessageDateRange.addEventListener("click", () => {
      clearAllMessageDateRanges();
      renderMessageList();
      renderContent();
    });

    callSearch.addEventListener("input", event => {
      state.callSearch = event.target.value.trim().toLowerCase();
      renderCallList();
      renderContent();
    });

    callGroupFilter.addEventListener("change", event => {
      state.callGroup = event.target.value;
      renderCallList();
      renderContent();
    });

    callDirectionFilter.addEventListener("change", event => {
      state.callDirection = event.target.value;
      renderCallList();
      renderContent();
    });

    callDateFrom.addEventListener("change", event => {
      state.callDateFrom = event.target.value || "";
      renderCallList();
      renderContent();
    });

    callDateTo.addEventListener("change", event => {
      state.callDateTo = event.target.value || "";
      renderCallList();
      renderContent();
    });

    clearCallDateRange.addEventListener("click", () => {
      state.callDateFrom = "";
      state.callDateTo = "";
      callDateFrom.value = "";
      callDateTo.value = "";
      renderCallList();
      renderContent();
    });

    function callPrimaryTimestampValue(call) {
      const key = callKey(call);
      if (CALL_PRIMARY_TIME_VALUE_CACHE.has(key)) {
        return CALL_PRIMARY_TIME_VALUE_CACHE.get(key);
      }
      const value = timeValue(callPrimaryTimestamp(call));
      CALL_PRIMARY_TIME_VALUE_CACHE.set(key, value);
      return value;
    }

    function scheduleIdleWork(callback) {
      if (window.requestIdleCallback) {
        window.requestIdleCallback(callback, { timeout: 250 });
        return;
      }
      window.setTimeout(() => callback({ didTimeout: true, timeRemaining: () => 0 }), 48);
    }

    function primeCallCaches(deadline) {
      const calls = DATA.calls || [];
      let index = primeCallCaches.index || 0;
      let processed = 0;
      const maxBatch = deadline.didTimeout ? 12 : Number.POSITIVE_INFINITY;
      while (index < calls.length && processed < maxBatch && (deadline.didTimeout || deadline.timeRemaining() > 4)) {
        isPhoneCall(calls[index]);
        index += 1;
        processed += 1;
      }
      primeCallCaches.index = index;
      if (index < calls.length) {
        scheduleIdleWork(primeCallCaches);
      }
    }

    scheduleIdleWork(primeCallCaches);
    renderView();
  </script>
</body>
</html>
"""


def load_export(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


SUMMARY_FIELDS = (
    "threads_total",
    "messages_total",
    "calls_total",
)
PROFILE_FIELDS = (
    "oid",
    "display_name",
)
THREAD_FIELDS = (
    "id",
    "category",
    "label",
    "message_count",
    "messages",
    "metadata",
    "meeting",
    "participant_ids",
    "participants",
)
THREAD_METADATA_FIELDS = ("topic",)
THREAD_MEETING_FIELDS = ("subject", "startTime", "endTime")
MESSAGE_FIELDS = (
    "id",
    "timestamp",
    "sender_display_name",
    "sender_id",
    "message_type",
    "content_html",
    "content_text",
    "attachments",
    "quality",
)
MESSAGE_ATTACHMENT_FIELDS = ("id", "name", "url", "preview_url", "kind")
CALL_FIELDS = (
    "call_id",
    "call_state",
    "call_type",
    "connect_time",
    "conversation_id",
    "direction",
    "end_time",
    "group_chat_thread_id",
    "meeting_end_time",
    "meeting_series_kind",
    "meeting_start_time",
    "meeting_subject",
    "originator_display_name",
    "originator_endpoint",
    "originator_id",
    "originator_phone_number",
    "participant_display_names",
    "participant_ids",
    "participant_sessions",
    "quality",
    "shared_correlation_id",
    "start_time",
    "summary_text",
    "target_display_name",
    "target_endpoint",
    "target_id",
    "target_phone_number",
    "user_participation",
)
CALL_PARTICIPANT_SESSION_FIELDS = ("id", "display_name")


def pick_fields(payload: dict | None, allowed_fields: tuple[str, ...]) -> dict:
    source = payload or {}
    return {key: source[key] for key in allowed_fields if key in source}


def build_browser(export_data: dict, output_path: Path) -> None:
    threads = []
    for thread in export_data.get("threads") or []:
        row = pick_fields(thread, THREAD_FIELDS)
        row["metadata"] = pick_fields(row.get("metadata"), THREAD_METADATA_FIELDS)
        row["meeting"] = pick_fields(row.get("meeting"), THREAD_MEETING_FIELDS)
        row["messages"] = [
            {
                **pick_fields(message, MESSAGE_FIELDS),
                "attachments": [
                    pick_fields(attachment, MESSAGE_ATTACHMENT_FIELDS)
                    for attachment in (message.get("attachments") or [])
                ],
            }
            for message in (thread.get("messages") or [])
        ]
        row["csv_path"] = thread_csv_path(thread)
        threads.append(row)

    calls = []
    for call in export_data.get("calls") or []:
        row = pick_fields(call, CALL_FIELDS)
        row["participant_sessions"] = [
            pick_fields(session, CALL_PARTICIPANT_SESSION_FIELDS)
            for session in (call.get("participant_sessions") or [])
        ]
        calls.append(row)

    payload = {
        "summary": pick_fields(export_data.get("summary"), SUMMARY_FIELDS),
        "profile": pick_fields(export_data.get("profile"), PROFILE_FIELDS),
        "threads": threads,
        "calls": calls,
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
