#!/usr/bin/env python3

import argparse
import json
import webbrowser
from pathlib import Path

from build_teams_ccl_browser import build_browser
from dump_teams_indexeddb_ccl import resolve_teams_source
from export_teams_ccl_canonical import build_export
from export_teams_ccl_csv import ensure_dir, export_calls, export_conversations
from export_teams_ccl_summary import build_summary


def write_json(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def print_stage(number: int, total: int, message: str) -> None:
    print()
    print(f"[{number}/{total}] {message}")


def run_pipeline(root: Path | str | None, output_root: Path) -> dict:
    source = resolve_teams_source(root)
    source_root = source.profile_root
    discovery_root = root if root is not None else source.search_root

    ensure_dir(output_root)
    csv_dir = output_root / "teams_ccl_csv_v1"
    ensure_dir(csv_dir)

    print("Starting Teams export pipeline")
    print(f"Source profile: {source_root}")

    print_stage(1, 3, "Scanning IndexedDB stores and building the summary report...")
    summary = build_summary(discovery_root, show_decode_errors=False)
    summary_path = output_root / "teams_ccl_summary.json"
    write_json(summary_path, summary)

    print_stage(2, 3, "Exporting the canonical JSON dataset and CSV files...")
    export_data = build_export(discovery_root, show_decode_errors=False)
    canonical_path = output_root / "teams_ccl_canonical_v1.json"
    write_json(canonical_path, export_data)
    export_conversations(export_data, csv_dir)
    export_calls(export_data, csv_dir)
    write_json(csv_dir / "export_summary.json", export_data.get("summary") or {})

    print_stage(3, 3, "Building the standalone browser viewer...")
    browser_path = output_root / "teams_ccl_browser_v1.html"
    build_browser(export_data, browser_path)

    manifest = {
        "platform": source.platform_name,
        "search_root": str(source.search_root),
        "source_root": str(source_root),
        "leveldb_path": str(source.leveldb_path),
        "blob_path": str(source.blob_path) if source.blob_path.exists() else None,
        "local_storage_dir": str(source.local_storage_dir) if source.local_storage_dir else None,
        "discovery_method": source.discovery_method,
        "artifacts": {
            "summary_json": str(summary_path),
            "canonical_json": str(canonical_path),
            "csv_dir": str(csv_dir),
            "browser_html": str(browser_path),
        },
    }
    write_json(output_root / "run_manifest.json", manifest)
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the full Teams CCL pipeline with automatic local-source discovery.")
    parser.add_argument("--root", help="Optional Teams source root, profile directory, copied evidence root, or IndexedDB directory.")
    parser.add_argument("--output-root", default="teams_pipeline_output", help="Directory for all generated artifacts.")
    parser.add_argument("--open-browser", action="store_true", help="Open the generated HTML viewer after a successful run.")
    args = parser.parse_args()

    output_root = Path(args.output_root).expanduser().resolve()
    manifest = run_pipeline(Path(args.root).expanduser().resolve() if args.root else None, output_root)
    browser_path = Path(manifest["artifacts"]["browser_html"])

    if args.open_browser:
        print()
        print("Opening the browser viewer...")
        webbrowser.open(browser_path.resolve().as_uri())

    print()
    print("Completed.")
    print(f"Output folder: {output_root}")
    print(f"Browser viewer: {browser_path}")


if __name__ == "__main__":
    main()
