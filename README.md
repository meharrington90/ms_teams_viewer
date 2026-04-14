# Michaelsoft Teams

Repository name: `michaelsoft_teams`

Michaelsoft Teams is a local Microsoft Teams artifact exporter and standalone browser viewer. It reads the newer Teams client IndexedDB profile, produces structured JSON and CSV exports, and builds a single-file HTML interface for reviewing messages, meetings, calls, and attachment references offline.

## Current Scope

- Automatically discovers the current Teams profile on macOS and Windows, or accepts an explicit `--root` that points at a copied evidence root, profile root, IndexedDB root, or `.leveldb` directory.
- Uses `ccl_chromium_reader` to decode structured IndexedDB stores from the Teams profile and matching blob directory.
- Builds a summary report, a canonical JSON export, flat CSVs, per-conversation CSVs, a run manifest, and a standalone HTML viewer.
- Extracts message attachments currently recoverable from replychains, including inline images, inline videos, and file/share links surfaced through message HTML or message properties.
- Merges structured call-history records with call events inferred from conversation messages so the browser can cross-link related calls and message timelines.
- Includes a browser UI with separate Messages and Calls views, search, type filters, date-range filters, linked call/message navigation, and direct links to conversation CSVs.

## How To Run

Use one of the launchers from the repo root:

- macOS: `MacOS Launcher.command`
- Windows: `Windows Launcher.bat`

Both launchers:

- create `.venv/` with local Python when available
- fall back to a bundled `uv` runtime when Python is not installed
- install dependencies on first run
- run `python/run_teams_pipeline.py --open-browser`
- pause at the end so you can review any errors

After a successful run, the generated browser viewer opens automatically in your default browser.

If you want to run it manually:

```bash
python3 -m venv .venv
. .venv/bin/activate
python python/bootstrap_dependencies.py
python python/run_teams_pipeline.py --open-browser
```

Useful options:

- `python python/run_teams_pipeline.py --root /path/to/profile-or-copy`
- `python python/run_teams_pipeline.py --output-root /path/to/output`
- `python python/run_teams_pipeline.py --open-browser`

You can also set `TEAMS_CCL_ROOT` to override the default source search root.

## Source Discovery

Default search roots:

- macOS: `~/Library/Containers/com.microsoft.teams2/Data/Library`
- Windows: `%APPDATA%\Microsoft\Teams`
- Windows: `%LOCALAPPDATA%\Microsoft\MSTeams`
- Windows: `%LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache`

Known profile layouts searched first:

- `EBWebView/WV2Profile_tfw`
- `EBWebView/Default`

The resolver can also work against copied evidence folders by recursively locating `https_teams.microsoft.com_0.indexeddb.leveldb`.

## Output

The default output folder is `teams_pipeline_output/`.

Main artifacts:

- `teams_ccl_summary.json`: high-level store counts and representative samples from key stores
- `teams_ccl_canonical_v1.json`: canonical export with profile info, threads, messages, calls, GUID directory, and limitations
- `teams_ccl_browser_v1.html`: single-file browser viewer branded as `Michaelsoft Teams`
- `run_manifest.json`: platform, discovery paths, discovery method, and artifact paths for the run
- `teams_ccl_csv_v1/conversation_manifest.csv`: one row per conversation/thread with participant and CSV metadata
- `teams_ccl_csv_v1/conversations_all_flat.csv`: flat all-messages export across conversations
- `teams_ccl_csv_v1/conversations/*.csv`: per-conversation CSVs
- `teams_ccl_csv_v1/call_history.csv`: flat call-history export
- `teams_ccl_csv_v1/export_summary.json`: summary block copied from the canonical export

## Browser Viewer

The generated HTML viewer currently supports:

- Messages and Calls tabs
- full-text search across conversation metadata, message text, attachment names, and linked call details
- message-type and call-type filters
- message and call date-range filters
- conversation timelines that merge messages with related call events where possible
- inline attachment cards for image, video, and file links
- call detail pages with linked conversation context
- lightweight browser payload trimming so the HTML ships only fields used by the UI

## Canonical Export Notes

The canonical JSON currently includes:

- normalized thread records with labels, categories, participants, metadata quality, meeting metadata, and ordered messages
- message records with sender info, text/HTML content, message type, attachment references, and quality classification
- structured call-history records plus inferred thread-event calls merged into the calls dataset
- profile data when it can be recovered from Teams local storage
- a GUID-to-name directory built from people, conversations, messages, and calls

## Requirements

- No admin or root access is required.
- First-run setup needs network access.
- Dependency bootstrap installs `ccl_chromium_reader` and related packages from GitHub source archives.
- The bundled runtime path uses Astral `uv` when local Python is unavailable.

## Limitations

- Reading a live Teams profile is best-effort. If Teams is running and mutating files during export, results can vary.
- A small number of IndexedDB keys can still report decode errors and may be absent from the export.
- Attachment coverage is currently limited to what can be recovered from replychain message content and message properties.
- The export still does not merge every auxiliary Teams store, including full reaction history and notification surfaces.
- Generated outputs can contain sensitive Teams content and metadata.

## License

MIT. See [LICENSE](./LICENSE).
