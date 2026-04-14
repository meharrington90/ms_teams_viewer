# MS Teams Viewer

This software reads Microsoft Teams local IndexedDB data, builds a structured export, writes CSV files, and creates a standalone HTML viewer for browsing messages and calls.

## How To Run

Use one of the launchers from the repo root:

- macOS: `MacOS Launcher.command`
- Windows: `Windows Launcher.bat`

After a successful run, the launcher will automatically open the generated HTML viewer in your default browser.

If you want to run it manually:

```bash
python3 -m venv .venv
. .venv/bin/activate
python python/bootstrap_dependencies.py
python python/run_teams_pipeline.py --open-browser
```

## Important Details

- The launchers will use local Python if it is installed.
- If Python is not installed, the launchers will try to create a self-contained runtime automatically.
- First-run setup requires internet access so dependencies can be downloaded.
- The software does not require admin/root access.
- It reads the local Teams profile automatically on macOS and Windows.
- Reading live Teams data is best-effort. If Teams is actively running and changing files, results can vary.

## Output

The default output folder is:

`teams_pipeline_output/`

Main files created there:

- `teams_ccl_summary.json`
- `teams_ccl_canonical_v1.json`
- `teams_ccl_csv_v1/`
- `teams_ccl_browser_v1.html`

## Notes

- Generated outputs can contain sensitive Teams data.
- `.venv/`, `.tools/`, and generated outputs are excluded by `.gitignore`.

## License

MIT — see [LICENSE](./LICENSE)
