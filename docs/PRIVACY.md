# Privacy And Release Checklist

Use this checklist before publishing the repository to GitHub.

## Do Not Commit

- Real annual report HTML, HTM or PDF files.
- Internal company lists, screening workbooks, CUSIP mapping files or research notes.
- Generated Excel, CSV, JSON, JSONL or TXT outputs.
- LLM raw responses, prompts containing full report excerpts, run logs or progress files.
- `.env`, API keys, endpoint secrets, account ids or local credential files.
- Local absolute paths, user names, desktop paths or organization-specific directories.
- Python caches, virtual environments, packaged archives or temporary files.

## Safe To Commit

- Reusable Python scripts under `scripts/`.
- Public templates under `templates/`.
- Empty `.gitkeep` files in `data/downloads/` and `outputs/`.
- Documentation, license and dependency list.

## Recommended Release Check

Run these commands from the repository root before pushing:

```powershell
git status --short
git ls-files
rg -n "D:\\\\|C:\\\\|api_key|ARK_API_KEY|secret|token|password|\\.xlsx|\\.html|\\.pdf" .
```

Expected result:

- `git status` should not show generated data files.
- `git ls-files` should not include annual reports, workbooks or model outputs.
- `rg` should only find documentation examples or code that reads environment variables, not real secrets or local private paths.

The default `.gitignore` is intentionally strict. If examples are needed, create small synthetic files instead of using real filings or internal workbooks.
