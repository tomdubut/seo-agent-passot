Drop Search Console Performance report exports here so the GitHub Actions
workflow can find them via the `gsc_import_path` input.

## How to add a file

1. Get the export from whoever has Search Console access: **Performance**
   tab → set your date range → **Export** → **Download Excel**.
2. On GitHub, open this folder (`gsc-exports/`) → **Add file** → **Upload
   files** → drag in the `.xlsx` (or `.csv`) → commit directly to `main`.
3. When running the "SEO Audit" workflow (Actions tab → Run workflow), set
   `gsc_import_path` to the file's path, e.g. `gsc-exports/Performance.xlsx`.

Uploaded files stay in the repo's git history. Only put files here if
that's acceptable for this repo's visibility (fine for a private repo).
