# seo-agent-passot

A manually-run SEO audit tool for passot.co.jp (WordPress + Elementor, EN
pages under `/en/`, JP pages at the root). Crawls the site via
`sitemap.xml`, extracts on-page SEO signals, flags issues, compares EN/JP
counterparts, and writes everything to a single XLSX workbook.

No accounts, API keys, or scheduling required — this only reads pages
that are already public on the site.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
python seo_audit.py
```

This crawls `https://passot.co.jp` using `https://passot.co.jp/sitemap_index.xml`
and writes `passot_seo_audit.xlsx` in the current directory. A full run of
~100-120 pages (EN + JP) takes a few minutes at the default request pace.

Useful flags (see `python seo_audit.py --help` for the full list):

| Flag | Purpose |
|---|---|
| `--sitemap-url URL` | Point at a different sitemap (e.g. if it's `/sitemap.xml` instead of `/sitemap_index.xml` — check WP admin / your SEO plugin to confirm) |
| `--max-pages N` | Crawl only the first N pages — good for a quick smoke test before a full run |
| `--skip-broken-links` | Skip internal link status checks (faster, but you lose that check) |
| `--output FILE.xlsx` | Change the output filename |
| `--thin-words N` / `--thin-chars N` | Adjust the thin-content thresholds (default: 300 words for EN, 600 characters for JP) |
| `--delay SECONDS` | Seconds between requests (default 0.8s — increase if the site starts rate-limiting you) |

If the site blocks the crawler (403s), it's likely a WAF/bot-protection
rule reacting to the request pattern or User-Agent — try `--delay 2` first,
or check whether your host's security plugin is challenging automated
requests from your IP.

## What it checks

- **Per page:** title tag, meta description, canonical tag, H1/H2
  structure, image alt-text coverage, internal link count (within main
  content), word count (EN) / character count (JP), last-modified date
  (from the HTTP header if present, else the sitemap).
- **Site-wide:** missing/duplicate meta descriptions, missing/duplicate
  titles, missing/multiple H1s, missing alt text, thin content, broken
  internal links, missing/mismatched canonical tags, sitemap URLs that
  redirect.
- **EN/JP comparison:** pages are paired via `hreflang` tags first, falling
  back to URL pattern matching (`/en/x/` ↔ `/x/`). For each pair: reciprocal
  hreflang presence, meta description present on only one side, byte-identical
  (likely untranslated) titles/meta descriptions, and H1 count mismatches.
  Pages with no counterpart in the other language are listed separately
  (not necessarily wrong — e.g. a JP-only blog post — but worth a look).

## Output

`passot_seo_audit.xlsx` with one tab per check type, plus an **Issues
Summary** tab at the front — every flagged issue in one place, sorted by
severity (High/Medium/Low/Info) and color-coded.

Severity guide:
- **High** — page unreachable, missing title, missing canonical, broken
  internal link, missing/incomplete hreflang pairing between EN/JP
  counterparts.
- **Medium** — missing/duplicate meta description, duplicate title,
  canonical pointing elsewhere, missing/multiple H1, thin content,
  EN/JP content inconsistencies (untranslated fields, meta present on
  only one side).
- **Low** — missing alt text, sitemap URL redirects, minor EN/JP
  structural differences (e.g. H1 count).
- **Info** — a page has no counterpart in the other language (informational,
  not necessarily a problem).

## Known limitations

- **Discovery is sitemap-only** — pages not listed in the sitemap (orphans,
  noindexed pages) won't be found.
- **No JS rendering** — pages are fetched as static HTML (`requests`, no
  headless browser). Elementor renders server-side so this is normally
  fine, but any content injected purely client-side won't be captured. If
  a page's extracted word count looks suspiciously low, that's the tell to
  check manually.
- **No indexing/ranking/traffic data** — that requires Google Search
  Console, which isn't wired up here.
- **No page speed / Core Web Vitals** — requires PageSpeed Insights API,
  not included.
- **No backlink data** — intentionally excluded; not built, not estimated.
- **No soft-404 detection** — a page returning HTTP 200 with an empty or
  "not found"-looking template will read as healthy; only real HTTP status
  codes are checked.
- **Japanese word count** uses character count (whitespace stripped) as a
  proxy, not real tokenization (no MeCab/fugashi dependency) — good enough
  as a thin-content signal, not a precise word count.
- **hreflang correctness** is checked for presence and reciprocity between
  matched pairs, not whether Google actually honors it.
