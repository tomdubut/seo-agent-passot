# seo-agent-passot

A manually-run SEO audit tool for passot.co.jp (WordPress + Elementor, EN
pages under `/en/`, JP pages at the root). Crawls the site via
`sitemap.xml`, extracts on-page SEO signals, flags issues, compares EN/JP
counterparts, and writes everything to a single XLSX workbook.

No accounts, API keys, or scheduling required for the core audit — it
only reads pages that are already public on the site. Two optional modes
add real data and do need accounts: `--ai-analysis` (an Anthropic API
key) and `--search-data` (a Google service account with Search Console +
GA4 access) — see below for both. If you don't have Search Console admin
access, `--gsc-import` lets you import a manually-exported report instead
of using the API — see below.

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

This crawls `https://www.passot.co.jp` using `https://www.passot.co.jp/sitemap_index.xml`
and writes `passot_seo_audit.xlsx` in the current directory. A full run of
~100-120 pages (EN + JP) takes a few minutes at the default request pace.

Useful flags (see `python seo_audit.py --help` for the full list):

| Flag | Purpose |
|---|---|
| `--sitemap-url URL` | Point at a different sitemap (e.g. if it's `/sitemap.xml` instead of `/sitemap_index.xml` — check WP admin / your SEO plugin to confirm) |
| `--max-pages N` | Crawl only the first N pages — good for a quick smoke test before a full run |
| `--skip-broken-links` | Skip internal link status checks (faster, but you lose that check) |
| `--output FILE.xlsx` | Change the output filename |
| `--thin-words N` / `--thin-chars N` | Fallback thin-content thresholds for pages that don't match a known page type below (default: 300 words for EN, 600 characters for JP) |
| `--delay SECONDS` | Seconds between requests (default 0.8s — increase if the site starts rate-limiting you) |
| `--debug-url URL` | Fetch a single URL and print exactly which container was used for word/char count, and why — use this to troubleshoot an unexpected 0 or very low count instead of running a full audit |
| `--ai-analysis` | Also run an AI (Claude) review of keyword fit, meta quality, and content quality per page — see below. Requires `ANTHROPIC_API_KEY` in the environment. |
| `--ai-model MODEL` | Model to use for `--ai-analysis` (default: `claude-sonnet-5`) |
| `--ai-keyword-map FILE.csv` | CSV with `url,keyword` columns giving explicit target keywords for specific pages — see below |
| `--search-data` | Pull real per-page Search Console + GA4 data — see below. Needs a Google service account. |
| `--gsc-site-url URL` | Your verified Search Console property, e.g. `https://www.passot.co.jp/` or `sc-domain:passot.co.jp` |
| `--ga4-property-id ID` | Your GA4 numeric Property ID (Admin → Property Settings) |
| `--google-credentials-file PATH` | Path to the service account JSON key. If omitted, falls back to `GOOGLE_SERVICE_ACCOUNT_JSON` or `GOOGLE_APPLICATION_CREDENTIALS` in the environment. |
| `--search-data-days N` | Lookback window in days for Search Console/GA4 data (default: 28) |
| `--gsc-import FILE` | Import a manually-exported Search Console Performance report (CSV or XLSX) instead of using the API — see below. No Google credentials needed, and works with or without `--search-data`. |

If the site blocks the crawler (403s), it's likely a WAF/bot-protection
rule reacting to the request pattern or User-Agent — try `--delay 2` first,
or check whether your host's security plugin is challenging automated
requests from your IP.

## About the thin-content thresholds

Google has no official minimum word count — word count is not a ranking
factor. The threshold here is a heuristic proxy ("very short pages are
often not substantive"), and a single number for the whole site is a blunt
instrument: a contact page and a competitive blog article shouldn't be
held to the same bar. So the check classifies each page by URL pattern
and applies a different threshold per type:

| Page type | How it's detected | EN threshold | JP threshold |
|---|---|---|---|
| Homepage | path is `/` or the EN prefix root | exempt | exempt |
| Blog Index | path is exactly `/column/` or `/en/blog/` | exempt | exempt |
| Product | first path segment is `products` or `retail-displays` | 150 words | 300 chars |
| Company/About | first path segment is `company`, `about`, `our-process`, etc. | 150 words | 300 chars |
| Utility (contact/FAQ/policy) | path contains one of those keywords | 80 words | 150 chars |
| Blog/Column Article | first segment starts with `column`, or the page's matched EN/JP counterpart does | 600 words | 1000 chars |
| Other (unclassified) | fallback | `--thin-words` (300) | `--thin-chars` (600) |

EN pages with no keyword match in their own URL (common with SEO-friendly
slugs) inherit their JP counterpart's classification via the already-matched
EN/JP pair, rather than guessing from scratch. The "All Pages" tab shows
which page type and threshold was applied to every page, and each Thin
Content issue names the page type it was judged against — check both if a
flag looks wrong, and adjust the rules in `PAGE_TYPE_BY_FIRST_SEGMENT`,
`UTILITY_TOKENS`, etc. near the top of `seo_audit.py` to match your own
URL structure and content strategy.

## AI content & keyword review (optional, `--ai-analysis`)

Everything else in this tool is a mechanical check (presence, length,
duplication) — it can't judge whether a page is actually *good* for the
keyword it's trying to rank for. `--ai-analysis` adds that: it sends each
page's title, meta description, headings, and main content to Claude and
asks for a genuine qualitative judgment — keyword/topic fit, whether the
title and meta description are compelling (not just present), whether the
content substantively covers the topic, and concrete suggestions. Results
land in their own **AI Content & Keyword Analysis** tab, sorted worst-rated
first and color-coded on a five-level scale (Critical → Poor → Needs
Improvement → Good → Excellent), plus a short rollup in the Executive
Summary. Assessments and suggestions are always written in English
regardless of the page's language — only the identified keyword/topic
stays in the page's own language, since that's what a real searcher would
actually type.

**Setup:**
1. Requires an Anthropic API key (console.anthropic.com).
2. Locally: `export ANTHROPIC_API_KEY=your-key` before running.
3. Via GitHub Actions: add it as a repo secret named `ANTHROPIC_API_KEY`
   (Settings → Secrets and variables → Actions → New repository secret),
   then tick "ai_analysis" when you run the workflow.

**Cost:** roughly $0.20–$1 for a full ~60-page run depending on the model
(Sonnet 5 is the default and lands around $0.40) — small, but not free
like the rest of the tool, and it's a real network call each time you run
it, so it's opt-in rather than bundled into every run.

**Target keywords:** by default the AI infers each page's apparent target
topic from its own content — useful for a general quality read, but it's
judging the page against a guess it made from that same page. Three
sources are checked in priority order: (1) `--ai-keyword-map keywords.csv`
(`url,keyword` columns) if you've specified real targets yourself, (2) if
`--search-data` is also enabled, the page's actual top Google Search query
by impressions — no guessing at all, judged against what people really
type, (3) otherwise, AI inference from the page's own content as a
fallback. Note: there's no `<meta name="keywords">` or Yoast/RankMath
"focus keyphrase" to crawl for this automatically — the focus keyphrase
lives only in the WordPress editor, never in the public HTML, and the old
`keywords` meta tag has been ignored by Google since ~2009.

## Real search & traffic data (optional, `--search-data`)

Everything else in this tool — even `--ai-analysis` on its own — is
judging pages in a vacuum, with no visibility into whether anyone actually
finds or visits them. `--search-data` pulls real per-page data from two
Google APIs:

- **Search Console**: clicks, impressions, click-through rate, and average
  ranking position per page, plus each page's actual top search query.
  Lands in a new **Search Performance (GSC)** tab.
- **GA4**: sessions, pageviews, engagement rate, bounce rate, and
  conversions per page. Lands in a new **Traffic (GA4)** tab.

Both also feed back into the rest of the report: the Executive Summary
gets a "Real traffic & search performance" section ranking your top 10
pages by actual traffic against their open issue count (so you know which
fixes matter most), plus a callout for any URLs getting real search
impressions that weren't found in your current sitemap crawl (a signal of
orphaned or removed-but-still-valuable pages). If `--ai-analysis` is also
on, each page's real top search query and performance numbers get passed
to the AI instead of letting it guess — see "Target keywords" above.

**Setup** (more involved than `--ai-analysis` — this uses two separate
Google products):

1. In [Google Cloud Console](https://console.cloud.google.com/), create or
   pick a project, then enable two APIs for it: **Search Console API** and
   **Google Analytics Data API** (APIs & Services → Library → search for
   each → Enable). Both are free — neither requires a billing account on
   the project, and this tool's usage (a handful of report queries per
   run) is nowhere near either API's free quota regardless.
2. Create a **service account** (IAM & Admin → Service Accounts → Create
   Service Account) — this is a "robot" identity, separate from your own
   Google login, used for unattended access. Give it any name (e.g.
   `seo-audit-reader`); it doesn't need any project-level role.
3. Create a JSON key for it (open the service account → Keys → Add Key →
   Create new key → JSON) and download the file. **Treat this like a
   password** — it grants read access to whatever you attach it to below.
4. Grant that service account read access in **both** products, using the
   email address shown on the service account (looks like
   `seo-audit-reader@your-project.iam.gserviceaccount.com`):
   - **Search Console**: Settings → Users and permissions → Add user →
     paste the service account email → "Restricted" permission is enough.
   - **GA4**: Admin → Property Access Management → paste the service
     account email → "Viewer" role is enough.
5. Find your **GA4 Property ID** (Admin → Property Settings — a short
   number, not the "Measurement ID" that starts with `G-`) and your
   **verified Search Console property URL** (exactly as shown in Search
   Console, e.g. `https://www.passot.co.jp/` or `sc-domain:passot.co.jp`).

**Running it:**
- Locally: `export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat your-key.json)"` (or
  pass `--google-credentials-file your-key.json` directly), then
  `python seo_audit.py --search-data --gsc-site-url "https://www.passot.co.jp/" --ga4-property-id 123456789`
- Via GitHub Actions: add the entire JSON key file's contents as a repo
  secret named `GOOGLE_SERVICE_ACCOUNT_JSON` (Settings → Secrets and
  variables → Actions → New repository secret, paste the whole JSON), then
  tick "search_data" and fill in `ga4_property_id` when you run the
  workflow (`gsc_site_url` already defaults to the site).

You can use just one of the two products (e.g. `--search-data
--gsc-site-url ...` with no `--ga4-property-id`) if you only have access
set up for one.

### No Search Console admin access? Import the export instead (`--gsc-import`)

If you don't have (or can't get) admin access to Search Console — for
example, someone else owns the property and only gave you GA4 — you don't
need the API or a service account for the Search Console half at all.
Search Console lets anyone with at least "Restricted" access export the
Performance report as a file, which you can hand off and import directly:

1. Whoever has Search Console access opens **Performance** → sets the date
   range you want → clicks **Export** → **Download Excel** (a single
   `.xlsx` file with one tab per dimension — this tool reads the "Pages"
   tab). CSV also works, but comes as a zip of several files, so Excel is
   simpler to hand off as one file.
2. Get you that file, then import it — how depends on where you run the
   audit:
   - **Locally:** save the file anywhere on your machine and run
     `python seo_audit.py --gsc-import "Search Console Performance.xlsx"`
   - **Via GitHub Actions:** `workflow_dispatch` inputs can't take file
     uploads directly, so upload the file into the repo first — open the
     [`gsc-exports/`](gsc-exports/) folder on GitHub → **Add file** →
     **Upload files** → drag in the export → commit. Then, when running
     the "SEO Audit" workflow, set `gsc_import_path` to the file's path
     (e.g. `gsc-exports/Performance.xlsx`).

This works standalone (no `--search-data`, no Google credentials, nothing
else needed) and populates the same **Search Performance (GSC)** tab and
Executive Summary sections as the API path. You can also combine it with
`--search-data --ga4-property-id ...` to get GA4 via the API while GSC
comes from the imported file — in that case `--gsc-import` takes
precedence over `--gsc-site-url` for the Search Console data specifically,
so you don't need Search Console API access at all.

**Trade-off:** the exported report doesn't include each page's individual
top query (Search Console only ties queries to pages in the live API
response, not the flat export), so `--ai-analysis` keyword grounding falls
back to `--ai-keyword-map` or its own inference for keyword matching, even
though clicks/impressions/CTR/position are still real numbers either way.

**Limitations:** Search Console data has a 2-3 day reporting lag and
suppresses very low-volume queries for privacy, so long-tail/low-traffic
pages will show sparser data. GA4 only has data from whenever tracking was
actually installed — if it's new, the lookback window won't have much
history. Neither source is fabricated or estimated; if a page has no data,
it's shown as blank rather than guessed at.

## What it checks

- **Per page:** title tag, meta description, canonical tag, H1/H2
  structure, image alt-text coverage, internal link count in and out
  (within main content), word count (EN) / character count (JP),
  last-modified date (from the HTTP header if present, else the sitemap),
  JSON-LD structured data types present, Open Graph / Twitter Card tags.
- **Site-wide:** missing/duplicate meta descriptions, missing/duplicate
  titles, missing/multiple H1s, missing alt text, thin content, broken
  internal links, missing/mismatched canonical tags, sitemap URLs that
  redirect, missing structured data, missing Open Graph tags, pages with
  very few internal links pointing to them ("weakly linked").
- **EN/JP comparison:** pages are paired via `hreflang` tags first, falling
  back to URL pattern matching (`/en/x/` ↔ `/x/`). For each pair: reciprocal
  hreflang presence, meta description present on only one side, byte-identical
  (likely untranslated) titles/meta descriptions, and H1 count mismatches.
  Pages with no counterpart in the other language are listed separately
  (not necessarily wrong — e.g. a JP-only blog post — but worth a look) and
  called out as a dedicated "translation / content parity gap" section in
  the Executive Summary, not just buried as Info-severity rows.

## Output

`passot_seo_audit.xlsx` with an **Executive Summary** tab first — site
totals, issues grouped by category with a recommended action and effort
tag, a table of images missing alt text on multiple pages (fix the shared
component once instead of page by page), the translation/content-parity
gap list, and an auto-generated fix-first priority list. Then an **Issues
Summary** tab with every flagged issue individually, sorted by severity
and color-coded. Then one tab per check type for the full underlying data.

Severity guide:
- **High** — page unreachable, missing title, missing canonical, broken
  internal link, missing/incomplete hreflang pairing between EN/JP
  counterparts.
- **Medium** — missing/duplicate meta description, duplicate title,
  canonical pointing elsewhere, missing/multiple H1, thin content,
  EN/JP content inconsistencies (untranslated fields, meta present on
  only one side).
- **Low** — missing alt text, sitemap URL redirects, minor EN/JP
  structural differences (e.g. H1 count), missing structured data, missing
  Open Graph tags, weakly-linked pages.
- **Info** — a page has no counterpart in the other language (informational,
  not necessarily a problem — see the Executive Summary's dedicated
  translation-gap section instead of hunting for these rows).

## What it can't tell you (and why)

This audits pages that already exist. With `--search-data` on, you do get
real Search Console query data — but only for queries where your pages
already get *some* impressions. It still can't discover keywords you have
**zero** visibility into (topics you'd need a dedicated keyword-research
tool to find, since Search Console only reports on queries that already
surface one of your pages at all), and it doesn't estimate search volume
for anything — no fabricated numbers, ever. `--ai-analysis` judges how
well an *existing* page fits a keyword (your own, Search Console's real
top query, or an AI guess as a last resort) — it doesn't invent net-new
content ideas. The translation-parity, weak-internal-linking, and (with
`--search-data`) orphan-page checks are the closest this tool gets to
"opportunities" from real data alone.

## Known limitations

- **Discovery is sitemap-only** — pages not listed in the sitemap (orphans,
  noindexed pages) won't be found.
- **No JS rendering** — pages are fetched as static HTML (`requests`, no
  headless browser). Elementor renders server-side so this is normally
  fine, but any content injected purely client-side won't be captured. If
  a page's extracted word/char count looks suspiciously low, run
  `python seo_audit.py --debug-url "https://www.passot.co.jp/the/page/"`
  to see exactly which container was used and why, and whether the real
  text exists elsewhere in the raw HTML (a selection issue) or genuinely
  isn't there as crawlable text (e.g. content delivered as images).
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
