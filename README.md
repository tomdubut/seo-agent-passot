# seo-agent-passot

A manually-run SEO audit tool for passot.co.jp (WordPress + Elementor, EN
pages under `/en/`, JP pages at the root). Crawls the site via
`sitemap.xml`, extracts on-page SEO signals, flags issues, compares EN/JP
counterparts, and writes everything to a single XLSX workbook.

No accounts, API keys, or scheduling required for the core audit — it
only reads pages that are already public on the site. An optional
`--ai-analysis` mode (see below) adds AI-judged keyword/content quality
and does need an Anthropic API key.

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
land in their own **AI Content & Keyword Analysis** tab, color-coded
Good/Needs Improvement/Poor, plus a short rollup in the Executive Summary.

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
judging the page against a guess it made from that same page. If you have
real target keywords in mind, pass `--ai-keyword-map keywords.csv` with
`url,keyword` columns; listed pages get judged against your actual intent
instead, which is far more useful for catching true gaps ("this page is
trying to rank for X but never actually says X"). Note: there's no
`<meta name="keywords">` or Yoast/RankMath "focus keyphrase" to crawl for
this automatically — the focus keyphrase lives only in the WordPress
editor, never in the public HTML, and the old `keywords` meta tag has been
ignored by Google since ~2009 — so a keyword list has to come from you.

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

This audits pages that already exist. Even with `--ai-analysis` turned on,
it does **not** attempt to identify content or pages you don't have yet
(keyword gaps, "what are people searching for that you don't rank for")
— that needs real search-demand data (Google Search Console query data or
a keyword research tool), which isn't wired up here, and fabricating
numbers that look like search volume would be worse than not having them.
`--ai-analysis` judges how well an *existing* page fits a keyword (inferred
or one you provide) — it doesn't discover keywords you're missing
entirely. The translation-parity and weak-internal-linking checks are the
closest this tool gets to "opportunities" from crawled data alone.

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
