#!/usr/bin/env python3
"""
SEO audit tool for passot.co.jp (WordPress + Elementor, EN at /en/, JP at root).

Discovers pages via sitemap.xml, crawls each one, extracts on-page SEO
signals, flags issues, compares EN/JP counterparts, and writes everything
to a single XLSX workbook.

Usage:
    pip install -r requirements.txt
    python seo_audit.py

Run `python seo_audit.py --help` for all options.
"""

import argparse
import re
import sys
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Config defaults (override via CLI flags — see --help)
# --------------------------------------------------------------------------

DEFAULT_BASE_URL = "https://www.passot.co.jp"
DEFAULT_SITEMAP_URL = "https://www.passot.co.jp/sitemap_index.xml"
DEFAULT_EN_PREFIX = "/en/"
DEFAULT_OUTPUT = "passot_seo_audit.xlsx"
DEFAULT_DELAY = 0.8
DEFAULT_TIMEOUT = 15
DEFAULT_THIN_WORDS_EN = 300
DEFAULT_THIN_CHARS_JA = 600
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 PassotSEOAudit/1.0"
)
# Sub-sitemaps whose filenames contain any of these are skipped (archive
# listings, not real content pages).
DEFAULT_SITEMAP_EXCLUDE = ["author-sitemap", "category-sitemap", "tag-sitemap", "feed"]
NON_HTML_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".ico",
    ".pdf", ".zip", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    ".mp4", ".mp3", ".css", ".js", ".xml", ".json",
}
MAIN_CONTENT_SELECTORS = ["main", "article", ".elementor", "#content", ".entry-content", ".site-content"]
BOILERPLATE_SELECTOR = "nav, header, footer, aside, script, style, form"

SITEMAP_NS = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}

HIGH, MEDIUM, LOW, INFO = "High", "Medium", "Low", "Info"
SEVERITY_ORDER = {HIGH: 0, MEDIUM: 1, LOW: 2, INFO: 3}


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------

@dataclass
class Page:
    url: str
    lang: str
    sitemap_lastmod: str = ""
    status_code: int = None
    final_url: str = ""
    redirected: bool = False
    fetch_error: str = ""
    title: str = ""
    meta_description: str = ""
    canonical: str = ""
    hreflang: dict = field(default_factory=dict)
    h1_list: list = field(default_factory=list)
    h2_list: list = field(default_factory=list)
    image_total: int = 0
    image_missing_alt: int = 0
    missing_alt_examples: list = field(default_factory=list)
    internal_link_count: int = 0
    internal_links: set = field(default_factory=set)
    all_internal_links: set = field(default_factory=set)
    word_count: int = None
    char_count: int = None
    http_last_modified: str = ""

    @property
    def ok(self):
        return not self.fetch_error and self.status_code == 200


@dataclass
class Issue:
    severity: str
    category: str
    url: str
    lang: str
    detail: str


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

def normalize_netloc(netloc):
    netloc = netloc.lower()
    if netloc.startswith("www."):
        netloc = netloc[4:]
    return netloc


def is_probably_html(url):
    path = urlparse(url).path
    ext = path.rsplit(".", 1)
    if len(ext) == 2 and f".{ext[1].lower()}" in NON_HTML_EXTENSIONS:
        return False
    return True


def normalize_link(href, page_url):
    href = (href or "").strip()
    if not href or href.startswith(("mailto:", "tel:", "javascript:", "#")):
        return None
    absolute = urljoin(page_url, href)
    absolute, _frag = urldefrag(absolute)
    return absolute


def resolve_links(link_tags, page_url, base_netloc):
    result = set()
    for a in link_tags:
        norm = normalize_link(a.get("href", ""), page_url)
        if not norm:
            continue
        if normalize_netloc(urlparse(norm).netloc) != base_netloc:
            continue
        if norm.rstrip("/") == page_url.rstrip("/"):
            continue
        result.add(norm)
    return result


def truncate_join(items, limit=5, sep="; "):
    items = [str(i) for i in items]
    if len(items) <= limit:
        return sep.join(items)
    return sep.join(items[:limit]) + f" ... (+{len(items) - limit} more)"


def classify_lang(url, en_prefix):
    path = urlparse(url).path
    return "en" if path.startswith(en_prefix) else "ja"


# --------------------------------------------------------------------------
# Sitemap discovery
# --------------------------------------------------------------------------

def fetch_xml(session, url, timeout):
    resp = session.get(url, timeout=timeout)
    resp.raise_for_status()
    return ET.fromstring(resp.content)


def collect_sitemap_urls(session, sitemap_url, timeout, exclude_patterns, seen=None):
    if seen is None:
        seen = set()
    if sitemap_url in seen:
        return []
    seen.add(sitemap_url)

    root = fetch_xml(session, sitemap_url, timeout)
    tag = root.tag.lower()
    results = []

    if tag.endswith("sitemapindex"):
        for sm in root.findall("sm:sitemap", SITEMAP_NS):
            loc = (sm.findtext("sm:loc", default="", namespaces=SITEMAP_NS) or "").strip()
            if not loc or any(p in loc.lower() for p in exclude_patterns):
                continue
            results.extend(collect_sitemap_urls(session, loc, timeout, exclude_patterns, seen))
    elif tag.endswith("urlset"):
        for u in root.findall("sm:url", SITEMAP_NS):
            loc = (u.findtext("sm:loc", default="", namespaces=SITEMAP_NS) or "").strip()
            lastmod = (u.findtext("sm:lastmod", default="", namespaces=SITEMAP_NS) or "").strip()
            if loc:
                results.append((loc, lastmod))

    return results


# --------------------------------------------------------------------------
# Page fetching / extraction
# --------------------------------------------------------------------------

def fetch_page(session, url, timeout, retries=3, backoff=2):
    last_exc = None
    for attempt in range(retries):
        try:
            return session.get(url, timeout=timeout, allow_redirects=True), None
        except requests.RequestException as e:
            last_exc = e
            time.sleep(backoff * (attempt + 1))
    return None, str(last_exc)


def check_link_status(session, url, timeout):
    try:
        resp = session.head(url, timeout=timeout, allow_redirects=True)
        if resp.status_code >= 400:
            resp = session.get(url, timeout=timeout, allow_redirects=True, stream=True)
            resp.close()
        return resp.status_code
    except requests.RequestException as e:
        return f"ERROR: {e.__class__.__name__}"


def find_main_container(soup):
    for selector in MAIN_CONTENT_SELECTORS:
        el = soup.select_one(selector)
        if el:
            return el
    return soup.body or soup


def extract_head_fields(soup):
    title = soup.title.get_text(strip=True) if soup.title else ""

    meta_tag = soup.find("meta", attrs={"name": re.compile("^description$", re.I)})
    meta_description = meta_tag["content"].strip() if meta_tag and meta_tag.get("content") else ""

    canonical_tag = soup.find("link", rel=lambda v: v and "canonical" in v)
    canonical = canonical_tag["href"].strip() if canonical_tag and canonical_tag.get("href") else ""

    hreflang = {}
    for link in soup.find_all("link", rel=lambda v: v and "alternate" in v):
        hl, href = link.get("hreflang"), link.get("href")
        if hl and href:
            hreflang[hl.lower()] = href.strip()

    h1_list = [h.get_text(strip=True) for h in soup.find_all("h1")]
    h2_list = [h.get_text(strip=True) for h in soup.find_all("h2")]
    return title, meta_description, canonical, hreflang, h1_list, h2_list


def extract_content_metrics(soup, lang):
    container = find_main_container(soup)
    # Re-parse a copy so we can strip boilerplate without mutating `soup`
    # (the caller still needs the untouched tree for full-page link scanning).
    working = BeautifulSoup(str(container), "lxml")
    for tag in working.select(BOILERPLATE_SELECTOR):
        tag.decompose()

    text = working.get_text(separator=" ", strip=True)
    images = working.find_all("img")
    missing_alt = [img.get("src", "") for img in images if not img.get("alt", "").strip()]

    word_count = char_count = None
    if lang == "en":
        word_count = len(text.split())
    else:
        char_count = len(re.sub(r"\s+", "", text))

    main_link_tags = working.find_all("a", href=True)
    return word_count, char_count, len(images), missing_alt, main_link_tags


def build_page(session, url, lang, sitemap_lastmod, timeout, base_netloc):
    page = Page(url=url, lang=lang, sitemap_lastmod=sitemap_lastmod)
    resp, err = fetch_page(session, url, timeout)
    if err or resp is None:
        page.fetch_error = err or "Unknown error"
        return page

    page.status_code = resp.status_code
    page.final_url = resp.url
    page.redirected = bool(resp.history) and resp.url.rstrip("/") != url.rstrip("/")
    page.http_last_modified = resp.headers.get("Last-Modified", "")

    if resp.status_code != 200 or "text/html" not in resp.headers.get("Content-Type", ""):
        return page

    soup = BeautifulSoup(resp.text, "lxml")

    page.all_internal_links = resolve_links(soup.find_all("a", href=True), url, base_netloc)
    (page.title, page.meta_description, page.canonical, page.hreflang,
     page.h1_list, page.h2_list) = extract_head_fields(soup)
    (page.word_count, page.char_count, page.image_total,
     page.missing_alt_examples, main_link_tags) = extract_content_metrics(soup, lang)
    page.image_missing_alt = len(page.missing_alt_examples)
    page.internal_links = resolve_links(main_link_tags, url, base_netloc)
    page.internal_link_count = len(page.internal_links)
    return page


# --------------------------------------------------------------------------
# Cross-page analysis
# --------------------------------------------------------------------------

def find_duplicates(pages, field_getter):
    groups = defaultdict(list)
    for p in pages:
        if not p.ok:
            continue
        val = field_getter(p)
        if not val:
            continue
        groups[(p.lang, val.strip().lower())].append(p.url)
    return {k: v for k, v in groups.items() if len(v) > 1}


def pair_pages(pages, en_prefix):
    by_url = {p.url.rstrip("/"): p for p in pages}
    en_pages = [p for p in pages if p.lang == "en" and p.ok]
    pairs, matched_en, matched_ja = [], set(), set()

    for en_page in en_pages:
        candidate = None
        for hl, href in en_page.hreflang.items():
            if hl.startswith("ja"):
                candidate = by_url.get(href.rstrip("/"))
                break
        if candidate is None:
            parsed = urlparse(en_page.url)
            path = parsed.path
            if path.startswith(en_prefix):
                ja_path = "/" + path[len(en_prefix):]
            elif path.rstrip("/") == en_prefix.rstrip("/"):
                ja_path = "/"
            else:
                ja_path = None
            if ja_path is not None:
                guess = f"{parsed.scheme}://{parsed.netloc}{ja_path}".rstrip("/")
                candidate = by_url.get(guess)
        if candidate:
            pairs.append((en_page, candidate))
            matched_en.add(en_page.url)
            matched_ja.add(candidate.url)

    unmatched_en = [p for p in en_pages if p.url not in matched_en]
    unmatched_ja = [p for p in pages if p.lang == "ja" and p.ok and p.url not in matched_ja]
    return pairs, unmatched_en, unmatched_ja


def check_broken_links(session, pages, timeout, delay, skip=False):
    crawled_status = {p.url.rstrip("/"): ("ERROR" if p.fetch_error else p.status_code) for p in pages}

    link_sources = defaultdict(set)
    for p in pages:
        for link in p.all_internal_links:
            link_sources[link].add(p.url)

    results = {}
    if skip:
        return results, link_sources

    for i, link in enumerate(sorted(link_sources), 1):
        key = link.rstrip("/")
        if key in crawled_status:
            results[link] = crawled_status[key]
        else:
            results[link] = check_link_status(session, link, timeout)
            time.sleep(delay)
    return results, link_sources


# --------------------------------------------------------------------------
# Issue generation
# --------------------------------------------------------------------------

def generate_issues(pages, pairs, unmatched_en, unmatched_ja, broken_links, link_sources,
                     dup_titles, dup_meta, thin_words, thin_chars):
    issues = []

    def add(sev, cat, url, lang, detail):
        issues.append(Issue(sev, cat, url, lang, detail))

    for p in pages:
        if p.fetch_error:
            add(HIGH, "Page Unreachable", p.url, p.lang, f"Request failed: {p.fetch_error}")
            continue
        if p.status_code != 200:
            add(HIGH, "Page Unreachable", p.url, p.lang, f"HTTP {p.status_code}")
            continue
        if p.redirected:
            add(LOW, "Sitemap Redirect", p.url, p.lang, f"Sitemap URL redirects to {p.final_url}")

        if not p.title:
            add(HIGH, "Missing Title", p.url, p.lang, "No <title> tag content")

        if not p.meta_description:
            add(MEDIUM, "Missing Meta Description", p.url, p.lang, "No meta description found")

        if not p.canonical:
            add(HIGH, "Missing Canonical", p.url, p.lang, "No canonical link tag found")
        elif p.canonical.rstrip("/") != p.url.rstrip("/") and p.canonical.rstrip("/") != p.final_url.rstrip("/"):
            add(MEDIUM, "Canonical Mismatch", p.url, p.lang, f"Canonical points to {p.canonical}, not this URL")

        if len(p.h1_list) == 0:
            add(MEDIUM, "Missing H1", p.url, p.lang, "No H1 tag found in main content")
        elif len(p.h1_list) > 1:
            add(MEDIUM, "Multiple H1", p.url, p.lang, f"{len(p.h1_list)} H1 tags found: {truncate_join(p.h1_list, 3)}")

        if p.image_missing_alt > 0:
            add(LOW, "Missing Alt Text", p.url, p.lang,
                f"{p.image_missing_alt}/{p.image_total} images missing alt text: "
                f"{truncate_join(p.missing_alt_examples, 3)}")

        content_len = p.word_count if p.lang == "en" else p.char_count
        threshold = thin_words if p.lang == "en" else thin_chars
        unit = "words" if p.lang == "en" else "characters"
        if content_len is not None and content_len < threshold:
            add(MEDIUM, "Thin Content", p.url, p.lang, f"{content_len} {unit} of main content (threshold {threshold})")

    for (lang, _text), urls in dup_titles.items():
        add(MEDIUM, "Duplicate Title", urls[0], lang, f"Same title used on {len(urls)} pages: {truncate_join(urls)}")

    for (lang, _text), urls in dup_meta.items():
        add(MEDIUM, "Duplicate Meta Description", urls[0], lang,
            f"Same meta description used on {len(urls)} pages: {truncate_join(urls)}")

    for link, status in broken_links.items():
        is_broken = isinstance(status, str) or status >= 400
        if is_broken:
            sources = link_sources.get(link, set())
            add(HIGH, "Broken Internal Link", link, "", f"Status: {status}. Linked from: {truncate_join(sources, 4)}")

    for en_page, ja_page in pairs:
        en_to_ja = en_page.hreflang.get("ja") or en_page.hreflang.get("ja-jp")
        ja_to_en = ja_page.hreflang.get("en") or ja_page.hreflang.get("en-us") or ja_page.hreflang.get("en-gb")
        en_ok = bool(en_to_ja) and en_to_ja.rstrip("/") == ja_page.url.rstrip("/")
        ja_ok = bool(ja_to_en) and ja_to_en.rstrip("/") == en_page.url.rstrip("/")
        if not (en_ok and ja_ok):
            add(HIGH, "Missing Hreflang Cross-Link", en_page.url, "en",
                f"hreflang pairing with JP counterpart {ja_page.url} is missing or incomplete "
                f"(EN->JA present: {en_ok}, JA->EN present: {ja_ok})")

        if bool(en_page.meta_description) != bool(ja_page.meta_description):
            add(MEDIUM, "EN/JP Inconsistency", en_page.url, "en",
                f"Meta description present on only one language version (JP: {ja_page.url})")

        if en_page.meta_description and ja_page.meta_description and \
                en_page.meta_description.strip() == ja_page.meta_description.strip():
            add(MEDIUM, "EN/JP Inconsistency", en_page.url, "en",
                f"Meta description is byte-identical between EN and JP (likely untranslated): {ja_page.url}")

        if en_page.title and ja_page.title and en_page.title.strip() == ja_page.title.strip():
            add(MEDIUM, "EN/JP Inconsistency", en_page.url, "en",
                f"Title is byte-identical between EN and JP (likely untranslated): {ja_page.url}")

        if len(en_page.h1_list) != len(ja_page.h1_list):
            add(LOW, "EN/JP Inconsistency", en_page.url, "en",
                f"H1 count differs: EN={len(en_page.h1_list)} vs JP={len(ja_page.h1_list)} ({ja_page.url})")

    for p in unmatched_en:
        add(INFO, "No Language Counterpart", p.url, p.lang, "No matching JP page found (by hreflang or URL pattern)")
    for p in unmatched_ja:
        add(INFO, "No Language Counterpart", p.url, p.lang, "No matching EN page found (by hreflang or URL pattern)")

    issues.sort(key=lambda i: (SEVERITY_ORDER.get(i.severity, 9), i.category, i.url))
    return issues


def find_recurring_missing_alt(pages, min_pages=2):
    """Images missing alt text on multiple pages usually means one shared
    template component (e.g. a CTA icon), not N separate mistakes."""
    src_to_pages = defaultdict(set)
    for p in pages:
        if not p.ok:
            continue
        for src in p.missing_alt_examples:
            src_to_pages[src].add(p.url)
    return {src: urls for src, urls in src_to_pages.items() if len(urls) >= min_pages}


# One-line recommended action + rough effort per issue category, used in the
# Executive Summary tab. Effort is a coarse hint, not an estimate in hours:
# Quick = a single-field fix, Medium = template/many-page fix, Involved = needs
# real content work, Strategic = a judgment call, not a "bug".
CATEGORY_GUIDANCE = {
    "Page Unreachable": ("Fix the broken URL or update/remove links and sitemap entries pointing to it.", "Quick"),
    "Missing Title": ("Add a unique, descriptive <title> tag.", "Quick"),
    "Duplicate Title": ("Differentiate the title tag on each affected page.", "Quick"),
    "Missing Meta Description": ("Write a unique meta description (roughly 50-160 characters).", "Quick"),
    "Duplicate Meta Description": ("Differentiate the meta description on each affected page.", "Quick"),
    "Missing Canonical": ("Add a self-referencing canonical tag.", "Quick"),
    "Canonical Mismatch": ("Confirm the canonical target is correct; fix if it points to the wrong URL.", "Quick"),
    "Missing H1": ("Add a single, descriptive H1 to the page's main content.", "Quick"),
    "Multiple H1": ("Keep one H1; demote the rest to H2/H3.", "Quick"),
    "Missing Alt Text": ("Add descriptive alt text — check the Executive Summary's recurring-image table first.", "Medium"),
    "Thin Content": ("Expand the page with more substantive, unique copy.", "Involved"),
    "Broken Internal Link": ("Fix or remove the link, or update it to the correct URL.", "Quick"),
    "Sitemap Redirect": ("Point the sitemap/internal links at the final URL directly.", "Quick"),
    "Missing Hreflang Cross-Link": ("Add reciprocal hreflang tags linking the EN and JP versions.", "Medium"),
    "EN/JP Inconsistency": ("Review and align content between the EN and JP versions of the page.", "Medium"),
    "No Language Counterpart": ("Confirm this is intentional (e.g. JP-only blog content); translate if not.", "Strategic"),
}


# --------------------------------------------------------------------------
# XLSX report
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SEVERITY_FILL = {
    HIGH: PatternFill("solid", fgColor="FFC7CE"),
    MEDIUM: PatternFill("solid", fgColor="FFEB9C"),
    LOW: PatternFill("solid", fgColor="DDEBF7"),
    INFO: PatternFill("solid", fgColor="F2F2F2"),
}
SEVERITY_FONT = {
    HIGH: Font(color="9C0006"),
    MEDIUM: Font(color="9C6500"),
}


def write_sheet(wb, title, headers, rows, col_widths=None):
    ws = wb.create_sheet(title=title[:31])
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws.freeze_panes = "A2"
    for row in rows:
        ws.append(row)
    widths = col_widths or [18] * len(headers)
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    return ws


def write_issues_sheet(wb, issues):
    headers = ["Severity", "Category", "URL", "Language", "Detail"]
    ws = write_sheet(wb, "Issues Summary", headers, [], col_widths=[10, 26, 45, 10, 90])
    for issue in issues:
        ws.append([issue.severity, issue.category, issue.url, issue.lang, issue.detail])
        row_idx = ws.max_row
        fill = SEVERITY_FILL.get(issue.severity)
        font = SEVERITY_FONT.get(issue.severity)
        if fill:
            ws.cell(row=row_idx, column=1).fill = fill
        if font:
            ws.cell(row=row_idx, column=1).font = font
    # Move to front
    wb.move_sheet("Issues Summary", offset=-len(wb.sheetnames))

    counts = defaultdict(int)
    for i in issues:
        counts[i.severity] += 1
    summary_row = ws.max_row + 2
    ws.cell(row=summary_row, column=1, value="Totals:").font = Font(bold=True)
    for i, sev in enumerate([HIGH, MEDIUM, LOW, INFO]):
        ws.cell(row=summary_row, column=2 + i, value=f"{sev}: {counts.get(sev, 0)}")


def write_executive_summary(wb, pages, issues, pairs, unmatched_en, unmatched_ja, recurring_alt):
    ws = wb.create_sheet("Executive Summary")
    ok_pages = [p for p in pages if p.ok]
    en_count = sum(1 for p in ok_pages if p.lang == "en")
    ja_count = sum(1 for p in ok_pages if p.lang == "ja")

    state = {"row": 1}

    def header(text, size=14):
        c = ws.cell(row=state["row"], column=1, value=text)
        c.font = Font(bold=True, size=size)
        state["row"] += 2

    def subheader(text):
        c = ws.cell(row=state["row"], column=1, value=text)
        c.font = Font(bold=True, size=11, color="305496")
        state["row"] += 1

    def line(text=""):
        ws.cell(row=state["row"], column=1, value=text)
        state["row"] += 1

    def table_header(cols):
        for idx, col_name in enumerate(cols, 1):
            cell = ws.cell(row=state["row"], column=idx, value=col_name)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        state["row"] += 1

    header("Passot SEO Audit — Executive Summary")
    line(f"Pages audited: {len(ok_pages)}  ({en_count} EN, {ja_count} JP)")
    line(f"EN/JP pairs matched: {len(pairs)}  |  Unmatched: {len(unmatched_en)} EN-only, {len(unmatched_ja)} JP-only")
    state["row"] += 1

    counts_by_sev = defaultdict(int)
    counts_by_cat = defaultdict(lambda: defaultdict(int))
    for i in issues:
        counts_by_sev[i.severity] += 1
        counts_by_cat[i.category][i.severity] += 1

    subheader("Issues by severity")
    for sev in [HIGH, MEDIUM, LOW, INFO]:
        line(f"{sev}: {counts_by_sev.get(sev, 0)}")
    state["row"] += 1

    subheader("Issues by category")
    table_header(["Category", "Count", "Worst Severity", "Recommended Action", "Effort"])
    cat_totals = sorted(counts_by_cat.items(), key=lambda kv: (SEVERITY_ORDER.get(min(kv[1], key=lambda s: SEVERITY_ORDER.get(s, 9)), 9), -sum(kv[1].values())))
    for cat, sev_counts in cat_totals:
        total = sum(sev_counts.values())
        worst = min(sev_counts.keys(), key=lambda s: SEVERITY_ORDER.get(s, 9))
        action, effort = CATEGORY_GUIDANCE.get(cat, ("Review the flagged pages.", "Medium"))
        ws.cell(row=state["row"], column=1, value=cat)
        ws.cell(row=state["row"], column=2, value=total)
        ws.cell(row=state["row"], column=3, value=worst)
        ws.cell(row=state["row"], column=4, value=action)
        ws.cell(row=state["row"], column=5, value=effort)
        state["row"] += 1
    state["row"] += 1

    if recurring_alt:
        subheader("Recurring root causes (fix once, resolves many pages)")
        table_header(["Shared Image", "Pages Affected", "Example Pages"])
        for src, urls in sorted(recurring_alt.items(), key=lambda kv: -len(kv[1]))[:20]:
            ws.cell(row=state["row"], column=1, value=src)
            ws.cell(row=state["row"], column=2, value=len(urls))
            ws.cell(row=state["row"], column=3, value=truncate_join(sorted(urls), 3))
            state["row"] += 1
        state["row"] += 1

    subheader("Fix-first priority list")
    high_cats = sorted(c for c, s in counts_by_cat.items() if HIGH in s)
    med_cats = sorted(c for c, s in counts_by_cat.items() if MEDIUM in s and HIGH not in s)
    priorities = []
    if high_cats:
        priorities.append("Address High severity issues first: " + ", ".join(high_cats) + ".")
    if recurring_alt:
        instances = sum(len(urls) for urls in recurring_alt.values())
        priorities.append(
            f"Fix the {len(recurring_alt)} shared/reused image(s) missing alt text once — "
            f"together they account for {instances} page-level flags."
        )
    if med_cats:
        priorities.append("Schedule Medium severity issues next: " + ", ".join(med_cats) + ".")
    if not priorities:
        priorities.append("No High or Medium severity issues found. Review Low/Info items opportunistically.")
    for idx, text in enumerate(priorities, 1):
        line(f"{idx}. {text}")
    line()
    line("Full page-by-page detail behind every number above is in the Issues Summary tab.")

    for idx, width in enumerate([45, 12, 16, 60, 10], 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    wb.move_sheet("Executive Summary", offset=-len(wb.sheetnames))
    return ws


def build_workbook(pages, issues, pairs, unmatched_en, unmatched_ja, broken_links, link_sources, thin_words, thin_chars):
    wb = Workbook()
    wb.remove(wb.active)

    write_issues_sheet(wb, issues)
    recurring_alt = find_recurring_missing_alt(pages)
    write_executive_summary(wb, pages, issues, pairs, unmatched_en, unmatched_ja, recurring_alt)

    # All Pages
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        rows.append([
            p.url, p.lang, p.status_code or p.fetch_error, p.title, len(p.title),
            p.meta_description, len(p.meta_description), p.canonical,
            len(p.h1_list), len(p.h2_list), p.image_total, p.image_missing_alt,
            p.internal_link_count, p.word_count if p.lang == "en" else "",
            p.char_count if p.lang == "ja" else "", p.http_last_modified, p.sitemap_lastmod,
            len(p.hreflang),
        ])
    write_sheet(wb, "All Pages", [
        "URL", "Language", "Status", "Title", "Title Length", "Meta Description",
        "Meta Length", "Canonical", "H1 Count", "H2 Count", "Image Count",
        "Images Missing Alt", "Internal Link Count", "Word Count (EN)",
        "Char Count (JP)", "Last-Modified (HTTP)", "Last-Modified (Sitemap)", "Hreflang Tag Count",
    ], rows, col_widths=[45, 6, 8, 35, 10, 40, 10, 40, 8, 8, 8, 10, 10, 10, 10, 20, 14, 10])

    # Meta Descriptions
    dup_meta = find_duplicates(pages, lambda p: p.meta_description)
    dup_meta_urls = {u for urls in dup_meta.values() for u in urls}
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        rows.append([p.url, p.lang, p.meta_description, len(p.meta_description),
                     "Yes" if not p.meta_description else "No",
                     "Yes" if p.url in dup_meta_urls else "No"])
    write_sheet(wb, "Meta Descriptions", ["URL", "Language", "Meta Description", "Length", "Missing?", "Duplicate?"],
                rows, col_widths=[45, 6, 60, 8, 8, 10])

    # Titles & Headings
    dup_titles = find_duplicates(pages, lambda p: p.title)
    dup_title_urls = {u for urls in dup_titles.values() for u in urls}
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        rows.append([p.url, p.lang, p.title, len(p.title),
                     "Yes" if p.url in dup_title_urls else "No",
                     len(p.h1_list), truncate_join(p.h1_list, 3),
                     len(p.h2_list), truncate_join(p.h2_list, 5)])
    write_sheet(wb, "Titles & Headings",
                ["URL", "Language", "Title", "Title Length", "Duplicate Title?", "H1 Count", "H1 Text", "H2 Count", "H2 Text"],
                rows, col_widths=[45, 6, 35, 10, 12, 8, 40, 8, 50])

    # Images & Alt Text
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        pct = round(100 * p.image_missing_alt / p.image_total, 1) if p.image_total else 0
        rows.append([p.url, p.lang, p.image_total, p.image_missing_alt, pct,
                     truncate_join(p.missing_alt_examples, 4)])
    write_sheet(wb, "Images & Alt Text",
                ["URL", "Language", "Total Images", "Missing Alt", "% Missing", "Example Sources"],
                rows, col_widths=[45, 6, 10, 10, 10, 60])

    # Internal Links (per page)
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        rows.append([p.url, p.lang, p.internal_link_count])
    write_sheet(wb, "Internal Links", ["URL", "Language", "Internal Link Count (main content)"],
                rows, col_widths=[45, 6, 30])

    # Broken Links
    rows = []
    for link, status in sorted(broken_links.items()):
        is_broken = isinstance(status, str) or status >= 400
        if is_broken:
            rows.append([link, status, truncate_join(link_sources.get(link, set()), 6)])
    write_sheet(wb, "Broken Links", ["Broken URL", "Status", "Linked From"], rows, col_widths=[50, 12, 80])

    # Canonical Tags
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        self_ref = p.canonical and (p.canonical.rstrip("/") == p.url.rstrip("/") or p.canonical.rstrip("/") == p.final_url.rstrip("/"))
        rows.append([p.url, p.lang, p.canonical or "(missing)", "Yes" if self_ref else "No"])
    write_sheet(wb, "Canonical Tags", ["URL", "Language", "Canonical Value", "Self-Referencing?"],
                rows, col_widths=[45, 6, 45, 15])

    # EN-JP Comparison
    rows = []
    for en_page, ja_page in sorted(pairs, key=lambda pr: pr[0].url):
        en_to_ja = en_page.hreflang.get("ja") or en_page.hreflang.get("ja-jp")
        ja_to_en = ja_page.hreflang.get("en") or ja_page.hreflang.get("en-us") or ja_page.hreflang.get("en-gb")
        reciprocal = bool(en_to_ja) and bool(ja_to_en) and \
            en_to_ja.rstrip("/") == ja_page.url.rstrip("/") and ja_to_en.rstrip("/") == en_page.url.rstrip("/")
        rows.append([
            en_page.url, ja_page.url, "Yes" if reciprocal else "No",
            en_page.title, ja_page.title,
            "Yes" if en_page.meta_description else "No", "Yes" if ja_page.meta_description else "No",
            "Yes" if (en_page.meta_description and ja_page.meta_description and
                      en_page.meta_description.strip() == ja_page.meta_description.strip()) else "No",
            len(en_page.h1_list), len(ja_page.h1_list),
            en_page.canonical or "(missing)", ja_page.canonical or "(missing)",
        ])
    write_sheet(wb, "EN-JP Comparison", [
        "EN URL", "JP URL", "Hreflang Reciprocal?", "EN Title", "JP Title",
        "EN Meta Present?", "JP Meta Present?", "Meta Identical (Untranslated?)",
        "EN H1 Count", "JP H1 Count", "EN Canonical", "JP Canonical",
    ], rows, col_widths=[40, 40, 12, 30, 30, 10, 10, 16, 8, 8, 40, 40])

    unmatched_rows = [[p.url, p.lang] for p in unmatched_en + unmatched_ja]
    if unmatched_rows:
        ws = wb["EN-JP Comparison"]
        start = ws.max_row + 3
        ws.cell(row=start, column=1, value="Pages with no counterpart:").font = Font(bold=True)
        ws.cell(row=start + 1, column=1, value="URL").font = HEADER_FONT
        ws.cell(row=start + 1, column=2, value="Language").font = HEADER_FONT
        for i, (url, lang) in enumerate(unmatched_rows, start + 2):
            ws.cell(row=i, column=1, value=url)
            ws.cell(row=i, column=2, value=lang)

    return wb


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def parse_args(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL)
    ap.add_argument("--sitemap-url", default=DEFAULT_SITEMAP_URL)
    ap.add_argument("--en-prefix", default=DEFAULT_EN_PREFIX)
    ap.add_argument("--output", default=DEFAULT_OUTPUT)
    ap.add_argument("--delay", type=float, default=DEFAULT_DELAY, help="Seconds between requests")
    ap.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT)
    ap.add_argument("--thin-words", type=int, default=DEFAULT_THIN_WORDS_EN, help="Thin-content threshold for EN pages (words)")
    ap.add_argument("--thin-chars", type=int, default=DEFAULT_THIN_CHARS_JA, help="Thin-content threshold for JP pages (characters)")
    ap.add_argument("--user-agent", default=DEFAULT_USER_AGENT)
    ap.add_argument("--max-pages", type=int, default=None, help="Limit number of pages crawled (for a quick test run)")
    ap.add_argument("--skip-broken-links", action="store_true", help="Skip checking internal links for broken status")
    ap.add_argument("--exclude-sitemap", nargs="*", default=DEFAULT_SITEMAP_EXCLUDE,
                     help="Skip sub-sitemaps whose filename contains any of these substrings")
    return ap.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    base_netloc = normalize_netloc(urlparse(args.base_url).netloc)

    session = requests.Session()
    session.headers.update({
        "User-Agent": args.user_agent,
        "Accept-Language": "en,ja;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })

    print(f"Discovering pages from {args.sitemap_url} ...")
    try:
        sitemap_entries = collect_sitemap_urls(session, args.sitemap_url, args.timeout, args.exclude_sitemap)
    except Exception as e:
        print(f"ERROR: could not read sitemap: {e}", file=sys.stderr)
        return 1

    seen_urls = set()
    to_crawl = []
    for loc, lastmod in sitemap_entries:
        if normalize_netloc(urlparse(loc).netloc) != base_netloc:
            continue
        if not is_probably_html(loc):
            continue
        key = loc.rstrip("/")
        if key in seen_urls:
            continue
        seen_urls.add(key)
        to_crawl.append((loc, lastmod))

    if args.max_pages:
        to_crawl = to_crawl[:args.max_pages]

    print(f"Found {len(to_crawl)} pages to crawl.")
    if not to_crawl:
        print("No pages found — check --sitemap-url.", file=sys.stderr)
        return 1

    pages = []
    for i, (url, lastmod) in enumerate(to_crawl, 1):
        lang = classify_lang(url, args.en_prefix)
        print(f"[{i}/{len(to_crawl)}] ({lang}) {url}")
        page = build_page(session, url, lang, lastmod, args.timeout, base_netloc)
        pages.append(page)
        time.sleep(args.delay)

    print("Checking for duplicate titles / meta descriptions ...")
    dup_titles = find_duplicates(pages, lambda p: p.title)
    dup_meta = find_duplicates(pages, lambda p: p.meta_description)

    print("Pairing EN/JP page counterparts ...")
    pairs, unmatched_en, unmatched_ja = pair_pages(pages, args.en_prefix)
    print(f"  {len(pairs)} pairs matched, {len(unmatched_en)} EN pages and {len(unmatched_ja)} JP pages unmatched.")

    if args.skip_broken_links:
        print("Skipping broken-link check (--skip-broken-links).")
    else:
        print("Checking internal links for broken status (this can take a while) ...")
    broken_links, link_sources = check_broken_links(session, pages, args.timeout, args.delay, skip=args.skip_broken_links)

    print("Generating issue list ...")
    issues = generate_issues(pages, pairs, unmatched_en, unmatched_ja, broken_links, link_sources,
                              dup_titles, dup_meta, args.thin_words, args.thin_chars)

    print("Writing workbook ...")
    wb = build_workbook(pages, issues, pairs, unmatched_en, unmatched_ja, broken_links, link_sources,
                         args.thin_words, args.thin_chars)
    wb.save(args.output)

    counts = defaultdict(int)
    for i in issues:
        counts[i.severity] += 1
    print(f"\nDone. Wrote {args.output}")
    print(f"Issues found — High: {counts[HIGH]}, Medium: {counts[MEDIUM]}, Low: {counts[LOW]}, Info: {counts[INFO]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
