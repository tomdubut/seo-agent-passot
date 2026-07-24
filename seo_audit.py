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
import csv
import json
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Literal
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel

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
# Cap on stored main-content text per page — bounds memory and, when
# --ai-analysis is used, bounds per-page token cost. Generous relative to
# observed page sizes (even long-form articles run well under this).
MAX_CONTENT_TEXT_CHARS = 8000
DEFAULT_AI_MODEL = "claude-sonnet-5"
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
    structured_data_types: list = field(default_factory=list)
    og_tags: dict = field(default_factory=dict)
    twitter_card: bool = False
    content_text: str = ""

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


def find_main_container_labeled(soup):
    """Among all candidate containers found on the page, pick the one with
    the most text — not just the first matching selector. Themes/page
    builders (Elementor especially) sometimes render a near-empty <main>
    wrapper alongside the real content in a sibling container (e.g.
    `.elementor`); picking the first match by priority alone can grab the
    empty wrapper and report the page as having ~0 words."""
    candidates = [(selector, el) for selector in MAIN_CONTENT_SELECTORS for el in soup.select(selector)]
    if not candidates:
        return "body-fallback (no selector matched)", (soup.body or soup)
    return max(candidates, key=lambda pair: len(pair[1].get_text(strip=True)))


def find_main_container(soup):
    _, el = find_main_container_labeled(soup)
    return el


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


def extract_structured_data(soup):
    """Presence-only check: JSON-LD @type values, Open Graph tags, Twitter Card.
    Does not validate schema correctness, just whether it's there at all."""
    ld_types = []
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
        except (ValueError, TypeError):
            continue
        for item in data if isinstance(data, list) else [data]:
            if not isinstance(item, dict):
                continue
            for node in item.get("@graph", [item]):
                if not isinstance(node, dict):
                    continue
                t = node.get("@type")
                if isinstance(t, list):
                    ld_types.extend(str(x) for x in t)
                elif t:
                    ld_types.append(str(t))

    og_tags = {}
    for meta in soup.find_all("meta", property=re.compile("^og:", re.I)):
        content = (meta.get("content") or "").strip()
        if content:
            og_tags[meta["property"].lower()] = content

    twitter_card = bool(soup.find("meta", attrs={"name": re.compile("^twitter:card$", re.I)}))

    return sorted(set(ld_types)), og_tags, twitter_card


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
    return word_count, char_count, len(images), missing_alt, main_link_tags, text[:MAX_CONTENT_TEXT_CHARS]


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
    (page.structured_data_types, page.og_tags, page.twitter_card) = extract_structured_data(soup)
    (page.word_count, page.char_count, page.image_total,
     page.missing_alt_examples, main_link_tags, page.content_text) = extract_content_metrics(soup, lang)
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


# Thin-content thresholds are not a real Google ranking signal (there is no
# official minimum word count) — they're a heuristic proxy for "might not be
# substantive," so one number for an entire site is a blunt instrument. A
# product spec page and a competitive blog article don't need the same bar.
# EXEMPT means "this page type is legitimately short by design, don't flag it
# at all" (e.g. a homepage or a blog listing page).
EXEMPT = "exempt"

# (label, word threshold for EN, char threshold for JP) keyed by first URL
# path segment (after stripping the EN prefix) or, for Utility, by any
# hyphen-separated token in the path. Tune these to match your own site's
# URL structure and content strategy.
PAGE_TYPE_BY_FIRST_SEGMENT = {
    "products": ("Product", 150, 300),
    "retail-displays": ("Product", 150, 300),
    "company": ("Company/About", 150, 300),
    "about": ("Company/About", 150, 300),
    "our-process": ("Company/About", 150, 300),
    "process-retail-display": ("Company/About", 150, 300),
}
UTILITY_TOKENS = {"contact", "faq", "privacy", "terms", "security", "policy"}
BLOG_INDEX_SEGMENTS = {"column", "blog"}


def classify_page_type_by_path(url, en_prefix):
    path = urlparse(url).path
    if en_prefix and path.startswith(en_prefix):
        path = "/" + path[len(en_prefix):]
    path = path.strip("/")
    if not path:
        return "Homepage", EXEMPT, EXEMPT

    segments = path.split("/")
    if len(segments) == 1 and segments[0] in BLOG_INDEX_SEGMENTS:
        return "Blog Index", EXEMPT, EXEMPT
    if segments[0] in PAGE_TYPE_BY_FIRST_SEGMENT:
        return PAGE_TYPE_BY_FIRST_SEGMENT[segments[0]]

    tokens = set(re.split(r"[-_]", path.lower()))
    if tokens & UTILITY_TOKENS:
        return "Utility", 80, 150
    if segments[0].startswith("column") or "column" in tokens:
        return "Blog/Column Article", 600, 1000

    return None


def classify_page_types(pages, pairs, en_prefix, default_words, default_chars):
    """URL-path classification first. EN pages with no keyword match (e.g.
    SEO-friendly slugs that share no vocabulary with their JP counterpart's
    URL) inherit their JP pair's classification instead of guessing — that's
    real matched data, not a fabricated signal. Anything still unclassified
    falls back to the site-wide --thin-words/--thin-chars default."""
    by_url = {}
    for p in pages:
        by_url[p.url] = classify_page_type_by_path(p.url, en_prefix)

    for en_page, ja_page in pairs:
        if by_url.get(en_page.url) is None and by_url.get(ja_page.url) is not None:
            by_url[en_page.url] = by_url[ja_page.url]

    result = {}
    for p in pages:
        classified = by_url.get(p.url)
        if classified is None:
            result[p.url] = ("Other", default_words, default_chars)
        else:
            result[p.url] = classified
    return result


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


def compute_inbound_link_counts(pages):
    """How many other crawled pages link to each page from their main content
    (nav/footer links excluded, since those are identical on every page and
    would mask genuinely under-linked content)."""
    urls_by_key = {p.url.rstrip("/"): p.url for p in pages}
    inbound = defaultdict(set)
    for p in pages:
        if not p.ok:
            continue
        for link in p.internal_links:
            key = link.rstrip("/")
            target = urls_by_key.get(key)
            if target and target != p.url:
                inbound[target].add(p.url)
    return inbound


# --------------------------------------------------------------------------
# Issue generation
# --------------------------------------------------------------------------

# Homepages are expected to have few/no inbound links from other pages'
# main content (they're reached via nav, not content links) — exclude them
# from the "weakly linked" check to avoid noise.
WEAK_INBOUND_THRESHOLD = 1

def generate_issues(pages, pairs, unmatched_en, unmatched_ja, broken_links, link_sources,
                     dup_titles, dup_meta, thin_words, thin_chars, en_prefix=DEFAULT_EN_PREFIX):
    issues = []
    inbound_counts = compute_inbound_link_counts(pages)
    page_types = classify_page_types(pages, pairs, en_prefix, thin_words, thin_chars)
    homepage_paths = {"/", en_prefix.rstrip("/") or "/en"}

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

        page_type, word_th, char_th = page_types.get(p.url, ("Other", thin_words, thin_chars))
        threshold = word_th if p.lang == "en" else char_th
        content_len = p.word_count if p.lang == "en" else p.char_count
        unit = "words" if p.lang == "en" else "characters"
        if threshold != EXEMPT and content_len is not None and content_len < threshold:
            add(MEDIUM, "Thin Content", p.url, p.lang,
                f"{content_len} {unit} of main content (threshold {threshold} for page type '{page_type}')")

        if not p.structured_data_types:
            add(LOW, "Missing Structured Data", p.url, p.lang, "No JSON-LD structured data (schema.org) found on this page")

        missing_og = [k for k in ("og:title", "og:description", "og:image") if k not in p.og_tags]
        if missing_og:
            add(LOW, "Missing Open Graph Tags", p.url, p.lang,
                f"Missing: {', '.join(missing_og)} (affects link preview appearance on social/chat apps)")

        path = urlparse(p.url).path.rstrip("/") or "/"
        if path not in homepage_paths:
            inbound = len(inbound_counts.get(p.url, set()))
            if inbound <= WEAK_INBOUND_THRESHOLD:
                add(LOW, "Weakly Linked Internally", p.url, p.lang,
                    f"Only {inbound} other page(s) link to this page from their main content — "
                    "consider adding internal links from related pages")

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
    "Missing Structured Data": ("Add JSON-LD schema.org markup (Article, FAQPage, Organization, etc. as relevant).", "Involved"),
    "Missing Open Graph Tags": ("Add og:title, og:description and og:image meta tags for social/chat link previews.", "Quick"),
    "Weakly Linked Internally": ("Add internal links to this page from related content elsewhere on the site.", "Quick"),
}


# --------------------------------------------------------------------------
# AI content & keyword analysis (opt-in via --ai-analysis)
# --------------------------------------------------------------------------

RATING_VALUES = Literal["Good", "Needs Improvement", "Poor"]


class PageAIAnalysis(BaseModel):
    keyword_or_topic: str
    keyword_is_inferred: bool
    title_assessment: str
    title_rating: RATING_VALUES
    meta_description_assessment: str
    meta_description_rating: RATING_VALUES
    content_assessment: str
    content_rating: RATING_VALUES
    overall_rating: RATING_VALUES
    suggestions: List[str]


AI_ANALYSIS_SYSTEM_PROMPT = (
    "You are an experienced SEO analyst reviewing a single web page's on-page optimization. "
    "Give concise, specific, and actionable judgments grounded in what is actually on the page — "
    "not generic SEO advice a template could produce. Judge keyword/topic fit, whether the title "
    "and meta description are compelling and relevant (not just present), and whether the content "
    "substantively covers the topic a searcher would expect, in the page's own language. Be honest: "
    "most pages have room to improve, so reserve 'Good' for genuinely strong pages, and make every "
    "suggestion concrete enough that someone could act on it without asking a follow-up question."
)


def load_keyword_map(path):
    """CSV with 'url' and 'keyword' columns, mapping specific pages to a
    site-owner-provided target keyword/topic instead of letting the AI infer
    one from the page's own content."""
    keyword_map = {}
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            url = (row.get("url") or row.get("URL") or "").strip()
            keyword = (row.get("keyword") or row.get("Keyword") or "").strip()
            if url and keyword:
                keyword_map[url.rstrip("/")] = keyword
    return keyword_map


def analyze_page_with_ai(client, page, model, target_keyword=None):
    if target_keyword:
        keyword_line = (
            f"The site owner has specified this page's target keyword/topic as: {target_keyword!r}. "
            "Judge fit against this specific target."
        )
    else:
        keyword_line = (
            "No target keyword was specified — infer the page's likely intended topic/keyword from "
            "its own title, headings, and content, and say so."
        )

    user_content = (
        f"URL: {page.url}\n"
        f"Language: {'English' if page.lang == 'en' else 'Japanese'}\n"
        f"{keyword_line}\n\n"
        f"Title tag ({len(page.title)} chars): {page.title!r}\n"
        f"Meta description ({len(page.meta_description)} chars): {page.meta_description!r}\n"
        f"H1: {page.h1_list}\n"
        f"H2s: {page.h2_list}\n\n"
        f"Main page content:\n{page.content_text}"
    )

    response = client.messages.parse(
        model=model,
        max_tokens=1024,
        system=AI_ANALYSIS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
        output_format=PageAIAnalysis,
    )
    return response.parsed_output


def run_ai_analysis(pages, model, keyword_map):
    import anthropic

    client = anthropic.Anthropic()
    ok_pages = [p for p in pages if p.ok]
    results = {}
    print(f"Running AI content/keyword analysis on {len(ok_pages)} pages using {model} ...")
    for i, p in enumerate(ok_pages, 1):
        target_keyword = keyword_map.get(p.url.rstrip("/"))
        print(f"  [{i}/{len(ok_pages)}] {p.url}")
        try:
            results[p.url] = analyze_page_with_ai(client, p, model, target_keyword)
        except Exception as e:
            print(f"    AI analysis failed: {e}", file=sys.stderr)
    return results


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


def write_executive_summary(wb, pages, issues, pairs, unmatched_en, unmatched_ja, recurring_alt, ai_results=None):
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

    content_gaps = unmatched_en + unmatched_ja
    if content_gaps:
        subheader("Translation / content parity gaps (not in the Issues tab's severity ranking)")
        line(f"{len(unmatched_ja)} JP page(s) have no English version. {len(unmatched_en)} EN page(s) have no Japanese version.")
        line("Not necessarily wrong (e.g. intentionally JP-only blog posts) — but a deliberate content decision, not a bug to fix.")
        table_header(["URL", "Language", "Content Length"])
        for p in sorted(content_gaps, key=lambda x: (x.lang, x.url)):
            if p.lang == "en" and p.word_count is not None:
                length = f"{p.word_count} words"
            elif p.char_count is not None:
                length = f"{p.char_count} characters"
            else:
                length = ""
            ws.cell(row=state["row"], column=1, value=p.url)
            ws.cell(row=state["row"], column=2, value=p.lang)
            ws.cell(row=state["row"], column=3, value=length)
            state["row"] += 1
        state["row"] += 1

    if ai_results:
        subheader("AI content & keyword review")
        rating_counts = defaultdict(int)
        poor_pages = []
        for url, r in ai_results.items():
            rating_counts[r.overall_rating] += 1
            if r.overall_rating == "Poor":
                poor_pages.append(url)
        line(
            f"{len(ai_results)} page(s) reviewed. Good: {rating_counts.get('Good', 0)}, "
            f"Needs Improvement: {rating_counts.get('Needs Improvement', 0)}, Poor: {rating_counts.get('Poor', 0)}."
        )
        if poor_pages:
            line("Pages rated 'Poor' overall: " + truncate_join(sorted(poor_pages), 8))
        line("Full per-page assessments and suggestions are in the AI Content & Keyword Analysis tab.")
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
    if content_gaps:
        priorities.append(
            f"Decide on the {len(content_gaps)} page(s) with no counterpart in the other language — "
            "translate for parity, or confirm it's intentional (see the table above)."
        )
    if ai_results:
        poor_count = sum(1 for r in ai_results.values() if r.overall_rating == "Poor")
        if poor_count:
            priorities.append(f"Review the {poor_count} page(s) the AI rated 'Poor' overall first.")
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


def build_workbook(pages, issues, pairs, unmatched_en, unmatched_ja, broken_links, link_sources,
                    thin_words, thin_chars, en_prefix=DEFAULT_EN_PREFIX, ai_results=None):
    wb = Workbook()
    wb.remove(wb.active)

    write_issues_sheet(wb, issues)
    recurring_alt = find_recurring_missing_alt(pages)
    write_executive_summary(wb, pages, issues, pairs, unmatched_en, unmatched_ja, recurring_alt, ai_results)

    # All Pages
    page_types = classify_page_types(pages, pairs, en_prefix, thin_words, thin_chars)
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        page_type, word_th, char_th = page_types.get(p.url, ("Other", thin_words, thin_chars))
        threshold = word_th if p.lang == "en" else char_th
        rows.append([
            p.url, p.lang, p.status_code or p.fetch_error, p.title, len(p.title),
            p.meta_description, len(p.meta_description), p.canonical,
            len(p.h1_list), len(p.h2_list), p.image_total, p.image_missing_alt,
            p.internal_link_count, p.word_count if p.lang == "en" else "",
            p.char_count if p.lang == "ja" else "", p.http_last_modified, p.sitemap_lastmod,
            len(p.hreflang), page_type, "exempt" if threshold == EXEMPT else threshold,
        ])
    write_sheet(wb, "All Pages", [
        "URL", "Language", "Status", "Title", "Title Length", "Meta Description",
        "Meta Length", "Canonical", "H1 Count", "H2 Count", "Image Count",
        "Images Missing Alt", "Internal Link Count", "Word Count (EN)",
        "Char Count (JP)", "Last-Modified (HTTP)", "Last-Modified (Sitemap)", "Hreflang Tag Count",
        "Page Type", "Thin-Content Threshold Used",
    ], rows, col_widths=[45, 6, 8, 35, 10, 40, 10, 40, 8, 8, 8, 10, 10, 10, 10, 20, 14, 10, 16, 14])

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
    inbound_counts = compute_inbound_link_counts(pages)
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        rows.append([p.url, p.lang, p.internal_link_count, len(inbound_counts.get(p.url, set()))])
    write_sheet(wb, "Internal Links",
                ["URL", "Language", "Outbound Links (main content)", "Inbound Links (from other pages' main content)"],
                rows, col_widths=[45, 6, 22, 30])

    # Structured Data & Social Tags
    rows = []
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        if not p.ok:
            continue
        rows.append([
            p.url, p.lang, truncate_join(p.structured_data_types, 5) or "(none)",
            "Yes" if p.og_tags else "No",
            "Yes" if "og:title" in p.og_tags else "No",
            "Yes" if "og:description" in p.og_tags else "No",
            "Yes" if "og:image" in p.og_tags else "No",
            "Yes" if p.twitter_card else "No",
        ])
    write_sheet(wb, "Structured Data & Social", [
        "URL", "Language", "JSON-LD Types Found", "Has Open Graph Tags?",
        "OG Title?", "OG Description?", "OG Image?", "Twitter Card?",
    ], rows, col_widths=[45, 6, 35, 16, 10, 14, 10, 12])

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

    if ai_results:
        write_ai_analysis_sheet(wb, pages, ai_results)

    return wb


RATING_FILL = {
    "Poor": PatternFill("solid", fgColor="FFC7CE"),
    "Needs Improvement": PatternFill("solid", fgColor="FFEB9C"),
    "Good": PatternFill("solid", fgColor="C6EFCE"),
}


def write_ai_analysis_sheet(wb, pages, ai_results):
    headers = [
        "URL", "Language", "Keyword/Topic", "Keyword Inferred?",
        "Title Rating", "Title Assessment", "Meta Rating", "Meta Assessment",
        "Content Rating", "Content Assessment", "Overall Rating", "Suggestions",
    ]
    ws = write_sheet(wb, "AI Content & Keyword Analysis", headers, [],
                      col_widths=[45, 6, 25, 10, 14, 45, 14, 45, 14, 45, 14, 60])

    rating_cols = {5: "title_rating", 7: "meta_description_rating", 9: "content_rating", 11: "overall_rating"}
    for p in sorted(pages, key=lambda x: (x.lang, x.url)):
        r = ai_results.get(p.url)
        if not r:
            continue
        ws.append([
            p.url, p.lang, r.keyword_or_topic, "Yes" if r.keyword_is_inferred else "No",
            r.title_rating, r.title_assessment,
            r.meta_description_rating, r.meta_description_assessment,
            r.content_rating, r.content_assessment,
            r.overall_rating, "; ".join(r.suggestions),
        ])
        row_idx = ws.max_row
        for col_idx, attr in rating_cols.items():
            fill = RATING_FILL.get(getattr(r, attr))
            if fill:
                ws.cell(row=row_idx, column=col_idx).fill = fill


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
    ap.add_argument("--debug-url", default=None,
                     help="Fetch a single URL and print diagnostics about what content was extracted "
                          "and why (use this to troubleshoot an unexpected word/char count) instead "
                          "of running a full audit")
    ap.add_argument("--ai-analysis", action="store_true",
                     help="Use the Claude API to judge keyword fit, meta quality, and content quality "
                          "per page (costs a small amount per run — see README). Requires "
                          "ANTHROPIC_API_KEY in the environment.")
    ap.add_argument("--ai-model", default=DEFAULT_AI_MODEL,
                     help=f"Claude model to use for --ai-analysis (default: {DEFAULT_AI_MODEL})")
    ap.add_argument("--ai-keyword-map", default=None,
                     help="Optional CSV file with 'url' and 'keyword' columns giving explicit target "
                          "keywords for specific pages; pages not listed fall back to AI-inferred topic")
    return ap.parse_args(argv)


def debug_single_url(session, url, en_prefix, timeout):
    lang = classify_lang(url, en_prefix)
    print(f"Fetching {url}  (classified as: {lang})\n")
    resp, err = fetch_page(session, url, timeout)
    if err or resp is None:
        print(f"FETCH FAILED: {err}")
        return 1

    print(f"HTTP status: {resp.status_code}")
    print(f"Content-Type: {resp.headers.get('Content-Type')}")
    if resp.url != url:
        print(f"Redirected to: {resp.url}")
    if resp.status_code != 200 or "text/html" not in resp.headers.get("Content-Type", ""):
        print("\nNot a 200 HTML response, so no content was extracted from it during a real audit run.")
        return 1

    soup = BeautifulSoup(resp.text, "lxml")
    print(f"Raw HTML size: {len(resp.text)} characters\n")

    print("Main-content selectors checked (in priority order), and what each finds on this page:")
    for selector in MAIN_CONTENT_SELECTORS:
        matches = soup.select(selector)
        if matches:
            lengths = [len(m.get_text(strip=True)) for m in matches]
            print(f"  {selector!r:12} -> {len(matches)} match(es), text length(s): {lengths}")
        else:
            print(f"  {selector!r:12} -> no match")

    chosen_selector, container = find_main_container_labeled(soup)
    print(f"\n=> This audit run would use: {chosen_selector!r}")

    word_count, char_count, image_total, missing_alt, _ = extract_content_metrics(soup, lang)
    print(f"\nAfter stripping {BOILERPLATE_SELECTOR!r} from that container:")
    print(f"  Word count (EN): {word_count}")
    print(f"  Char count (JP): {char_count}")
    print(f"  Images found: {image_total} (missing alt: {len(missing_alt)})")

    working = BeautifulSoup(str(container), "lxml")
    for tag in working.select(BOILERPLATE_SELECTOR):
        tag.decompose()
    preview = working.get_text(separator=" ", strip=True)[:400]
    print(f"\nExtracted text preview (first 400 chars): {preview!r}")

    full_body_len = len(soup.get_text(separator=" ", strip=True))
    extracted_len = (word_count or 0) + (char_count or 0)
    print(f"\nFor comparison, the ENTIRE page's text (not just the chosen container): {full_body_len} characters")
    if full_body_len > 300 and extracted_len < 20:
        print(
            "\n>>> The page clearly has real text somewhere, but almost none of it came from the chosen\n"
            "    container. That points to a container-selection bug (wrong element picked) rather than\n"
            "    the page genuinely being empty. Share this output so it can be fixed."
        )
    elif full_body_len < 300:
        print(
            "\n>>> The whole page has very little real HTML text, even outside the chosen container.\n"
            "    If the page visually looks content-rich, that content is likely delivered as images\n"
            "    or JavaScript-rendered widgets rather than crawlable HTML text — worth checking\n"
            "    view-source (not just the rendered page) in your browser to confirm."
        )
    return 0


def main(argv=None):
    args = parse_args(argv)

    if args.ai_analysis and not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
        print("ERROR: --ai-analysis requires ANTHROPIC_API_KEY (or ANTHROPIC_AUTH_TOKEN) to be set "
              "in the environment.", file=sys.stderr)
        return 1

    base_netloc = normalize_netloc(urlparse(args.base_url).netloc)

    session = requests.Session()
    session.headers.update({
        "User-Agent": args.user_agent,
        "Accept-Language": "en,ja;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })

    if args.debug_url:
        return debug_single_url(session, args.debug_url, args.en_prefix, args.timeout)

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
                              dup_titles, dup_meta, args.thin_words, args.thin_chars, args.en_prefix)

    ai_results = None
    if args.ai_analysis:
        keyword_map = load_keyword_map(args.ai_keyword_map) if args.ai_keyword_map else {}
        ai_results = run_ai_analysis(pages, args.ai_model, keyword_map)

    print("Writing workbook ...")
    wb = build_workbook(pages, issues, pairs, unmatched_en, unmatched_ja, broken_links, link_sources,
                         args.thin_words, args.thin_chars, args.en_prefix, ai_results)
    wb.save(args.output)

    counts = defaultdict(int)
    for i in issues:
        counts[i.severity] += 1
    print(f"\nDone. Wrote {args.output}")
    print(f"Issues found — High: {counts[HIGH]}, Medium: {counts[MEDIUM]}, Low: {counts[LOW]}, Info: {counts[INFO]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
