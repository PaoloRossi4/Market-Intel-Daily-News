#!/usr/bin/env python3
"""
Daily VC & Growth Equity News Digest Agent
-------------------------------------------
Fetches articles from VC/tech RSS feeds, filters and summarises them with
Claude, then sends a clean HTML email via Outlook SMTP every morning at 9 AM.

Usage:
    python news_digest.py           # runs scheduler (fires at 09:00 daily)
    python news_digest.py --now     # runs one digest immediately, then keeps scheduling
"""

import json
import logging
import os
import re
import smtplib
import sys
import time
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import anthropic
import feedparser
import openpyxl
import requests
import schedule
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from rapidfuzz import process as fuzz_process

load_dotenv()

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
# (FT session initialised after logging is configured — see _init_ft_session)
_FT_SESSION: requests.Session | None = None
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
RSS_FEEDS: dict[str, str] = {
    "TechCrunch": "https://techcrunch.com/feed/",
    "Financial Times": "https://www.ft.com/companies/technology?format=rss",
    "Wired": "https://feeds.wired.com/wired/index",
    "PitchBook News": "https://pitchbook.com/rss/news",
    "Sifted": "https://sifted.eu/feed",
    "EU-Startups": "https://www.eu-startups.com/feed/",
    "Tech.eu": "https://tech.eu/feed/",
    "Tech Funding News": "https://techfundingnews.com/feed/",
    "Pulse 2.0": "https://pulse2.com/venture-capital/feed/",
}

ANTHROPIC_MODEL = "claude-sonnet-4-6"
SEND_TIME = "09:00"  # local time, 24 h

SECTORS: list[str] = [
    "Climate Solutions",
    "Cybersecurity",
    "Data & AI Infrastructure",
    "Deep Tech",
    "DevOps & DevTools",
    "Digital Health",
    "Fintech & Insurtech",
    "Horizontal SW",
    "Internet",
    "Vertical SW",
    "Others",
]

# Maps Excel sub-sector labels → our canonical SECTORS
_EXCEL_SECTOR_MAP: dict[str, str] = {
    "Climate Solution":                 "Climate Solutions",
    "Cybersecurity":                    "Cybersecurity",
    "Data, Infrastructure, DevTools":   "Data & AI Infrastructure",
    "Data/AI Infrastructure, DevTools": "Data & AI Infrastructure",
    "DeepTech":                         "Deep Tech",
    "Digital Health":                   "Digital Health",
    "Fintech/Insurtech":                "Fintech & Insurtech",
    "Horizontal SW":                    "Horizontal SW",
    "Internet":                         "Internet",
    "Others":                           "Others",
    "Vertical SW":                      "Vertical SW",
}

TAXONOMY_PATH = Path(__file__).parent / "SP Opp by vertical.xlsx"


def load_taxonomy() -> dict[str, str]:
    """Load {company_name_lower: sector} from the Excel portfolio file.
    Returns an empty dict if the file is missing (e.g. in CI/GitHub Actions)."""
    if not TAXONOMY_PATH.exists():
        log.warning("Taxonomy file not found at %s — skipping company lookup.", TAXONOMY_PATH)
        return {}
    try:
        wb = openpyxl.load_workbook(TAXONOMY_PATH, read_only=True, data_only=True)
        ws = wb["SP Opp by vertical"]
        taxonomy: dict[str, str] = {}
        for row in ws.iter_rows(min_row=3, values_only=True):
            name, subsector = row[1], row[3]
            if name and subsector:
                sector = _EXCEL_SECTOR_MAP.get(str(subsector).strip(), "Others")
                taxonomy[re.sub(r"[^a-z0-9 ]", "", str(name).lower()).strip()] = sector
        log.info("Loaded taxonomy: %d companies.", len(taxonomy))
        return taxonomy
    except Exception as exc:
        log.warning("Could not load taxonomy: %s", exc)
        return {}


def _normalise(name: str) -> str:
    """Lowercase, strip punctuation and extra spaces for name comparison."""
    return re.sub(r"[^a-z0-9 ]", "", name.lower()).strip()


def lookup_sector(company_name: str, taxonomy: dict[str, str], threshold: int = 97) -> str | None:
    """Return sector from taxonomy only when the company name is clearly the same.

    Strategy:
    1. Exact match after normalisation (e.g. 'Qevlar AI' == 'qevlar ai')
    2. Narrow fuzzy match (≥97) to catch trivial typos/abbreviations only
       (e.g. 'Qflow' ~ 'Qualis Flow')
    Different companies with similar-sounding names will NOT match.
    """
    if not taxonomy or not company_name:
        return None

    needle = _normalise(company_name)

    # 1. Exact match after normalisation
    if needle in taxonomy:
        log.info("Taxonomy exact match: %r", company_name)
        return taxonomy[needle]

    # 2. Narrow fuzzy match — only for near-identical names
    result = fuzz_process.extractOne(needle, taxonomy.keys(), score_cutoff=threshold)
    if result:
        matched_name, score, _ = result
        log.info("Taxonomy fuzzy match: %r → %r (score %d)", company_name, matched_name, score)
        return taxonomy[matched_name]

    return None


# Loaded once at import time
COMPANY_TAXONOMY: dict[str, str] = load_taxonomy()

# ---------------------------------------------------------------------------
# Deduplication cache
# ---------------------------------------------------------------------------

SEEN_CACHE_PATH = Path(__file__).parent / "seen_stories.json"
_DEDUP_TTL_HOURS = 168  # 7 days — prevents same story re-appearing across the week
_DEDUP_SIMILARITY_THRESHOLD = 90  # rapidfuzz score to consider two titles the same story


def _load_seen_cache() -> dict[str, str]:
    """Return {title_lower: iso_timestamp} for stories seen recently."""
    if not SEEN_CACHE_PATH.exists():
        return {}
    try:
        with open(SEEN_CACHE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_seen_cache(cache: dict[str, str]) -> None:
    try:
        with open(SEEN_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2)
    except Exception as exc:
        log.warning("Could not save dedup cache: %s", exc)


def deduplicate(digest: dict[str, list[dict]]) -> dict[str, list[dict]]:
    """Remove stories whose title is too similar to one seen in the last 48 h."""
    cache = _load_seen_cache()
    cutoff = datetime.now(tz=timezone.utc)

    # Expire old entries
    cache = {
        title: ts
        for title, ts in cache.items()
        if (cutoff - datetime.fromisoformat(ts)).total_seconds() < _DEDUP_TTL_HOURS * 3600
    }

    cached_titles = list(cache.keys())
    result: dict[str, list[dict]] = {s: [] for s in SECTORS}
    dropped = 0

    for sector, stories in digest.items():
        for story in stories:
            title = story.get("headline", story.get("title", "")).lower()
            if not title:
                result[sector].append(story)
                continue

            # Check similarity against all cached titles
            is_dup = False
            if cached_titles:
                match = fuzz_process.extractOne(title, cached_titles, score_cutoff=_DEDUP_SIMILARITY_THRESHOLD)
                if match:
                    log.info("Dedup: dropping %r (similar to cached %r, score %d)", title, match[0], match[1])
                    is_dup = True
                    dropped += 1

            if not is_dup:
                result[sector].append(story)
                cache[title] = cutoff.isoformat()
                cached_titles.append(title)  # prevent two stories in same batch matching each other

    _save_seen_cache(cache)
    if dropped:
        log.info("Deduplication removed %d repeated stor%s.", dropped, "y" if dropped == 1 else "ies")
    return result

# ---------------------------------------------------------------------------
# Financial Times – authenticated session & full-text scraping
# ---------------------------------------------------------------------------

_FT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


def _init_ft_session() -> requests.Session | None:
    """Log in to FT and return an authenticated session. Returns None if
    credentials are missing or login fails."""
    email = os.getenv("FT_EMAIL", "").strip()
    password = os.getenv("FT_PASSWORD", "").strip()
    if not email or not password:
        log.info("FT credentials not set — FT articles will use RSS summaries only.")
        return None

    try:
        s = requests.Session()
        s.headers.update(_FT_HEADERS)

        # Step 1: load login page, grab CSRF token and hidden fields
        r1 = s.get("https://accounts.ft.com/login", timeout=15)
        soup1 = BeautifulSoup(r1.text, "html.parser")

        def _val(name: str) -> str:
            tag = soup1.find("input", {"name": name})
            return tag["value"] if tag and tag.get("value") else ""

        csrf1 = _val("_csrf")
        location = _val("location") or "https://www.ft.com"

        # Step 2: submit email
        r2 = s.post(
            "https://accounts.ft.com/login",
            data={
                "_csrf": csrf1,
                "formType": "enter-email",
                "location": location,
                "noScript": "true",
                "email": email,
            },
            timeout=15,
        )
        soup2 = BeautifulSoup(r2.text, "html.parser")

        def _val2(name: str) -> str:
            tag = soup2.find("input", {"name": name})
            return tag["value"] if tag and tag.get("value") else ""

        csrf2 = _val2("_csrf") or csrf1

        # Step 3: submit password
        r3 = s.post(
            "https://accounts.ft.com/login",
            data={
                "_csrf": csrf2,
                "formType": "login",
                "location": location,
                "email": email,
                "password": password,
                "rememberMe": "true",
            },
            allow_redirects=True,
            timeout=15,
        )

        # Verify login succeeded by checking for FT session cookie
        if "FTSession" in s.cookies or "FTSession_s" in s.cookies or r3.url.startswith("https://www.ft.com"):
            log.info("FT login successful.")
            return s
        else:
            log.warning("FT login may have failed (no FT session cookie). Will use RSS summaries.")
            return None

    except Exception as exc:
        log.warning("FT login error: %s — will use RSS summaries.", exc)
        return None


def _fetch_ft_fulltext(url: str, session: requests.Session, char_limit: int = 2000) -> str:
    """Fetch and return the full text of an FT article (up to char_limit chars).
    Returns empty string on any failure."""
    try:
        r = session.get(url, timeout=15)
        if r.status_code != 200:
            return ""
        soup = BeautifulSoup(r.text, "html.parser")

        # FT article body selectors (in order of preference)
        for selector in [
            {"data-component": "article-body"},
            {"class": "article__content-body"},
            {"class": "n-content-body"},
        ]:
            body = soup.find(attrs=selector)
            if body:
                text = body.get_text(separator=" ", strip=True)
                return text[:char_limit]

        # Fallback: grab all <p> inside <article>
        article = soup.find("article")
        if article:
            paragraphs = " ".join(p.get_text(strip=True) for p in article.find_all("p"))
            return paragraphs[:char_limit]

    except Exception as exc:
        log.debug("FT full-text fetch failed for %s: %s", url, exc)
    return ""


# ---------------------------------------------------------------------------
# Step 1 – Fetch RSS articles
# ---------------------------------------------------------------------------

def _strip_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    return re.sub(r"\s+", " ", text).strip()


def fetch_articles(lookback_hours: int = 24) -> list[dict]:
    """Return articles published within the last *lookback_hours* hours."""
    cutoff = datetime.now(tz=timezone.utc) - timedelta(hours=lookback_hours)
    articles: list[dict] = []

    for source, url in RSS_FEEDS.items():
        log.info("Fetching %-25s  %s", source, url)
        try:
            feed = feedparser.parse(url, request_headers={"User-Agent": "vc-digest-bot/1.0"})
        except Exception as exc:
            log.warning("Could not fetch %s: %s", source, exc)
            continue

        for entry in feed.entries:
            published: datetime | None = None
            for attr in ("published_parsed", "updated_parsed"):
                raw = getattr(entry, attr, None)
                if raw:
                    try:
                        published = datetime(*raw[:6], tzinfo=timezone.utc)
                    except Exception:
                        pass
                    break

            # Skip articles older than the cutoff (unknown date: keep)
            if published and published < cutoff:
                continue

            article_link = entry.get("link", "")

            # For FT articles, attempt to fetch full text with authenticated session
            if source == "Financial Times" and _FT_SESSION is not None:
                full_text = _fetch_ft_fulltext(article_link, _FT_SESSION)
                if full_text:
                    clean_summary = full_text
                    log.debug("FT full text fetched for: %s", entry.get("title", "")[:60])
                else:
                    raw_summary = entry.get("summary") or entry.get("description") or ""
                    clean_summary = _strip_html(raw_summary)[:700]
            else:
                raw_summary = entry.get("summary") or entry.get("description") or ""
                clean_summary = _strip_html(raw_summary)[:700]

            articles.append(
                {
                    "source": source,
                    "title": _strip_html(entry.get("title", "")).strip(),
                    "summary": clean_summary,
                    "link": article_link,
                    "published": (
                        published.strftime("%Y-%m-%d %H:%M UTC") if published else "Unknown"
                    ),
                }
            )

    log.info("Fetched %d articles across all feeds.", len(articles))
    return articles


# ---------------------------------------------------------------------------
# Step 2 – Filter, summarise, and rank with Claude
# ---------------------------------------------------------------------------

FILTER_PROMPT = """\
You are a senior growth equity analyst curating a daily investment brief for the Eurazeo Growth team.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PRIMARY PURPOSE — what this newsletter is for
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. FUNDING ROUNDS (core content — be comprehensive)
   Report all noteworthy funding rounds. Do NOT cap the number of rounds included.
   • Europe: include from Seed upward if the company or sector is interesting. Err on the side of inclusion.
   • US: apply a high bar. Only include a US funding round if AT LEAST ONE of these is true:
       - A major tech corporation participates as investor (Google, Microsoft, Nvidia, Apple, Meta,
         Amazon, Salesforce, SAP, or similar large strategic investor)
       - A European tech investor participates (e.g. Index Ventures, Balderton, Northzone, HV Capital,
         Eurazeo, Partech, Idinvest, Breega, Notion Capital, Atomico, Creandum, EQT, Verdane,
         Maind Capital, or any other fund headquartered in Europe)
       - The company was founded by European founders (check founder names/backgrounds if mentioned)
       - The round is $50M or above in a relevant sector (AI, cybersecurity, fintech, deeptech,
         digital health, vertical/horizontal SaaS)
     Routine US rounds with only US investors and no European angle = EXCLUDE.
   • Israel: include strong AI, cyber, or deep tech rounds with clear relevance to European investors.
   • Do NOT apply a hard minimum size — a €2M seed in a genuinely novel European deep tech company
     is more relevant than a $10M US round in a commodity SaaS.

2. M&A (include if significant)
   Acquisitions, exits, IPOs, and secondary transactions involving tech companies in our sectors.
   Include cross-border deals and exits by known VC-backed companies.

3. GENERAL TECH NEWS (very selective — high bar)
   Only include if the story is a major, landscape-changing event:
   • A market-defining product launch by a sector leader (e.g. a new frontier AI model, a major
     platform shift, a new chip architecture)
   • A regulatory change with direct, immediate impact on the whole sector
   • A major strategic partnership or compute deal that reshapes competitive dynamics
   Do NOT include: opinion pieces, incremental product updates, minor partnerships, earnings commentary,
   analyst notes, company blog posts, or anything that would be forgotten in a week.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STRICT RULES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

• GEOGRAPHY: Europe and Israel are primary. US stories are secondary — apply a higher bar.
  Exclude all other geographies unless the deal is exceptionally large and globally relevant.
• FRESHNESS — this is critical: always assess whether the UNDERLYING EVENT is genuinely new,
  not just whether the article was published today. News sources frequently publish delayed
  coverage, recap articles, or "deal closes" weeks after the original announcement.
  Ask yourself: "Was this event (the fundraise, acquisition, launch) already widely known
  before the last 48–72 hours?" If yes, exclude it — the team will already know.
  Examples of what to EXCLUDE on freshness grounds:
    - "Google completes acquisition of Wiz" if the deal was announced months ago
    - "Company X closes $50M round" if the announcement was made 2+ weeks ago
    - Any article that is clearly a follow-up recap of an event the market already digested
  Use your own training knowledge to assess whether an event predates the last 48–72 hours.
  When in doubt and the article gives no clear indication of when the event actually happened,
  err on the side of exclusion.
• NO FORCING: It is perfectly fine if some sectors have zero stories on a given day.
  Never include a story just to fill a sector. Quality and relevance always over quantity.
• NO DUPLICATES ACROSS DAYS: See the "Already Reported" section below — strictly exclude
  any article covering the same event already sent, regardless of source or wording.
• CROSS-SOURCE SIGNAL: If a story appears across 3–4+ sources in today's batch, treat that
  as a strong signal of relevance and include it (picking the best-sourced article index).
• EXCLUDE ALWAYS: stock price/earnings/market cap commentary, macroeconomic news with no
  direct tech investing angle, consumer lifestyle, politics, sports, non-tech industries,
  outage/incident reports, opinion pieces without a concrete investment signal.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ALREADY REPORTED — last 7 days (DO NOT repeat)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{seen_block}
Exclude any article covering the same company + same event as above, even if worded differently
or from a different source. Only include a follow-up if it is a genuinely new development
(e.g. a second close, a product launch after a fundraise, regulatory approval after an IPO filing).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TODAY'S ARTICLES (0-indexed)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{articles_block}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

For every INCLUDED article output EXACTLY these fields:
1. "company_name": the primary company or fund name (e.g. "Stripe", "Mistral", "Sequoia")
2. "headline": a 5–7 word headline in the format "Company X did/announced Y"
   (e.g. "Mistral raises €600M Series B", "Wiz acquires Dazz for $450M",
   "Anthropic launches Claude 4 with agentic reasoning")
3. "investor_relevance": exactly 2 sentences. Be specific — mention round size, valuation,
   lead investor, strategic angle, or market signal. No vague generalities.
4. "sector": classify into exactly one sector using the definitions below.
   Use your own knowledge of what the company does — not just what the article says.
   Classify by core business, not by the topic of this specific news item.
   • Climate Solutions: cleantech, renewable energy software, carbon tracking, ESG data,
     energy transition infrastructure, smart grid, sustainability analytics
   • Cybersecurity: network security, endpoint protection, identity & access management,
     threat detection, AI security testing, SOC automation, zero trust, compliance software
   • Data & AI Infrastructure: AI/ML platforms, data pipelines, vector databases, cloud
     infrastructure, LLM tooling, AI agents infrastructure, observability, data warehousing
   • Deep Tech: frontier AI research, robotics, physical AI, world models, semiconductor IP,
     quantum computing, photonics, advanced materials, biotech, autonomous vehicles
   • DevOps & DevTools: CI/CD, developer productivity, code generation, testing automation,
     API management, platform engineering, software supply chain
   • Digital Health: healthcare SaaS, clinical AI, remote patient monitoring, mental health
     tech, medical device software, pharma tech, health data platforms
   • Fintech & Insurtech: payments, banking infrastructure, lending, wealth management,
     insurance tech, RegTech, embedded finance, crypto infrastructure
   • Horizontal SW: business process automation, ERP, CRM, HR tech, legal tech, procurement
     software, collaboration tools — applicable across all industries
   • Internet: consumer platforms, marketplaces, e-commerce enablement, creator economy,
     AdTech, social platforms, gaming
   • Vertical SW: industry-specific SaaS — construction tech, agritech, proptech, retail
     tech, logistics software, manufacturing software, education tech
   • Others: VC fund news, LP activity, macro investment theses, anything that does not
     clearly fit the above
5. "source_name": the publication name (e.g. "TechCrunch", "Sifted", "PitchBook")
6. "source_url": the article URL

Return ONLY a valid JSON array — no markdown fences, no extra text:
[
  {{
    "index": <int>,
    "company_name": "<name>",
    "sector": "<sector>",
    "headline": "<5-7 word headline>",
    "investor_relevance": "<Sentence one. Sentence two.>",
    "source_name": "<publication>",
    "source_url": "<URL>"
  }}
]

If no articles qualify, return an empty array: []
"""


def filter_and_rank(articles: list[dict], seen_headlines: list[str] | None = None) -> dict[str, list[dict]]:
    """Ask Claude to filter, summarise, and categorise the article list."""
    empty_result: dict[str, list[dict]] = {s: [] for s in SECTORS}

    if not articles:
        log.info("No articles to filter.")
        return empty_result

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError("ANTHROPIC_API_KEY is not set in .env")

    client = anthropic.Anthropic(api_key=api_key)

    articles_block = ""
    for i, a in enumerate(articles):
        articles_block += (
            f"[{i}] SOURCE: {a['source']} | DATE: {a['published']}\n"
            f"TITLE: {a['title']}\n"
            f"SUMMARY: {a['summary']}\n"
            f"URL: {a['link']}\n\n"
        )

    if seen_headlines:
        seen_block = "\n".join(f"• {h}" for h in seen_headlines)
        log.info("Passing %d recently seen headlines to Claude for semantic dedup.", len(seen_headlines))
    else:
        seen_block = "(none — this is the first digest or no recent history)"

    prompt = FILTER_PROMPT.format(articles_block=articles_block, seen_block=seen_block)

    log.info("Calling Claude (%s) to filter %d articles …", ANTHROPIC_MODEL, len(articles))
    response = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )

    raw = response.content[0].text.strip()

    # Strip accidental markdown fences
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
        raw = raw.strip()

    try:
        ranked: list[dict] = json.loads(raw)
    except json.JSONDecodeError as exc:
        log.error("Claude response is not valid JSON (%s). First 500 chars:\n%s", exc, raw[:500])
        return empty_result

    result: dict[str, list[dict]] = {s: [] for s in SECTORS}
    for item in ranked:
        idx = item.get("index")
        if idx is None or not isinstance(idx, int) or idx >= len(articles):
            continue

        # 1. Claude's classification is always primary
        company_name = item.get("company_name", "")
        claude_sector = str(item.get("sector", "Others"))

        # 2. Excel taxonomy used only as a hint for small/unknown companies —
        #    never overrides Claude
        taxonomy_sector = lookup_sector(company_name, COMPANY_TAXONOMY)
        if taxonomy_sector and taxonomy_sector != claude_sector:
            log.info(
                "Taxonomy hint for %r: Excel=%r, Claude=%r → keeping Claude",
                company_name, taxonomy_sector, claude_sector,
            )

        sector = claude_sector
        if sector not in result:
            sector = "Others"

        enriched = {
            **articles[idx],
            "headline": item.get("headline", articles[idx]["title"]),
            "investor_relevance": item.get("investor_relevance", ""),
            "source_name": item.get("source_name", articles[idx]["source"]),
            "source_url": item.get("source_url", articles[idx]["link"]),
            "sector_source": "claude",
        }
        result[sector].append(enriched)

    total = sum(len(v) for v in result.values())
    sector_summary = "  ".join(f"{s}:{len(result[s])}" for s in SECTORS if result[s])
    log.info("Claude selected %d articles → %s", total, sector_summary)
    return result


# ---------------------------------------------------------------------------
# Step 3 – Build HTML email
# ---------------------------------------------------------------------------

def _esc(text: str) -> str:
    """Minimal HTML entity escaping."""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def build_html(digest: dict[str, list[dict]], date_str: str) -> str:
    # Only render sectors that have at least one story
    sections_html = ""
    for sector in SECTORS:
        stories = digest.get(sector, [])
        if not stories:
            continue

        stories_html = ""
        for s in stories:
            stories_html += f"""
          <tr>
            <td style="padding:12px 0;border-bottom:1px solid #f0f0f0;vertical-align:top;">
              <div style="font-size:13px;font-weight:700;color:#111;line-height:1.4;margin-bottom:5px;">
                {_esc(s['headline'])}
              </div>
              <div style="font-size:12px;color:#444;line-height:1.6;margin-bottom:5px;">
                {_esc(s['investor_relevance'])}
              </div>
              <div style="font-size:11px;color:#888;">
                <a href="{_esc(s['source_url'])}" target="_blank"
                   style="color:#1a56db;text-decoration:none;">{_esc(s['source_name'])}</a>
              </div>
            </td>
          </tr>"""

        sections_html += f"""
      <tr>
        <td style="padding:20px 0 0 0;">
          <!-- Sector header bar — inset, navy background -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td style="background:#0a1628;padding:7px 12px;">
                <span style="font-size:10px;font-weight:700;letter-spacing:2px;
                             text-transform:uppercase;color:#ffffff;">
                  {_esc(sector)}
                </span>
              </td>
            </tr>
            <!-- Teal accent line below sector header -->
            <tr>
              <td style="height:2px;background:#00d4aa;font-size:0;line-height:0;">&nbsp;</td>
            </tr>
          </table>
          <!-- Stories -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            {stories_html}
          </table>
        </td>
      </tr>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Daily Market Intel News - {_esc(date_str)}</title>
</head>
<body style="margin:0;padding:0;background:#ffffff;font-family:'Helvetica Neue',Arial,sans-serif;">

<!-- Branded header — full width, dark navy, compact -->
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="background:#0a1628;">
  <tr>
    <td style="padding:16px 32px 12px 32px;">
      <div style="font-size:18px;font-weight:700;color:#ffffff;letter-spacing:0.5px;line-height:1.2;">
        Eurazeo Growth
      </div>
      <div style="font-size:11px;color:#a0b4cc;margin-top:4px;letter-spacing:0.3px;">
        Daily Market Intel News &nbsp;&middot;&nbsp; {_esc(date_str)}
      </div>
    </td>
  </tr>
  <tr>
    <td style="height:3px;background:#00d4aa;font-size:0;line-height:0;">&nbsp;</td>
  </tr>
</table>

<!-- Content -->
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff;">
  <tr>
    <td style="padding:24px 32px 8px 32px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        {sections_html}
      </table>
    </td>
  </tr>
</table>

<!-- Footer -->
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:24px;">
  <tr>
    <td style="height:3px;background:#00d4aa;font-size:0;line-height:0;">&nbsp;</td>
  </tr>
  <tr>
    <td style="height:24px;background:#0a1628;font-size:0;line-height:0;">&nbsp;</td>
  </tr>
</table>

</body>
</html>"""


# ---------------------------------------------------------------------------
# Step 4 – Send via Outlook SMTP
# ---------------------------------------------------------------------------

def send_email(html_body: str, date_str: str) -> None:
    sender = os.getenv("GMAIL_USER", "").strip()
    password = os.getenv("GMAIL_APP_PASSWORD", "").strip()
    recipient = os.getenv("RECIPIENT_EMAIL", "").strip()

    missing = [k for k, v in [
        ("GMAIL_USER", sender),
        ("GMAIL_APP_PASSWORD", password),
        ("RECIPIENT_EMAIL", recipient),
    ] if not v]
    if missing:
        raise EnvironmentError(f"Missing .env variable(s): {', '.join(missing)}")

    subject = f"Daily Market Intel News - {date_str}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient

    # Plain-text fallback
    plain = (
        f"Daily Market Intel News - {date_str}\n\n"
        "This email is best viewed in an HTML-capable email client."
    )
    msg.attach(MIMEText(plain, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    log.info("Connecting to smtp.gmail.com:587 …")
    with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(sender, password)
        smtp.sendmail(sender, [recipient], msg.as_string())

    log.info("✓ Digest sent to %s", recipient)


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------

def run_digest() -> None:
    log.info("=" * 60)
    log.info("Starting daily VC digest run")
    log.info("=" * 60)
    date_str = datetime.now().strftime("%B %d, %Y")
    try:
        articles = fetch_articles()
        # Load recent headlines so Claude can skip semantic duplicates from past 7 days
        recent_cache = _load_seen_cache()
        seen_headlines = list(recent_cache.keys())
        digest = filter_and_rank(articles, seen_headlines=seen_headlines)
        digest = deduplicate(digest)
        html = build_html(digest, date_str)

        # Wait until exactly SEND_TIME before delivering
        send_h, send_m = map(int, SEND_TIME.split(":"))
        target = datetime.now().replace(hour=send_h, minute=send_m, second=0, microsecond=0)
        wait_secs = (target - datetime.now()).total_seconds()
        if wait_secs > 0:
            log.info("Digest ready — holding %.0f s until %s to send …", wait_secs, SEND_TIME)
            time.sleep(wait_secs)

        send_email(html, date_str)
        log.info("=== Digest completed successfully ===")
    except Exception:
        log.exception("Digest run failed")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    global _FT_SESSION
    log.info("VC News Digest Agent starting up.")
    _FT_SESSION = _init_ft_session()
    log.info("Digest scheduled at %s local time every day.", SEND_TIME)

    schedule.every().day.at(SEND_TIME).do(run_digest)

    if "--once" in sys.argv:
        log.info("--once flag detected: running digest once and exiting.")
        run_digest()
        return

    if "--now" in sys.argv:
        log.info("--now flag detected: running digest immediately.")
        run_digest()

    log.info("Scheduler running. Press Ctrl-C to stop.")
    try:
        while True:
            schedule.run_pending()
            time.sleep(30)
    except KeyboardInterrupt:
        log.info("Shutting down.")


if __name__ == "__main__":
    main()
