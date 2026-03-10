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

import anthropic
import feedparser
import schedule
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
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
    "Axios Pro Rata": "https://api.axios.com/feed/pro-rata",
    "Fortune Term Sheet": "https://fortune.com/feed/fortune-term-sheet/",
    "PitchBook News": "https://pitchbook.com/rss/news",
    "The Information": "https://www.theinformation.com/feed",
    "Sifted": "https://sifted.eu/feed",
    "EU-Startups": "https://www.eu-startups.com/feed/",
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

            raw_summary = entry.get("summary") or entry.get("description") or ""
            clean_summary = _strip_html(raw_summary)[:700]

            articles.append(
                {
                    "source": source,
                    "title": _strip_html(entry.get("title", "")).strip(),
                    "summary": clean_summary,
                    "link": entry.get("link", ""),
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
You are a senior growth equity analyst curating a daily investment brief for a VC/growth equity team.

Below are articles (0-indexed) fetched from VC/tech news sources in the past 24 hours.

──────────────────────────────────────────
{articles_block}
──────────────────────────────────────────

INCLUDE stories that fall into at least one of these categories:
• Funding rounds (€5M+ / $5.5M+) at Series A or beyond, in Europe, the US, or Israel
• M&A, acquisitions, exits, IPOs involving tech companies
• Notable product launches or significant market moves by growth-stage tech companies
• VC / growth equity fund news: new funds raised, LP activity, notable GP moves, fund strategies
• Market thesis or investment analysis pieces from reputable VC funds or analysts
  (e.g. "AI is bigger than SaaS", sector deep-dives, paradigm-shift arguments with investment implications)

EXCLUDE: seed rounds under €5M/$5.5M, consumer lifestyle, politics, sports, non-tech companies,
general macro news unrelated to tech investing, outage/incident reports, companies outside Europe/US/Israel.

For every INCLUDED article output EXACTLY these fields:
1. "headline": a 5–7 word headline in the format "Company/Fund X did/argues Y"
   (e.g. "Stripe launches new B2B payments product", "Acme raises €20M led by Sequoia",
   "NFX argues AI dwarfs SaaS opportunity")
2. "investor_relevance": exactly 2 sentences explaining why this matters for a growth equity investor.
   Be specific — mention round size, valuation, strategic angle, or market signal.
3. "sector": classify into exactly one of these sectors:
   Climate Solutions | Cybersecurity | Data & AI Infrastructure | Deep Tech |
   DevOps & DevTools | Digital Health | Fintech & Insurtech | Horizontal SW |
   Internet | Vertical SW | Others
4. "source_name": the publication name (e.g. "TechCrunch", "PitchBook", "Sifted")
5. "source_url": the article URL

Return ONLY a valid JSON array — no markdown fences, no extra text:
[
  {{
    "index": <int>,
    "sector": "<sector name>",
    "headline": "<5-7 word headline>",
    "investor_relevance": "<Sentence one. Sentence two.>",
    "source_name": "<publication name>",
    "source_url": "<article URL>"
  }}
]

If no articles qualify, return an empty array: []
"""


def filter_and_rank(articles: list[dict]) -> dict[str, list[dict]]:
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

    prompt = FILTER_PROMPT.format(articles_block=articles_block)

    log.info("Calling Claude (%s) to filter %d articles …", ANTHROPIC_MODEL, len(articles))
    response = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=4096,
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
        sector = str(item.get("sector", "Others"))
        if idx is None or not isinstance(idx, int) or idx >= len(articles):
            continue
        if sector not in result:
            sector = "Others"
        enriched = {
            **articles[idx],
            "headline": item.get("headline", articles[idx]["title"]),
            "investor_relevance": item.get("investor_relevance", ""),
            "source_name": item.get("source_name", articles[idx]["source"]),
            "source_url": item.get("source_url", articles[idx]["link"]),
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
        stories = digest.get(sector, [])[:5]  # cap at 5 per sector
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
        <td style="padding:20px 0 4px;">
          <div style="font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
                      color:#1a56db;border-bottom:1px solid #e0e0e0;padding-bottom:6px;">
            {_esc(sector)}
          </div>
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
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff;">
  <tr>
    <td align="center" style="padding:32px 20px;">
      <table width="800" cellpadding="0" cellspacing="0" border="0"
             style="max-width:800px;width:100%;">

        <!-- Header -->
        <tr>
          <td style="border-bottom:2px solid #111;padding-bottom:10px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td style="font-size:13px;font-weight:700;letter-spacing:1.5px;
                            text-transform:uppercase;color:#111;">
                  Daily Market Intel News
                </td>
                <td align="right" style="font-size:11px;color:#888;">
                  {_esc(date_str)}
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- Sections -->
        {sections_html}

      </table>
    </td>
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
        digest = filter_and_rank(articles)
        html = build_html(digest, date_str)
        send_email(html, date_str)
        log.info("=== Digest completed successfully ===")
    except Exception:
        log.exception("Digest run failed")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    log.info("VC News Digest Agent starting up.")
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
