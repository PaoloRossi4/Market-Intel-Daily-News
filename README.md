# Daily VC & Growth Equity News Digest

A Python agent that runs every morning at **9 AM**, fetches VC and tech funding
news from seven RSS sources, filters and summarises them with **Claude**, and
delivers a clean HTML email to your inbox via **Outlook SMTP**.

---

## Project structure

```
vc-news-digest/
├── news_digest.py      # main agent
├── requirements.txt
├── .env.example
└── README.md
```

---

## Quick-start

### 1. Clone / copy the project

```bash
cd ~/vc-news-digest
```

### 2. Create and activate a virtual environment

```bash
python -m venv .venv

# macOS / Linux
source .venv/bin/activate

# Windows (PowerShell)
.\.venv\Scripts\Activate.ps1

# Windows (CMD)
.\.venv\Scripts\activate.bat
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure credentials

Copy the example file and fill in your values:

```bash
cp .env.example .env
```

Open `.env` in any editor and set:

| Variable | Description |
|---|---|
| `ANTHROPIC_API_KEY` | Your Anthropic API key (from console.anthropic.com) |
| `OUTLOOK_USER` | Your Microsoft 365 / Outlook email address |
| `OUTLOOK_PASSWORD` | Your Outlook password **or** App Password (see below) |
| `RECIPIENT_EMAIL` | Where to deliver the digest (can match `OUTLOOK_USER`) |

#### Outlook App Password (required if 2FA is enabled)

If your Microsoft account has two-factor authentication (MFA) enabled, SMTP
login with your normal password will be rejected. You must create an
**App Password**:

1. Go to **https://account.microsoft.com/security**
2. Sign in with your Microsoft account.
3. Click **"Advanced security options"**.
4. Under **"App passwords"**, click **"Create a new app password"**.
5. Copy the generated password and paste it as `OUTLOOK_PASSWORD` in your
   `.env` file.

> **Corporate / Exchange accounts:** If your organisation uses Microsoft 365
> with Conditional Access or Basic Auth disabled, ask your IT admin to enable
> SMTP AUTH for your mailbox:
> *Microsoft 365 admin center → Users → Active users → [your user] →
> Mail → Manage email apps → enable "Authenticated SMTP"*.

### 5. Test immediately

Run a single digest right now without waiting for 9 AM:

```bash
python news_digest.py --now
```

You should see log lines ending with `✓ Digest sent to you@example.com`.
Check your inbox – the email usually arrives within a minute.

### 6. Start the daily scheduler

```bash
python news_digest.py
```

The process will block and fire the digest every day at **09:00 local time**.
Keep it running in the background (see options below).

---

## Keeping the agent running

### Option A – `screen` or `tmux` (Linux / macOS)

```bash
screen -S vc-digest
python news_digest.py
# Ctrl-A D  to detach
```

### Option B – Windows Task Scheduler

1. Open **Task Scheduler** → Create Basic Task.
2. Trigger: **Daily at 09:00**.
3. Action: **Start a program**
   - Program: `C:\path\to\vc-news-digest\.venv\Scripts\python.exe`
   - Arguments: `C:\path\to\vc-news-digest\news_digest.py --now`
4. Use `--now` so the Task Scheduler handles the timing instead of the
   internal `schedule` loop.

### Option C – `systemd` service (Linux)

Create `/etc/systemd/system/vc-digest.service`:

```ini
[Unit]
Description=VC News Digest Agent
After=network.target

[Service]
ExecStart=/path/to/.venv/bin/python /path/to/news_digest.py
WorkingDirectory=/path/to/vc-news-digest
Restart=on-failure
User=youruser

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl enable --now vc-digest
```

---

## Customisation

| What | Where |
|---|---|
| Change send time | `SEND_TIME = "09:00"` in `news_digest.py` |
| Add / remove RSS feeds | `RSS_FEEDS` dict in `news_digest.py` |
| Adjust filtering criteria | `FILTER_PROMPT` string in `news_digest.py` |
| Change Claude model | `ANTHROPIC_MODEL` in `news_digest.py` |
| Cap stories per section | `.[:5]` slice in `build_html()` |

---

## RSS sources

| Source | URL |
|---|---|
| TechCrunch | https://techcrunch.com/feed/ |
| Axios Pro Rata | https://api.axios.com/feed/pro-rata |
| Fortune Term Sheet | https://fortune.com/feed/fortune-term-sheet/ |
| PitchBook News | https://pitchbook.com/rss/news |
| The Information | https://www.theinformation.com/feed |
| Sifted | https://sifted.eu/feed |
| EU-Startups | https://www.eu-startups.com/feed/ |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `SMTPAuthenticationError` | Wrong password or App Password needed (see §4) |
| `SMTP AUTH extension not supported` | Enable SMTP AUTH in M365 admin (see §4) |
| Empty digest / no stories | Check feed connectivity; try `--now` to see logs |
| `JSONDecodeError` from Claude | Transient API issue; the agent will retry next day |
| `ANTHROPIC_API_KEY not set` | Ensure `.env` is in the same folder you run from |
