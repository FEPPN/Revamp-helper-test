# Setup

One-time setup, done once by whoever owns/hosts the tool. Copywriters never
touch anything in this file — they only get a link and a form.

## 1. Python dependencies (only needed for local testing, not for the hosted app)

```
python -m pip install -r requirements.txt
```

## 2. SerpAPI (live Google SERP data)

1. Create an account at https://serpapi.com (free tier: 250 searches/month).
2. Copy your key from https://serpapi.com/manage-api-key.
3. Copy `.env.example` to `.env` and set `SERPAPI_API_KEY` (for local testing).

## 3. Ahrefs — CSV upload, no API key needed

No Ahrefs API key is required. In the app, whoever runs a report exports a
CSV by hand from Ahrefs' normal UI first (this doesn't need admin
permissions, unlike creating an API key):

Ahrefs → Keywords Explorer → paste the keyword → **Matching terms** → Terms
match: **All** → **Export** → CSV. Upload that file in the app's form.

(If an Ahrefs API key becomes available later — someone with workspace
admin rights creates one under Account settings > API keys — this step can
be swapped for a direct API call without touching anything else.)

## 4. Google Search Console — CSV upload, no Google Cloud Console needed

Same idea as Ahrefs: no OAuth, no Google Cloud project, no admin permission
required — just the Search Console UI you already have access to.

1. Open https://search.google.com/search-console and select the
   papernest.com property.
2. **Performance** report (left menu).
3. Click **+ New** (filters bar) → **Page** → **Exact URL** → paste the
   exact target page URL for the keyword.
4. Set the date range (e.g. last 6 months, top right).
5. Make sure the table below is showing **Queries** (not Pages/Countries/...).
6. Click **Export** (top right) → **Download CSV**.
7. Upload that file in the app's form.

## 5. Hosted app for copywriters

This turns the whole pipeline into a web page copywriters open with a link
— no Python, no Antigravity, no keys on their end, and (as of this version)
no Google Cloud Console / OAuth setup either.

### 5.1 Push the code to GitHub

1. Create a **private** GitHub repo (private because it's an internal tool,
   even though it holds no real secrets itself).
2. Push this whole folder to it (`git init`, `git add .`, `git commit`, `git push`
   — `.env` and `secrets.toml` are already gitignored, so real credentials
   never get pushed).

### 5.2 Deploy on Streamlit Community Cloud (free)

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **New app** → pick your repo → set **Main file path** to `app.py`.
3. Go to **Settings → Secrets** for this app and paste the content of
   `secrets.toml.example`, filled in with your real SerpAPI key. This is the
   *only* place this credential lives — not in the repo, not on any
   copywriter's machine.
4. Deploy. You get a URL like `https://<something>.streamlit.app` — that's
   the link to share with the copywriting team.

### 5.3 Using the app

Whoever wants a report: opens the link, fills in the keyword + brand,
exports the two CSVs by hand (Ahrefs + Search Console, steps 3 and 4 above),
uploads them, clicks "Générer le rapport", downloads the finished Excel
file. SERP data and competitor page structure are fetched automatically —
no manual step for those.

## (Optional, legacy) Google Search Console via API instead of CSV export

`scripts/fetch_gsc.py` still exists for anyone who *does* have Google Cloud
Console access and wants live API queries instead of manual CSV exports —
useful for the command-line workflow in SKILL.md, not required for the
hosted app. See its docstring for the OAuth setup if you want to use it;
skip this entirely if the CSV export flow above works for you.

## Security

- `.env` and `secrets.toml` are real credentials. They are gitignored — keep it that way, never force-add them.
- Never paste real key values into `SKILL.md`, prompts, or output files. If you accidentally commit a real key, rotate it immediately (don't just remove it from a later commit — it stays in git history).
