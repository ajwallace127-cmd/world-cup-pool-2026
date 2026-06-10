# ⚽ World Cup Pool 2026 — Standings Tracker

A self-contained system that reads your Google Forms pool entries and generates a live leaderboard website — automatically updated from ESPN every 2 hours during the tournament.

---

## How it works

1. Your **Google Form** collects picks → saved to a **Google Sheet**
2. You download the sheet as `.xlsx`, rename it `pool_entries.xlsx`, and commit it to this repo **once** (entries close June 11)
3. **GitHub Actions** runs every 2 hours, fetches live scores from ESPN, regenerates the leaderboard, and deploys it automatically
4. Participants visit your **GitHub Pages URL** to see live standings

---

## Complete GitHub setup guide

### Part 1 — Create a GitHub account (skip if you already have one)

1. Go to [github.com](https://github.com) and click **Sign up**
2. Enter an email, password, and username (your username becomes part of your site URL, e.g. `alexwallace.github.io/...`)
3. Verify your email and finish onboarding — the free plan is all you need

---

### Part 2 — Create a new repository

A **repository** (repo) is just a folder on GitHub that holds your project files.

1. Once signed in, click the **+** icon in the top-right corner → **New repository**
2. Fill in the details:
   - **Repository name:** `world-cup-pool-2026` (no spaces)
   - **Description:** World Cup Pool standings tracker (optional)
   - **Visibility:** ✅ **Public** — required for free GitHub Pages hosting
   - Leave "Initialize with README" **unchecked** (you'll upload your own files)
3. Click **Create repository**

You'll land on an empty repo page. Keep this tab open.

---

### Part 3 — Upload the project files

You need to upload everything from the `World Cup Pool` folder on your Mac, **including** the hidden `.github` folder that contains the auto-update workflow.

**The files GitHub needs:**

```
process_pool.py
pool_entries.xlsx
requirements.txt
index.html
README.md
images/               ← folder (upload photos here, see below)
.github/
  workflows/
    update-standings.yml
```

**How to upload:**

> **Note on the `.github` folder:** Mac Finder hides folders that start with a dot. The easiest way to upload it is via GitHub's web interface by typing the path manually (covered below).

**Option A — Upload everything at once via GitHub web (easiest):**

1. On your empty repo page, click **uploading an existing file** (or **Add file → Upload files**)
2. Open Finder, press **Cmd+Shift+.** to show hidden files — you should now see the `.github` folder
3. Select all files and the `.github` folder and drag them into the GitHub upload area
4. Scroll down, leave the commit message as-is, and click **Commit changes**

**Option B — Upload the `.github` workflow file manually:**

If drag-and-drop doesn't pick up the `.github` folder:

1. Upload all the regular files first (drag everything except `.github`)
2. Click **Commit changes**
3. Then click **Add file → Create new file**
4. In the filename box at the top, type exactly: `.github/workflows/update-standings.yml`
   - GitHub will automatically create the nested folders as you type slashes
5. Open the file `update-standings.yml` from your Mac in a text editor, copy all the text, and paste it into the GitHub editor
6. Click **Commit new file**

---

### Part 4 — Add your pool entries

After entries close:

1. Open your Google Sheet with the form responses
2. Click **File → Download → Microsoft Excel (.xlsx)**
3. Rename the downloaded file to exactly `pool_entries.xlsx`
4. In your GitHub repo, click **Add file → Upload files**, drag in the file, and commit

---

### Part 5 — Upload the background photos

1. In your repo, click **Add file → Upload files**
2. Before dropping files, click into the path box at the top of the page (it shows `world-cup-pool-2026 /`) and type `images/` — this tells GitHub to put the files inside an `images` subfolder
3. Drag in your 6 renamed photos (see photo list below) and click **Commit changes**

| Filename | Tab it appears on |
|----------|-------------------|
| `hero.jpg` | Main banner at the top |
| `standings_bg.jpg` | Standings tab |
| `groups_bg.jpg` | Group Stage tab |
| `bracket_bg.jpg` | Bracket tab |
| `scorers_bg.jpg` | Goal Scorers tab |
| `allpicks_bg.jpg` | All Picks tab |

Photos appear as a dark atmospheric strip behind each tab header. The site looks great without them too — it falls back to a deep green gradient.

> **Preview locally with photos:** Photos only load when served over HTTP, not from a file. To preview: open Terminal, `cd` to your project folder, run `python3 -m http.server`, and visit `http://localhost:8000` in your browser.

---

### Part 6 — Enable GitHub Pages (your public URL)

1. In your repo, click **Settings** (top tab bar)
2. In the left sidebar, click **Pages**
3. Under **Source**, click the dropdown and select **Deploy from a branch**
4. Set Branch: `main`, Folder: `/ (root)` → click **Save**
5. Wait about 60 seconds, then refresh the page — your live URL will appear:
   ```
   https://YOUR-USERNAME.github.io/world-cup-pool-2026/
   ```
6. Share this URL with your pool participants

---

### Part 7 — Enable auto-update permissions

GitHub Actions needs permission to push the updated `index.html` back to your repo every 2 hours.

1. In your repo, click **Settings**
2. In the left sidebar, click **Actions → General**
3. Scroll to **Workflow permissions**
4. Select **Read and write permissions**
5. Click **Save**

---

### Part 8 — Verify the auto-update is working

1. In your repo, click the **Actions** tab
2. You should see a workflow called **Update Standings**
3. To trigger it manually right now: click **Update Standings** → **Run workflow** → **Run workflow** (green button)
4. The workflow will run, generate a fresh `index.html`, and push it to your repo
5. Your GitHub Pages site will reflect the update within a minute

If the workflow shows a red ✗, click it to see the error log — common causes are missing files or the workflow permissions not being set in Step 7.

---

### Making changes after setup

**To re-run an update immediately:** Actions tab → Update Standings → Run workflow

**To update entries** (if someone submits late): re-download the Google Sheet as `.xlsx`, rename it `pool_entries.xlsx`, and upload it to the repo — the next auto-run picks it up

**To edit any setting** (scoring, team aliases, etc.): click the file in your repo → click the pencil ✏️ icon → edit → Commit changes — GitHub Actions picks it up on the next run

---

## Adding the background photos

The site has image hooks for 6 World Cup photos — one per tab. Save your photos to an `images/` folder in the repo with these exact filenames:

| File | Used on |
|------|---------|
| `images/hero.jpg` | Main banner (e.g. Van Persie diving header) |
| `images/standings_bg.jpg` | Standings tab strip |
| `images/groups_bg.jpg` | Group Stage tab strip (e.g. Maradona) |
| `images/bracket_bg.jpg` | Bracket tab strip (e.g. Götze goal) |
| `images/scorers_bg.jpg` | Scorers tab strip (e.g. Ronaldo running) |
| `images/allpicks_bg.jpg` | All Picks tab strip (e.g. England celebrating) |

The site looks great without them too — each tab falls back to a dark pitch-green gradient.

---

## During the tournament

### Updating entries

If late entries come in before the June 11 deadline, just re-download your Google Sheet and re-upload `pool_entries.xlsx`. The next automated run will pick up the changes.

### Goal scorer tracking

Goal scorer data is fetched **automatically** — the script pulls per-game summaries from ESPN for every completed match and aggregates goals by player. No manual updates needed during the tournament.

If ESPN's scorer data is wrong for a specific game, you can override individual players using `GOAL_SCORER_OVERRIDE` in `process_pool.py`:

```python
GOAL_SCORER_OVERRIDE: dict = {
    "Kylian Mbappe": 5,   # override if ESPN count is off
}
```

The scorer matcher handles accented names automatically (Mbappé → Mbappe) and uses fuzzy matching so partial names like "Vinicius" will still match "Vinicius Junior".

### Group advancement bonuses

These award automatically once ESPN reports knockout round games. If you want to award the bonus immediately after the group stage ends (before knockout games start), add team names to `MANUAL_ADVANCED`:

```python
MANUAL_ADVANCED: set = {
    "Spain", "France", "Brazil",  # ... etc
}
```

### Manual trigger

To force an immediate update, go to your repo's **Actions** tab → select **Update Standings** → click **Run workflow**.

---

## Scoring reference

| Event | Points |
|-------|--------|
| Win (regulation or ET/PK) | 300 |
| Draw (regulation) | 100 |
| ET/PK loss | 100 |
| Regulation loss | 0 |
| Group stage advancement — Tiers A & B | +100 |
| Group stage advancement — Tiers C & D | +200 |
| Correct tournament winner | +450 |
| Each goal by a picked goal scorer | +150 |

---

## Troubleshooting

### Picks aren't being detected

Run the script locally and check the printed column map:
```bash
pip install -r requirements.txt
python process_pool.py pool_entries.xlsx
```

Look for output like:
```
  Column map:
    Tier A: ['Tier A (Favorites) [Pick 2]']
    Tier B: ['Tier B (Dark Horses) [Pick 4]']
    ...
```

If a tier shows `[]`, the column name didn't match. Open `process_pool.py`, find `_detect_columns()`, and add a pattern matching your actual column name to `tier_patterns`.

### ESPN not finding a team

Add an alias to `TEAM_ALIASES` in `process_pool.py`:
```python
TEAM_ALIASES = {
    "usa": "united states",
    "south korea": "korea republic",
    # Add more as needed
}
```

### Running locally

```bash
pip install -r requirements.txt
python process_pool.py pool_entries.xlsx
open index.html      # macOS
start index.html     # Windows
```

---

## Website features

- **Live leaderboard** — all participants ranked by score, searchable
- **Expandable rows** — click any row to see all 12 team picks + goal scorers + winner pick, color-coded by status
- **Pick Popularity charts** — per-tier bar charts showing how picks were distributed, plus winner and goal scorer popularity
- **Team Tracker** — all 48 teams with current status, record, and how many participants picked them
- **All Picks grid** — every participant's 12 picks + winner + scorers in one scrollable table, color-coded by team status, with name filter
- **Auto-refresh** — GitHub Actions re-runs every 2 hours and pushes updated standings
