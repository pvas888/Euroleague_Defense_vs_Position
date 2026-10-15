# Euroleague defense vs position (PIR) - automatic web page

A web page that shows, for every team, how much PIR the opposing guards, forwards and
centers produce against it. A free GitHub job refreshes it every night, so it stays
current after each round with no action from you.

## One-time setup (about 10 minutes)

1. **Create a GitHub account** (free) at github.com if you don't have one.
2. **New repository**: click **+ -> New repository**. Name it e.g. `euroleague-dvp`.
   Choose **Public** (free GitHub Pages needs a public repository). Create it.
3. **Upload the files from this folder**: on the new repository page click
   **uploading an existing file** and drag in *everything* from the unzipped folder,
   including the hidden `.github` folder (drag the whole folder contents, not just the
   visible files). Click **Commit changes**.
   *If the `.github` folder did not upload:* click **Add file -> Create new file**,
   type `.github/workflows/update.yml` as the name (typing the slashes creates the
   folders), paste the contents of `update.yml`, and commit.
4. **Allow the job to save its results**: **Settings -> Actions -> General ->
   Workflow permissions -> Read and write permissions -> Save**.
5. **Run it once**: **Actions** tab -> **Update defense-vs-position page** ->
   **Run workflow**. The first run downloads the whole season so far (a few minutes).
6. **Turn on the web page**: **Settings -> Pages -> Build and deployment ->
   Source: Deploy from a branch -> Branch: `main`, folder `/docs` -> Save.**
7. After a minute your page is at `https://<your-username>.github.io/<repo-name>/`.
   Bookmark it on your phone.

## What happens afterwards

- Every night at 23:30 UTC and every morning at 08:00 UTC the job downloads only the
  games played since the last run, rebuilds the page and saves it. If nothing new was
  played, nothing changes.
- The "updated ..." line at the top of the page shows when it last refreshed.
- Want it right now? **Actions -> Update defense-vs-position page -> Run workflow**.

## Keeping it healthy

- **New players / transfers**: the positions come from `data/positions.csv`
  (columns `ID, Name, Surname, Team, Position`). Players missing from it get an
  estimated position and are flagged on the page. To fix or add players, edit that file
  on GitHub (pencil icon) and run the job again.
- **Next season**: change `DVP_SEASONS: "2026"` in `.github/workflows/update.yml`
  (use `"2027"` for 2027-28).
- **If the page stops updating**: open the **Actions** tab. GitHub switches off
  scheduled jobs in repositories with no activity for a long time; click
  **Enable workflow** if you see that. If a run shows a red cross, open it to read the
  error and send it to Claude.
- **Euroleague data access**: the job needs to reach the Euroleague data service from
  GitHub's servers. If every run fails with connection errors there, run the Colab
  version instead; it produces the same `index.html` you can open or share.

## Files

| File | Purpose |
|---|---|
| `euroleague_defense_vs_position.py` | Downloads games, attaches positions, builds the tables and page (same script as the Colab one) |
| `.github/workflows/update.yml` | The schedule |
| `data/positions.csv` | Your players and their positions |
| `data/cache/` | Saved box scores (created by the job) |
| `docs/index.html` | The web page |
