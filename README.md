# Daily Job Digest Bot

Runs once a day for free on GitHub's servers, checks public job feeds for
junior/entry DevOps & Cloud roles in Cairo or remote, scores each one
against your CV, and updates `tracker.xlsx` in this repo. You just open
the file, glance at the top rows (best matches), click the links that
look good, and mark the "Applied?" column as you go.

**What it does NOT do:** log into any of your accounts, or click Apply for
you. LinkedIn and Indeed both forbid automated account access in their
Terms of Service, so this only reads public feeds (Wuzzuf's public RSS,
RemoteOK's public API) - no login required, no ToS risk.

## One-time setup (10 minutes)

1. **Create a GitHub account** if you don't have one: https://github.com/join

2. **Create a new repository**
   - Click the `+` in the top right of GitHub -> "New repository"
   - Name it something like `job-digest`
   - Set it to **Private** (so the digest isn't public) - private repos
     still get free GitHub Actions minutes for personal accounts
   - Click "Create repository"

3. **Upload these files** to the new repo. Easiest way: on the repo page,
   click "Add file" -> "Upload files", then drag in:
   - `job_digest.py`
   - `requirements.txt`
   - `tracker.xlsx`
   - the whole `.github` folder (drag the folder in directly - GitHub
     preserves the path)

   Commit them to the `main` branch.

4. **Check Actions is enabled**
   - Go to the "Actions" tab of your repo
   - If it asks you to enable workflows, click the button to enable them
   - You should see "Daily Job Digest" listed as a workflow

5. **Test it manually** (don't wait for tomorrow)
   - Actions tab -> "Daily Job Digest" -> "Run workflow" -> "Run workflow"
   - Wait ~30 seconds, refresh - you'll see a green checkmark when done
   - Open `tracker.xlsx` in the repo (GitHub previews Excel files in
     the browser) to see the results

## Using it day to day

- Every morning around 8-9am Cairo time, the workflow runs automatically
  and commits an updated `tracker.xlsx`.
- Open the file (GitHub's web preview works, or clone the repo and open
  it in Excel/LibreOffice for the full experience with the dropdown).
- Rows are sorted with the best CV matches first.
- Click a link in "Apply Link", apply on the actual site, then set
  "Applied?" for that row (it's a dropdown: No / Applied / Interviewing /
  Rejected / Not Interested).
- Your "Applied?" choices are never overwritten by later runs - the
  script matches rows by URL and leaves anything you've already set alone.
- If you edit the sheet locally, commit and push your changes back before
  the next scheduled run, or your edits could conflict with the bot's commit.

## Adjusting what it looks for

Open `job_digest.py` and edit the CONFIG section near the top:

- `TITLE_MUST_INCLUDE` / `TITLE_EXCLUDE` - which job titles count
- `JUNIOR_HINTS` - words that boost a posting's match score as junior-friendly
- `CV_SKILLS` - keywords it scores postings against (keep this in sync
  whenever you update your real CV)
- `CAIRO_AREA_HINTS` - which cities count as "Cairo area" for Wuzzuf results

## Changing the schedule

Open `.github/workflows/daily-job-digest.yml` and edit the `cron` line.
Cron times are always in UTC. For example `"0 6 * * *"` = 6:00 UTC daily.

## Honest limitations

- **RemoteOK is global.** A "remote" tag doesn't guarantee the company
  accepts candidates from Egypt without visa/work-authorization issues -
  the sheet flags this so you check each listing yourself before applying.
- **Wuzzuf's feed reflects newly *posted* jobs**, not a live "still open"
  status - always confirm on the actual page before spending time on an
  application, the same way we found stale listings during manual search.
- **LinkedIn and Indeed are intentionally excluded** because they prohibit
  this kind of automated access. If you want their listings included too,
  the safest route is checking them yourself with the "Date posted: Past
  week" filter, or asking me to search for you (I run within a chat
  session and can always do this for you in a normal conversation, no
  automation needed).
