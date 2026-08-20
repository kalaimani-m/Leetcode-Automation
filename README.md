# LeetCode Daily Student Progress Report

Automatically checks the LeetCode profiles of all students in `students/leetcode_Links.xlsx`,
builds a daily Excel report, and emails it to `kalaimani.cybersec@prathyusha.edu.in` every
day at 9:00 AM IST via GitHub Actions.

## What it reports

Main sheet — exactly these 6 columns: S.No., Registration Number, Student Name, Current Date,
Problems Solved on Current Date, Overall Problems Solved.

A second **Notes** sheet lists any student whose profile was unavailable, or whose "today"
count might be a lower bound (see limitation below).

## Known limitation (please read)

LeetCode has no official API. This uses LeetCode's own public GraphQL endpoint, which only
exposes a student's most recent ~20 accepted submissions -- not a full daily log. If a student
accepts more than ~20 problems in one day, or has accepted submissions between their last one
today and the time the report runs, "today's" count can be an undercount. The script flags this
in the Notes sheet whenever it detects the recent-activity window was full. There's no public
LeetCode endpoint that avoids this; it's a genuine ceiling on accuracy, not a bug.

## One-time setup

1. **Add these files to your repo** (`kalaimani-m/Leetcode-Automation`):
   - `main.py`, `leetcode_client.py`, `requirements.txt`
   - `.github/workflows/daily-report.yml`
   - `students/leetcode_Links.xlsx` (your master list -- commit the real file here so the
     GitHub Actions runner can read it; it has no access to your local uploads)

   Easiest way: on the repo page, "Add file" -> "Upload files", drag all of the above in
   (recreate the `.github/workflows/` and `students/` folders by typing the path in the
   filename box), then commit.

2. **Create a Gmail App Password** to send from (any Gmail account you control -- it can be
   your personal Gmail, it just needs to be able to send email; the report goes to your
   college address regardless of which account sends it):
   - Go to your Google Account -> Security -> turn on **2-Step Verification** if it isn't already on.
   - Go to https://myaccount.google.com/apppasswords, create an app password (name it
     "LeetCode Report"), copy the 16-character code it gives you.

3. **Add two repo secrets** (Settings -> Secrets and variables -> Actions -> New repository secret):
   - `SENDER_EMAIL` = the Gmail address you're sending from
   - `SENDER_APP_PASSWORD` = the 16-character app password from step 2

4. **Run the test first** (Actions tab -> "Daily LeetCode Progress Report" -> "Run workflow" ->
   set `test_run` to `true` -> Run). This writes `reports/LeetCode_Daily_Report_TEST.xlsx`
   and commits it back to the repo -- no email is sent in test mode. Open that file in the repo
   and check:
   - all 48 students appear
   - the Notes sheet for anyone flagged N/A
   - a few "Overall Problems Solved" numbers against their real profiles, to sanity-check

5. Once that looks right, either wait for the 9:00 AM IST daily run, or trigger the workflow
   again with `test_run` left as `false` to get a real emailed report immediately.

## Running locally instead (optional)

```
pip install -r requirements.txt
python main.py --test          # writes reports/LeetCode_Daily_Report_TEST.xlsx, no email
python main.py                 # real run -- needs SENDER_EMAIL / SENDER_APP_PASSWORD env vars set
```

## Files

- `leetcode_client.py` -- talks to LeetCode's GraphQL endpoint, retries on rate-limits
- `main.py` -- reads the master Excel file, orchestrates the run, builds the report, emails it
- `data/history.csv` -- append-only log of every student's daily/overall numbers, for trend analysis
- `reports/` -- one dated `.xlsx` per day, never overwritten
