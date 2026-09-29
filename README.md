# NIFTY H1 / AVG_H1 + Option Score -> Google Sheets

`nifty_to_sheets.py` runs every 5 minutes (GitHub Actions cron, or a loop on your PC),
computes H1 breadth + option score, logs both to CSV in `data/`, and writes today's data
into your Google Sheet through a tiny Apps Script web app. No HTML is generated.

Tabs (created automatically, one pair per trading day):
- `H1_<date>`           : time | H1 | AVG_H1
- `OptionScore_<date>`  : Strike price x HH:MM grid (green positive / red negative)

## One-time setup (about 5 minutes)
1. Open your Google Sheet -> Extensions -> Apps Script.
2. Delete the sample code, paste all of `apps_script.gs`, and change `TOKEN` to your own password.
3. Save -> Deploy -> New deployment -> gear icon -> Web app:
   Execute as: Me | Who has access: Anyone -> Deploy -> Authorize (Advanced -> Go to project -> Allow).
4. Copy the Web app URL (ends in /exec).

## GitHub Actions
Repo -> Settings -> Secrets and variables -> Actions -> New repository secret:
- `SHEET_WEBHOOK_URL` = the Web app URL
- `SHEET_TOKEN`       = the same password you put in the script

Settings -> Actions -> General -> Workflow permissions -> Read and write.
Then Actions tab -> Run workflow once to test.

## Local PC
    pip install -r requirements.txt
    PowerShell:  $env:SHEET_WEBHOOK_URL="https://script.google.com/macros/s/.../exec"
                 $env:SHEET_TOKEN="your password"
    python nifty_to_sheets.py            # loops every 5 minutes
    python nifty_to_sheets.py --once     # single cycle
    python nifty_to_sheets.py --once --no-sheets   # CSV only

## If you change apps_script.gs later
Deploy -> Manage deployments -> edit (pencil) -> Version: New version -> Deploy.
(The URL stays the same.)
