# LinkedIn Job Scraper (Advanced & Robust Filtering Edition)

An automated LinkedIn job search tool built with Selenium and Microsoft Edge. It searches for CAE, FEA, FEM, CFD, and simulation related roles (thesis, internship, or full time) across Germany and other European regions, filters results with keyword scoring and negative keyword exclusions, and exports matched jobs to CSV and Excel in real time. It runs with a dedicated stealth browser profile so your normal Edge profile stays untouched.

## Features

- **Granular URL filters**: date posted, sort order, experience level, job type, workplace type, and exact geoIds for Germany, DACH, Nordics, and other European regions.
- **AI natural language search**: uses LinkedIn's semantic AI search (with URL parameter, prompt page, and search bar fallback strategies) so full sentence queries are matched instead of single keywords.
- **Multi stage text filtering**: expands truncated job descriptions ("See more" / "Mehr anzeigen"), applies title and company blacklists, and uses smart German and English compound keyword matching (for example FEM Berechnung, Crashsimulation).
- **Context aware keyword validation**: ambiguous acronyms such as META, ANSA, and CFD only count as matches when they appear near relevant domain context (for example solver, meshing, fluid, flow).
- **Relevance scoring**: each matched job gets a numeric score based on keyword count and whether keywords appear in the title.
- **Automatic export**: deduplicated, real time export to both CSV and a formatted Excel workbook with frozen header row, auto filter, and auto sized columns.
- **Resilient scraping**: multi layer fallback CSS selectors handle LinkedIn's frequently changing SDUI layout, plus stealth timing, mouse jitter, and automatic session recovery after crashes or connection errors.
- **Optional continuous mode**: loop the whole search on a timer instead of running once.

## Requirements

- Windows with Microsoft Edge installed
- Python 3.10 or later (uses modern type hint syntax such as `str | None`)
- Install dependencies:

```bash
pip install selenium webdriver-manager pyperclip pandas openpyxl
```

The script auto detects your installed Edge version and resolves a matching `msedgedriver.exe` (from the system, WinGet cache, Downloads, or by downloading it). No manual driver setup is required.

## Configuration

All settings live near the top of `linkedin_job_scraper_pwa.py`, under **Section 1: User Configuration**. Key options:

| Setting | Purpose |
|---|---|
| `LINKEDIN_EMAIL` / `LINKEDIN_PASSWORD` | Optional auto login. Leave blank to log in manually on first run. |
| `LOCATION_FILTER` | Country, region, or list of countries to search (for example `"Germany"`, `"DACH"`, or `["Germany", "Austria"]`). |
| `DATE_FILTER` | `"today"`, `"week"`, `"month"`, or `"any"`. |
| `SORT_BY` | `"recent"` or `"relevance"`. |
| `USE_AI_SEARCH` | `True` for natural language AI search queries, `False` for traditional keyword URL search. |
| `AI_SEARCH_QUERIES` / `PRIMARY_KEYWORDS` | The search queries used depending on the mode above. |
| `EXPERIENCE_LEVEL`, `JOB_TYPES`, `WORKPLACE_TYPES` | Lists that narrow results by seniority, employment type, and remote or on site status. |
| `MIN_SECONDARY_MATCHES` | Minimum number of secondary keywords a job description must contain to count as a match. |
| `SAVE_ON_LINKEDIN` | Automatically clicks "Save" on matched jobs in your LinkedIn account. |
| `CONTINUOUS_MODE`, `LOOP_WAIT_MINUTES`, `MAX_CYCLES` | Controls for running the scraper on a repeating schedule. |
| `CSV_OUTPUT_FILE`, `EXCEL_OUTPUT_FILE` | Output file names and paths. |

Keyword lists (`PRIMARY_KEYWORDS`, `SECONDARY_KEYWORDS`), title and line exclusion patterns, and the company blacklist are defined in **Section 2** and can be edited to match your target roles.

## Usage

1. Adjust the configuration values described above to match your search (location, keywords, filters).
2. Run the script:

```bash
python linkedin_job_scraper_pwa.py
```

3. On first run, log in manually in the Edge window that opens (the script waits up to 5 minutes). Your session is saved in a local stealth profile folder (`.edge_scraper_profile`), so subsequent runs stay logged in.
4. The scraper searches each configured keyword or AI query across the selected location(s), opens each job card, evaluates the description against your keyword rules, and writes matches to the CSV and Excel output files as they are found.
5. Press `Ctrl+C` at any time to stop gracefully. A summary of searched, matched, saved, and error counts is printed.

## Output

Each matched job record includes:

- Job ID and URL
- Title, company, location, and workplace type
- Posted date
- Match score and list of matched keywords
- The specific description lines that triggered the match
- The search query that found it
- A timestamp of when it was scraped

Results are deduplicated by Job ID and written to both `linkedin_matched_jobs.csv` and `linkedin_matched_jobs.xlsx` (file names configurable).

## Notes and limitations

- This script automates a real LinkedIn browser session. Use reasonable search volumes and pauses (already built in) to reduce the risk of rate limiting or account flags.
- LinkedIn's page structure changes frequently. If card or field detection stops working, check the CSS selector lists in the script (Sections 11 to 13) and update them.
- The AI search feature depends on LinkedIn surfaces that may be gated or A/B tested; the script automatically falls back to traditional keyword URL search if AI search is unavailable.
- A daily log file (`linkedin_scraper_YYYY-MM-DD.log`) is created next to the script for troubleshooting.

## Disclaimer

This tool is intended for personal job search automation. Review LinkedIn's Terms of Service before running automated browsing tools against your account, and use at your own risk.
