# Nasie Live Report — Developer Setup Guide

**Goal:** When the maintenance Excel file is updated, a styled live HTML dashboard
auto-rebuilds and publishes to a stable URL for management.

```
SharePoint (Excel)  ──Power Automate──►  GitHub repo: data/Nasie_Maintenance.xlsx
                                                │ (commit/push)
                                                ▼
                                     GitHub Actions (.github/workflows/build.yml)
                                                │ runs build_report.py
                                                ▼
                                     latest/index.html ──► GitHub Pages (live link)
```

Stable management link: `https://<user>.github.io/nasie-report/latest/`

---

## 1. Repository
1. Create or use a GitHub repo named **`nasie-report`** (Public, or Private + Pages if you have GitHub Pro/Enterprise).
2. Upload **all** files in this package, preserving structure:
   ```
   build_report.py
   logo.b64
   requirements.txt
   data/Nasie_Maintenance.xlsx
   .github/workflows/build.yml
   ```

## 2. GitHub Pages
- Repo **Settings → Pages**.
- Recommended: **Source = "GitHub Actions"**. (The classic "Deploy from a branch / root" also works since the workflow commits HTML into the repo; if you use branch mode, point it at `main` and serve `/latest` via the link above.)
- Confirm the published URL.

## 3. The build (already automated)
`.github/workflows/build.yml` triggers on:
- push to `data/Nasie_Maintenance.xlsx` or `build_report.py`
- manual run (Actions tab → Run workflow)
- (optional) hourly cron — uncomment the `schedule` block

It runs:
```bash
pip install -r requirements.txt
python build_report.py data/Nasie_Maintenance.xlsx latest
# + copies a dated snapshot to archive/<YYYY-MM-DD>/
```
then commits `latest/` and `archive/` back to the repo (Pages redeploys automatically).

## 4. SharePoint → GitHub sync (Power Automate)
Create a cloud flow:
1. **Trigger:** "When a file is modified (properties only)" on the SharePoint library holding `Nasie_Maintenance.xlsx`.
2. **Get file content** (SharePoint).
3. **HTTP** action → GitHub Contents API to update the file:
   - `PUT https://api.github.com/repos/<owner>/nasie-report/contents/data/Nasie_Maintenance.xlsx`
   - Headers: `Authorization: Bearer <GITHUB_PAT>`, `Accept: application/vnd.github+json`
   - Body:
     ```json
     {
       "message": "Update maintenance data from SharePoint",
       "content": "<base64 of file content>",
       "sha": "<current file sha, fetched via GET first>"
     }
     ```
   - Use a fine-grained PAT with `Contents: Read and write` on `nasie-report` only.
   *(Alternative if you prefer no PAT in Power Automate: trigger `workflow_dispatch` via the Actions API and commit the xlsx by other means, or use a GitHub App.)*

That's it — editing the Excel propagates to the live link within ~1–2 minutes.

---

## Data contract (important)
`build_report.py → load_data()` expects two sheets:
- **`Service Calls`**: columns (0-based) → 0:id, 2:listing, 3:description, 4:status, 6:contractor, 7:scheduled date, 9:material cost, 10:labor cost.
- **`Sheet1`** (enriched): 1:cat2(title), 2:cat3(classification), 3:contractor, 4:listing, 6:city, 7:area, 9:internal labor, 10:external labor, 11:external flag, 12:admin fee, 13:material, 14:total, 16:BOOM task URL (contains task id), 17:raiser.

Business rules already implemented:
- Count only `Done` and `In Review` (= Pending Approval) statuses.
- Exclude rows matching: Demo Listing, General Requests - Internal, Inquiry/استفسار, Guest Experience.
- Internal/external split from Sheet1; a request with external>0 counts as "with external".
- Materials use Material Cost only.
- Empty raiser → "AI". Category consolidation (electricity/AC/etc.) via `map_cat`.
- Weekly comparison auto-rolls via `latest/prev_summary.json`.

If SharePoint column order changes, adjust the index numbers in `load_data()` only.

## Local test
```bash
pip install -r requirements.txt
python build_report.py data/Nasie_Maintenance.xlsx latest
open latest/index.html
```

## Optional: PDF export
The HTML is print-ready (A4, each section on its own page). To auto-generate a PDF too,
add a step using Playwright/Chromium:
```bash
pip install playwright && playwright install chromium
# then render latest/index.html with emulate_media('print') → latest/report.pdf
```
