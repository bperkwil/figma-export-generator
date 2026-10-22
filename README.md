# Figma Export Generator

A small web app that takes a campaign Excel workbook and turns it into a clean,
Figma-ready export. You upload a file, it does the tidying automatically, and
you download the result — no spreadsheet surgery required.

## What it does

1. **Upload** — drag in an `.xlsx` workbook.
2. **Auto-detect** — it figures out which tabs are Onsite, Outbound Email, and
   Outbound Push, even if the tab names aren't exact.
3. **Clean up** — it applies three sets of rules:
   - **Asset map** — assigns the right Figma asset size based on the placement.
   - **ILM auto-correction** — fixes common message-code typos (e.g. `ilm1` → `ILM-01`).
   - **Copy priority** — when there are several copy columns, it picks the best one
     (Final → Approved → Revised → Draft …).
4. **Summary** — shows stats, every correction it made, and any alerts.
5. **Download** — hands back one `.xlsx` with three sheets:
   - `Figma Export` (onsite)
   - `Figma Export - Outbound` (email + push)
   - `Processing Log` (a record of everything it detected, fixed, or flagged)

## How to run it

In a terminal, from this folder:

```
pip install -r requirements.txt
streamlit run app.py
```

Your browser opens automatically. Upload a file and go.

## Changing the rules

The three rule sets live at the very top of `app.py`, in plainly-labelled
sections (`ONSITE_ASSET_MAP`, `ILM_CORRECTIONS`, `COPY_PRIORITY`). They're just
simple lists — edit the words and sizes there to match your real business rules.
I filled them with sensible placeholders so the app runs end-to-end today; tell
me your actual values and I'll drop them in.
