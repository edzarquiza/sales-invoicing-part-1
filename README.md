# Dexterous Invoicing Automation

A Streamlit web app that automates the generation of per-CM invoicing workbooks from XPM data — replacing the manual ChatGPT workflow.

---

## What it does

1. **Upload** your master `.xlsm` / `.xlsx` workbook (requires `XPM Data` and `CM assignment` sheets)
2. **Set three date ranges** — Main, Weekly, Monthly
3. **Click Run** — the app filters, pivots, and generates one formatted `.xlsx` workbook per Client Manager
4. **Download** individually or as a single ZIP with CM-named folders

### Output format (matches your existing ChatGPT process)
- One sheet per client (skipped if total hours = 0)
- Pivot-style layout with blanked repeated values
- Light blue header & grand total rows (bold)
- Yellow rows for non-billable entries
- Bold Row Labels and Staff Name columns
- Auto-fitted column widths (except [Time] Note)
- `Unassigned Clients.xlsx` for any clients not in CM assignment

---

## Local setup

```bash
# 1. Clone this repo
git clone https://github.com/YOUR_USERNAME/dexterous-invoicing.git
cd dexterous-invoicing

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
streamlit run app.py
```

---

## Deploy to Streamlit Community Cloud (free, always-on)

> This lets anyone on your team open the app in a browser — no Python required.

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/dexterous-invoicing.git
git push -u origin main
```

### Step 2 — Deploy on Streamlit Cloud
1. Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub
2. Click **"New app"**
3. Select your repository → branch `main` → main file `app.py`
4. Click **"Deploy"** — you'll get a public URL like `https://your-app.streamlit.app`

That's it. Share the URL with your team. No installation needed.

---

## File structure

```
dexterous-invoicing/
├── app.py              ← Main Streamlit application
├── requirements.txt    ← Python dependencies
└── README.md           ← This file
```

---

## Three-range filter logic

| Batch value | Date range applied |
|---|---|
| `"Weekly"` | Weekly Date Range |
| `"Monthly"` | Monthly Date Range |
| Anything else (incl. blank) | Main Date Range |

Both billable (`Yes`) and non-billable (`No`) rows are included in the output.
