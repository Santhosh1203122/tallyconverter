# Tally Sales Report Generator

A simple desktop app that turns a Tally ERP/Prime **DayBook XML** export
into a clean Excel sales report (one row per line item) — in the exact
column layout of your sales report template.

## What it does

1. Click **Browse** and pick the `DayBook.xml` exported from Tally.
2. Click **Convert to Excel**.
3. The Excel file is saved next to the XML, named
   `<input_name>_sales_report.xlsx` (e.g. `DayBook_sales_report.xlsx`).

The output sheet has these 14 columns:

```
Invoice No | Invoice Date | Client Name | Client GST No | Item | HSN Code |
Quantity   | Rate         | Amount (Excl. GST) | GST % |
CGST Amount | SGST Amount | IGST Amount | Total Invoice Value
```

Credit-note vouchers (sales returns) are emitted with negative amounts.

## Project files

| File | Purpose |
|------|---------|
| `converter.py` | Parsing + Excel-writing logic (no GUI) |
| `app.py` | Tkinter GUI front-end |
| `requirements.txt` | Python deps |
| `.github/workflows/build.yml` | **Auto-builds the Windows .exe in the cloud** |
| `build_windows.bat` | Optional: build .exe locally on a Windows machine |
| `build_macos.sh` | Optional: build .app + .dmg locally on a Mac |

---

## How to get the Windows .exe (no Windows machine needed)

You'll use **GitHub Actions** — a free service that builds your .exe on
real Windows servers in the cloud, then lets you download it. Total
time: ~10 minutes the first time.

### Step 1 — Push this folder to GitHub

If you don't have a GitHub account, create one at <https://github.com/signup>.

1. On GitHub, click **+** (top-right) → **New repository**.
2. Name it `tally-sales-report` (or anything). Set **Private** if you
   prefer. Click **Create repository**.
3. On the next page, click the **"uploading an existing file"** link.
4. Drag the **contents of this folder** (every file *and* the `.github`
   folder — make sure hidden files come along) into the upload area.
5. Click **Commit changes**.

### Step 2 — Wait for the build

1. Click the **Actions** tab at the top of the repo page.
2. You'll see a workflow run called **"Build Windows EXE"** with a
   yellow circle (running) or green check (done). Click it.
3. After about 3–4 minutes it turns green.

### Step 3 — Download the .exe

1. On that same workflow-run page, scroll to the bottom.
2. Under **Artifacts** you'll see **TallySalesReport-Windows**. Click it.
3. GitHub downloads a `.zip` containing `TallySalesReport.exe`.
4. Unzip → that .exe is your finished Windows app. Send it to anyone.

> **First-run note:** Windows SmartScreen may show a "Windows protected
> your PC" dialog because the .exe isn't code-signed. Click **More info**
> → **Run anyway**. This happens on any unsigned exe.

### Step 4 — Future updates

Whenever you change the code on GitHub, the workflow rebuilds the .exe
automatically. Just go to the **Actions** tab and download the new
artifact — same place as before.

You can also **manually trigger** a build:
1. Go to the **Actions** tab → click **Build Windows EXE** on the left.
2. Click the **Run workflow** button on the right → **Run workflow**.

---

## How to get a macOS .app/.dmg

Building macOS apps requires running on a Mac (Apple's tools won't run
on Windows or Linux). Two paths:

### Easy: build locally on your Mac

```bash
chmod +x build_macos.sh
./build_macos.sh
```

This produces `dist/TallySalesReport.app` and `dist/TallySalesReport.dmg`.

### Automated via GitHub Actions

If you also want macOS builds in the cloud, ask me to add a `macos-latest`
job to the workflow — it's a small addition.

---

## Run it directly from Python (no exe needed)

If you'd rather skip the .exe entirely and run from source:

```bash
python -m pip install -r requirements.txt
python app.py
```

You can also use the converter as a command-line tool:

```bash
python converter.py /path/to/DayBook.xml
# -> writes /path/to/DayBook_sales_report.xlsx
```

---

## Notes about your specific data

Your Tally exports use the `Sales GST (SYSTEM)` voucher class, which
**does not record per-item Quantity or Rate** — only the taxable
amount. The converter therefore leaves those two columns blank. All
the other fields (HSN, GST %, CGST/SGST/IGST, totals) are derived
correctly:

- **GST %** is parsed from the suffix of the stock-item name
  (e.g. `"3924  - BUCKET, PLASTIC GOODS - 18%"` → 18%).
- **CGST/SGST** apply when the place of supply matches the company's
  state (intra-state). **IGST** applies otherwise (inter-state).
- **Total Invoice Value** = Amount + CGST + SGST + IGST.

If a future Tally export starts including `<RATE>` and `<ACTUALQTY>`
per item, those columns will fill in too.
