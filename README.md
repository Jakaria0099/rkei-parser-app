# RKEI Form Processor — User Documentation

---

## 📌 What Is This?

The **RKEI Form Processor** is a web-based tool that reads your completed **RKEI Word (.docx) forms** and automatically extracts all the key data into a single, clean **Excel spreadsheet** called `SPRE_CodedReturns_REF2029.xlsx`.

Instead of manually opening dozens of Word forms, copying data into spreadsheets, and cross-referencing codes — this tool does it all for you in seconds.

---

## 🤔 Why Does This Exist?

RKEI forms contain structured data across multiple tables — staff details, research priorities, bids, events, engagement activities, and impact entries — each tagged with specific codes (SPRE, Stage, Partner type).

Manually consolidating this information from many forms is:

- **Time-consuming** — each form has multiple tables with dropdowns and codes.
- **Error-prone** — easy to misread or skip entries.
- **Hard to summarise** — turning raw form data into useful summaries and pivot tables takes extra work.

This tool **automates the entire process**. Upload your `.docx` files, click a button, and get a fully structured Excel workbook with summaries, breakdowns, pivot tables, and charts — ready for analysis or reporting.

---

## ⚙️ How Does It Work?

Behind the scenes the tool:

1. **Reads** each uploaded `.docx` file by inspecting the underlying XML (the same data Word uses internally).
2. **Locates** the specific tables in the RKEI template (Staff, Priorities, Bids, Events, Engagement, Impact).
3. **Extracts** all dropdown selections, free-text entries, and codes from each table row.
4. **Normalises** the data — for example, detecting and separating dates that may appear in unexpected positions in the staff row.
5. **Aggregates** everything into multiple Excel sheets with summaries, pivot tables, and charts.
6. **Packages** it all into a single downloadable `.xlsx` file.

---

## 🚀 How To Use It

### Step 1 — Open the App

Visit the app URL in your web browser. No installation or login is required.

### Step 2 — Upload Your Forms

Click **Browse files** (or drag and drop) to select one or more `.docx` RKEI forms. You can upload as many files as you like at once.

### Step 3 — Process

Click the **▶ Process** button. The app will read all uploaded forms and generate the Excel output. This usually takes just a few seconds.

### Step 4 — Download

Once processing is complete, click **⬇ Download Excel**. The file will be saved to your computer as:

```
SPRE_CodedReturns_REF2029.xlsx
```

That's it! Open the file in Excel to view your data.

---

## 📊 What's In The Excel File?

The downloaded spreadsheet contains **9 tabs (sheets)**. Here is what each one contains and what every column means:

---

### 1. `summary_counts`

**What it is:** A high-level count of every code that appeared across all uploaded forms, grouped by code family.

| Column | Meaning |
|--------|---------|
| `family` | The code family — one of `SPRE`, `STAGE`, or `PARTNER` |
| `code` | The specific three-letter code (e.g. `STR`, `PLN`, `ACD`) |
| `count` | How many times this code appeared across all forms |
| `meaning` | The human-readable description of the code |
| `percent` | The percentage this code represents within its family |

**Code reference:**

| Code | Family | Meaning |
|------|--------|---------|
| STR | SPRE | Strategic research output |
| PPL | SPRE | People / research culture |
| IIF | SPRE | Income / funding |
| CEI | SPRE | Civic / external impact |
| PLN | STAGE | Planning stage |
| DEV | STAGE | In development |
| EXT | STAGE | Externally submitted |
| LIV | STAGE | Live activity |
| CMP | STAGE | Completed |
| EVD | STAGE | Evidence of impact |
| ACD | PARTNER | Academic partner |
| IND | PARTNER | Industry partner |
| CUL | PARTNER | Cultural organisation |
| COM | PARTNER | Community group |
| PUB | PARTNER | Public sector |
| PRO | PARTNER | Professional body |
| NON | PARTNER | Non-profit |

---

### 2. `counts_by_file`

**What it is:** Shows how many data entries were extracted from each uploaded form.

| Column | Meaning |
|--------|---------|
| `file_name` | The name of the `.docx` file |
| `entries` | Total number of data rows extracted from that file |
|
Useful for quickly spotting forms that may be incomplete (very low count) or unusually large.

---

### 3. `distinct_staff_counts`

**What it is:** A list of every unique staff member found across all forms and how many entries are associated with them.

| Column | Meaning |
|--------|---------|
| `staff_name` | The staff member's name as it appeared in the form |
| `entries` | Total number of data rows linked to that person |

---

### 4. `codes_by_uoa`

**What it is:** A breakdown of code usage grouped by **Unit of Assessment (UoA)**.

| Column | Meaning |
|--------|---------|
| `uoa` | The Unit of Assessment value from the staff section of the form |
| `family` | Code family (`SPRE`, `STAGE`, or `PARTNER`) |
| `code` | The specific three-letter code |
| `count` | How many times this code appeared for this UoA |

---

### 5. `codes_by_pathway`

**What it is:** A breakdown of code usage grouped by **Pathway**.

| Column | Meaning |
|--------|---------|
| `pathway` | The Pathway value from the staff section of the form |
| `family` | Code family (`SPRE`, `STAGE`, or `PARTNER`) |
| `code` | The specific three-letter code |
| `count` | How many times this code appeared for this Pathway |

---

### 6. `master_rows`

**What it is:** The **complete raw dataset** — every single entry extracted from every form, one row per item. This is the most detailed sheet and the foundation for all other summaries.

| Column | Meaning |
|--------|---------|
| `file_name` | Which `.docx` file this row came from |
| `staff_name` | Name of the staff member (from the staff table in the form) |
| `position` | Staff member's position/role |
| `department` | Staff member's department |
| `pathway` | Research pathway (dropdown value from the form) |
| `uoa` | Unit of Assessment (dropdown value from the form) |
| `review_date` | The review date found in the staff row (if present) |
| `section` | Which section of the form this entry is from — one of: `Priorities`, `Bids`, `Events`, `Engagement`, or `Impact` |
| `row_id` | The row identifier within the section (e.g. 1, 2, 3) |
| `entry` | The free-text content of the entry |
| `spre_code` | SPRE code (Priorities section only) — `STR`, `PPL`, `IIF`, or `CEI` |
| `baseline` | Baseline stage code (Priorities section only) |
| `target` | Target stage code (Priorities section only) |
| `stage` | Stage code (Bids and Impact sections) |
| `partner` | Partner type code (Events and Engagement sections) |

> **Note:** Not every column applies to every row. For example, `spre_code` only has values for Priorities rows, and `partner` only has values for Events/Engagement rows. Empty cells are normal and expected.

---

### 7. `file_level_extraction`

**What it is:** A diagnostic view showing one row per uploaded file, with the Pathway and UoA that were extracted from it.

| Column | Meaning |
|--------|---------|
| `file_name` | The name of the `.docx` file |
| `pathway` | The Pathway value extracted from the staff table |
| `uoa` | The UoA value extracted from the staff table |
| `pathway_missing` | `True` if the Pathway could not be extracted (blank in the form) |
| `uoa_missing` | `True` if the UoA could not be extracted (blank in the form) |

Use this sheet to **validate** that each form was read correctly. If you see `True` in the missing columns, it means that form may not have had a Pathway or UoA selected in its dropdown.

---

### 8. `pivot_pathway_uoa`

**What it is:** A **cross-tabulation (pivot table)** showing the number of unique files for each combination of Pathway and UoA.

- **Rows** = Pathway values
- **Columns** = UoA values
- **Cells** = Number of unique files with that Pathway + UoA combination
- **Total_by_pathway** (rightmost column) = Row totals
- **Total_by_uoa** (bottom row) = Column totals

This gives you a quick bird's-eye view of how forms are distributed across Pathways and UoAs.

---

### 9. `charts`

**What it is:** Visual bar charts showing the distribution of codes across three families:

- **SPRE distribution** — how many times each SPRE code (STR, PPL, IIF, CEI) appeared
- **STAGE distribution** — how many times each Stage code (PLN, DEV, EXT, LIV, CMP, EVD) appeared
- **PARTNER distribution** — how many times each Partner code (ACD, IND, CUL, COM, PUB, PRO, NON) appeared

These charts are embedded as images directly in the Excel file for easy inclusion in reports or presentations.

---

## 📝 Requesting Changes, Updates, or New Features

If you need any changes to the tool — whether it's a bug fix, a new column, a different chart, support for a new form template, or any other enhancement — please provide the following information to the developer/owner:

### What to include in your request:

1. **What you want changed**
   - Be as specific as possible. For example: *"Add a new column called X to the master_rows sheet"* or *"The app isn't reading the Partner dropdown from the Events table correctly."

2. **Why you need it**
   - A brief explanation of the purpose helps the developer make the right design decision. For example: *"We need to report on this for the REF submission deadline."

3. **An example file (if relevant)**
   - If the issue is related to a specific form not being parsed correctly, attach the `.docx` file that is causing the problem. This is the fastest way to get a fix.

4. **Expected vs. actual behaviour**
   - What did you expect to see in the output? What did you actually see? Screenshots of the problematic cells in the Excel output are very helpful.

5. **Priority / timeline**
   - Is this urgent (blocking your work) or a nice-to-have for the future?

### How to send your request:

- **Email the developer directly** with the above details and any attachments.
- Or, if you have a GitHub account, **open an Issue** on the project repository at:
  **https://github.com/Jakaria0099/rkei-parser-app/issues**
  - Click **New Issue**, give it a clear title, and paste your details.

> **Important:** If the RKEI Word form template itself changes (e.g. new tables are added, columns are reordered, or new dropdown codes are introduced), the tool will need to be updated to match. Please notify the developer as soon as you become aware of any template changes.

---

## 🔧 Technical Notes (For the Developer)

- **App framework:** Streamlit (deployed on Streamlit Community Cloud)
- **Source repository:** [github.com/Jakaria0099/rkei-parser-app](https://github.com/Jakaria0099/rkei-parser-app)
- **Main module:** `app.py` (Streamlit front-end) imports `rkei_parser.py` (parsing logic)
- **Template dependency:** The parser relies on specific table indexes in the RKEI `.docx` template. If the template structure changes, update the `TABLE_IDX` dictionary in `rkei_parser.py`.
- **Auto-deploy:** Any push to the `main` branch automatically redeploys the live app.

---

*Last updated: March 2026*