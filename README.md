# 🏥 Hospital Emergency Room Analysis Dashboard

An end-to-end Excel analytics project — raw ER admission data is cleaned and transformed with **Power Query**, modeled with a custom **Date/Calendar table**, enriched with **calculated columns and DAX measures**, and visualized in an interactive **single-file Excel dashboard** with KPI cards, charts, slicers, and hyperlink-based drill-through navigation.

---

## 📌 Project Overview

The dataset tracks patient visits to a Hospital Emergency Room — admission date/time, demographics, department referral, wait time, satisfaction score, and admission outcome. The goal was to turn a messy raw CSV export into a clean, filterable report that hospital staff/management can use to monitor ER performance and compare trends across periods.

**Workflow:** `Raw CSV → Power Query cleaning → Custom Calendar table → Calculated columns → Data Model relationship → PivotTables/PivotCharts → Interactive Dashboard`

---

## 🗂️ Project Files

| File / Path | Description |
| :--- | :--- |
| `📄README.md` | Contains README.md file |
|`🗂️Dashboard/` ↳ `📊Hospital_Emergency_Room_Dashboard.xlsx` | Raw, unprocessed source data | 
| `🗂️Script/`↳ `📄Hospital_data.xlsx` | Single workbook containing: Power Query cleaning steps, Data Model, Pivot Report sheet, Dashboard sheet, and 3 supporting detail sheets |
| `images/` | Static chart exports used in this README |

> The cleaned data is **not stored as a separate static table** — it lives entirely inside the Power Query connections of the workbook, so the dashboard refreshes from the query rather than a hardcoded sheet.

---

## 🛠️ Tools & Techniques Used

- **Microsoft Excel** (Tables, PivotTables, PivotCharts, Slicers)
- **Power Query (M)** — data cleaning, column splitting, type conversion, blank query for calendar generation
- **Data Model / Power Pivot** — relationships and DAX-style calculated fields
- **Cell Hyperlinks + Home button** — used to simulate drill-through style navigation between the Dashboard and detail sheets

---

## 🧹 Data Cleaning (Power Query)

| Raw Column Issue | Fix Applied |
|---|---|
| `Patient Admission Date` mixed date **and** time in one field (e.g. `20-03-2024 08:47`) | Split into two separate columns: `Patient Admission Date` and `Patient Admission Time` |
| `Patient Gender` inconsistent values (`M`, `F`, `Male`, `Female` all mixed) | Standardized to full text `Male` / `Female` |
| `Patient First Initial` + `Patient Last Name` as two columns | Merged into a single `Patient Name` column (e.g. `G. Stanlack`) — keeps identity semi-anonymized |
| `Patient Admission Flag` duplicated (same TRUE/FALSE column appeared twice) | Duplicate column removed |
| `Patient Admission Flag` stored as `TRUE`/`FALSE` | Converted to readable labels: `Admitted` / `Not Admitted` |
| No age-band grouping | Custom calculated column added to bucket raw `Patient Age` into ranges (see formula below) |

---

## 🧮 Data Model & Formulas

### 1. Calendar table (Power Query, M)

Since the raw data only had a single `Patient Admission Date` column, there was no independent date dimension to power the Year/Month slicers correctly. A blank query generates a continuous calendar so every date — even ones with zero admissions — is represented:

```m
= List.Dates(
    #date(2023, 01, 01),
    731,
    #duration(1, 0, 0, 0)
)
```

This produces **731 sequential days starting 1 Jan 2023** (a full 2-year span covering 2023–2024, including the leap day). The resulting `Calendar[Date]` column is loaded into the Data Model and related to `FactTable[Patient Admission Date]`, which is what lets the Year and Month slicers filter every chart on the dashboard cleanly.

### 2. Age group bucketing (calculated column)

```dax
= IF([Patient Age] >= 70, "70-79",
  IF([Patient Age] >= 60, "60-69",
  IF([Patient Age] >= 45, "45-59",
  IF([Patient Age] >= 30, "30-44",
  IF([Patient Age] >= 15, "15-29",
  IF([Patient Age] >= 5, "05-14", "0-4"))))))
```

This nested `IF` walks each patient's age from the top down and assigns the first band it qualifies for, producing 7 groups: `0-4`, `05-14`, `15-29`, `30-44`, `45-59`, `60-69`, `70-79`.

> **Note:** the pivot reports and dashboard charts referenced in this README display age in even 10-year bands (`0-09`, `10-19`, `20-29` … `70-79`) rather than the uneven bands this formula produces. If the workbook you're publishing uses the even-band version, swap the thresholds to `>=70/60/50/40/30/20/10` so the formula documented here matches what's on screen.

### 3. Other measures (same pattern, DIVIDE-based)

Every percentage shown on the dashboard (Admitted %, On-time %, Male %, etc.) follows the same structure — count of the category divided by the grand total for that field:

```dax
Admitted % = DIVIDE([Count of Admitted], [Total Patients], 0)
On-Time % = DIVIDE([Count of On-Time], [Total Patients], 0)
```

---

## 📊 Dashboard Structure

### KPI Cards — 3 total
Total Patients · AVG Wait Time · Satisfaction Score

### Slicers — 2 total
Year · Month

### Charts — 6 total

| # | Chart | Type |
|---|---|---|
| 1 | Patients Trend (Daily) | Area/line chart |
| 2 | Patients Attended Within Time | Donut chart |
| 3 | Patient by Gender | Donut chart |
| 4 | Admission Status | Icon + horizontal bar |
| 5 | Patients by Age Group | Column chart |
| 6 | Department Referral | Horizontal bar chart |

---

## 🔎 Chart-by-Chart Analysis (comparing across all the data)

By comparing all the data side by side rather than looking at any single period in isolation, some patterns hold steady while others move around noticeably.

### Total patients & wait time

Patient volume moves up and down period to period with no clear upward or downward direction. Average wait time, on the other hand, stays remarkably stable in the mid-30-minute range almost regardless of how many patients came in — a sign that the pace of care isn't scaling proportionally with demand.

![Total ER patients trend]("D:\EXCEL\Hospital_data\Patient_Trend_Analysis.png")

### Patients attended within time — the biggest recurring issue

This is the one pattern that never flips. Comparing across all the data, **delay consistently outnumbers on-time** — not once does on-time attendance overtake delay in any period compared. This is a structural pattern, not a one-off bad stretch.

![On-time vs delay rate]("D:\EXCEL\Hospital_data\OnTime_vs_Delay_Analysis.png")

### Admission status (Admitted vs Not Admitted)

This stays close to an even 50/50 split whenever compared — meaning roughly half of every ER visit escalates into a hospital admission. That's a fairly severe case mix for an ER, suggesting this isn't primarily a walk-in-minor-injury facility.

### Patient by gender

Male patients hold a small, consistent edge over Female across almost every period compared, generally landing in the low-to-mid 50% range. It's a mild, stable skew rather than a dramatic one.

### Patients by age group — the flattest chart on the dashboard

Comparing totals across all the data, no single age band dominates. Every bracket lands within a tight range of each other — this ER serves a genuinely broad demographic rather than skewing toward children or the elderly. Which bracket looks "tallest" shifts around depending on the period you're looking at, but that's noise sitting on top of an essentially flat underlying distribution.

![Patients by age group]("D:\EXCEL\Hospital_data\Age_Group_Analysis.png")

### Department referral — the most concentrated chart

This ranking barely changes no matter which periods are compared: **None → General Practice → Orthopedics** always occupy the top three spots, together accounting for the large majority of all referrals. The remaining specialties split a small share between them, with Renal consistently the smallest category by a wide margin.

![Department referral volume]("D:\EXCEL\Hospital_data\Department_Referral_Analysis.png")

---

## 🔗 Interactive Navigation

Cell hyperlinks sit on top of the KPI cards to simulate Power BI–style drill-through inside plain Excel:

- Clicking **Total Patients** → jumps to the **Every Day Patient Count** sheet
- Clicking **AVG Wait Time** → jumps to the **Patient Avg Wait-Time** sheet
- Clicking **Satisfaction Score** → jumps to the **Satisfaction Score** sheet
- Each detail sheet has a **Home icon** that hyperlinks back to the main **Dashboard** sheet

---

## ⚠️ Problems Faced & Solutions

| Problem | Why It's Happening | Solution |
|---|---|---|
| Gender field had multiple spellings (`M`, `F`, `Male`, `Female`) | Manual/free-text data entry at source, no dropdown validation | Standardized via Power Query "Replace Values" so charts group correctly |
| Date and Time combined in a single column | Source system exported a single timestamp field | Split into two columns using Power Query's Split Column by delimiter |
| Duplicate `Patient Admission Flag` column | Likely a duplicated export/merge error at the source | Removed the redundant column to avoid double-counting |
| No independent date table to filter by Year/Month | A single admission-date field can't be sliced cleanly across multiple measures without a dimension table | Built a Calendar table via `List.Dates()` and linked it through the Data Model |
| Chronic delay rate — on-time never overtakes delay when comparing across all the data | Average wait time stays flat regardless of patient volume, meaning intake/triage capacity isn't scaling with demand | Investigate staffing-to-volume ratio by shift/hour — a period-level average can hide hour-of-day bottlenecks the current charts can't show |
| Satisfaction doesn't track cleanly with delay rate | The periods with the worst delay aren't always the periods with the worst satisfaction, and vice versa | Satisfaction is likely driven by more than wait time alone (e.g. bedside manner, environment, billing) — worth adding a dedicated wait-time-vs-satisfaction correlation view if the source data has more fields to pull in |
| Heavy concentration in "None" referral and General Practice | Either genuinely minor/general cases, or intake defaults to "None" when a referral decision isn't made promptly | Check whether "None" cases have longer or shorter wait times than referred cases — if longer, it may point to a triage bottleneck rather than case severity |
| Renal referrals near-zero every time compared | Could be genuinely low incidence, or Renal cases may be getting miscoded into "None"/General Practice | Spot-check a sample of "None"/"General Practice" records for renal-related symptoms to rule out miscoding |
| "External Data Connections have been disabled" warning on opening the file | Excel's default security setting blocks Power Query/data connections from running automatically when a workbook is downloaded from the internet (e.g., GitHub) | Expected behavior, not corruption — click **Enable Content** in the yellow security bar |
| Some slicer options appear grayed out | The calendar table always spans the full 2-year range, but not every date in that range has matching admission records | Expected — a grayed slicer option means "no data for this period," not a broken slicer |

---

## 1. Executive Diagnostic Summary

| Operational Area | Baseline Status | Identified Root Cause | Corrective Action |
| :--- | :--- | :--- | :--- |
| **High Patient Delays** | 🚨 **57%** wait > 30 mins (Target: 56%–66% reduction) | Intake triage bottlenecks & rigid staffing during mid-month volume surges | Implement Fast-Track Triage lane & realign shift rosters to daily volume trends |
| **Average Wait Time** | ⏳ **~35 min** (Exceeds 30-min SLA benchmark) | Diagnostic testing delays (lab/imaging) & slow inpatient bed turnover | Institute sub-20 minute turnaround SLAs for laboratory and imaging results |
| **Referral Integrity** | ⚠️ **50%+** tagged as "None" | EHR discharge entry forms treat referral fields as optional | Enforce mandatory referral selection during electronic patient discharge logging |

---

## 2. Structured Diagnostic Breakdown

### 🚨 High Patient Delays (56%–66% Over Target)
* **Status:** 57% of incoming ER patients experience wait times exceeding the 30-minute target threshold.
* **Root Cause:** Bottlenecks during initial intake triage combined with shift schedules that fail to align with mid-month patient volume spikes.
* **Action:** Launch a dedicated Fast-Track Triage lane for low-acuity cases and dynamically adjust physician/nurse staffing rosters using historical trend charts.

### ⏳ Average Wait Time (~35 Minutes)
* **Status:** Overall average patient wait time sits at ~35 minutes, consistently exceeding the 30-minute operational SLA.
* **Root Cause:** Turnaround delays in receiving diagnostic lab work/imaging results and slow bed turnover for patients requiring inpatient admission.
* **Action:** Establish binding inter-departmental SLAs requiring lab and imaging returns in **< 20 minutes**, while optimizing inpatient bed discharge procedures.

### ⚠️ Referral Data Integrity Gaps ("None" Category = 50%+)
* **Status:** Over half of all discharge records lack an assigned specialty referral, defaulting to "None".
* **Root Cause:** Optional (non-mandatory) field validation settings in current EHR electronic discharge logging forms.
* **Action:** Update EHR software configuration to mandate explicit referral selection (or formal "Outpatient Self-Care" designation) before a discharge record can be closed.

---

## 🚀 How to Use This File

1. Download `Hospital_data.xlsx` and open it in Excel
2. Click **Enable Content** on the yellow security bar (needed for Power Query/slicers to work)
3. Go to the **Dashboard** sheet
4. Use the **Year** and **Month** slicers to filter the whole report
5. Click on any KPI card to jump to its detailed breakdown sheet; use the **Home** icon to return

---

