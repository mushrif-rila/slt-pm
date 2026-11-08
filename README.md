# SLT Project Management

A web application for managing many projects in one place: **timeline (Gantt)**, **progress tracking**, **monthly status updates** and **CAPEX budget / payment tracking**. It follows the same rules as the DCI-2025 Excel progress tracker, reads ProjectLibre files directly, and exports to Excel, ProjectLibre XML, `.pod` and a one-page PDF status report.

---

## 1. Quick start

You need **Python 3.10 or newer**.

| System | Start |
|---|---|
| macOS / Linux | `./run.sh` |
| Windows | double-click `run.bat` |

The first start creates a virtual environment (`.venv`) and installs the packages in `requirements.txt`. Then the browser opens at **http://127.0.0.1:5000**.

Manual start:

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

### Load your existing project (DCI-2025)

1. Click **Import from files** on the Projects page.
2. Select the Excel tracker (`250702_DCI-2025_Progress_Tracker_budget.xlsx`) **and** all monthly `.pod` files (`250702-2026-04-DCI-2025-v1.pod` … `-10-…`) together.
3. Click **Create project**.

The app then:
- reads every month,
- takes the budget scopes, payment terms and milestone task IDs from the Excel `Budget` sheet,
- takes the *Planned Progress* values from the Excel Summary.

You can do the same from the command line: `python app.py import <files…>`.

### Share it with your team (intranet)

```bash
SLT_PM_HOST=0.0.0.0 SLT_PM_PORT=8080 python app.py      # Windows: set SLT_PM_HOST=0.0.0.0 & python app.py
```

Colleagues then open `http://<your-PC-or-server-IP>:8080`. A `Dockerfile` is included for running it on a server (`docker build -t slt-pm . && docker run -p 5000:5000 -v sltpm:/data slt-pm`). There is no login: run it on the internal network only. For wider use, put it behind a reverse proxy with SSO (for example Entra ID).

| Setting | Default |
|---|---|
| `SLT_PM_DB` | `sqlite:///slt_pm.db` (any SQLAlchemy URL works, e.g. PostgreSQL) |
| `SLT_PM_HOST` | `127.0.0.1` |
| `SLT_PM_PORT` | `5000` |

**Backup:** copy the file `slt_pm.db`.

---

## 2. What each page does

| Page | What you do there |
|---|---|
| **All Projects** | A card for every project: health (Green / Amber / Red), actual vs planned progress, progress trend sparkline, budget bar (due / allocated / reserved), overdue count. Create a new project or import one from files. |
| **Dashboard** | Main figures, a progress trend chart (actual vs planned input vs baseline), task-status donut, component progress vs baseline, budget by scope, component status table (current stage, slip, schedule status), overdue tasks, tasks due in the next 45 days, and remarks (Text2). |
| **Tasks & Progress** | Edit the live plan. Click any cell to change name, dates, duration, % complete, baseline dates, Text2 remarks or predecessors (`12FS; 14SS+2`). Tick **Excl.** to set Text1 = 1 (not counted). Add a task or subtask, indent or outdent, move up or down, delete, search and filter. **Baseline ▾** saves the current plan as the baseline. **Record status update** saves the month. |
| **Timeline (Gantt)** | Planned bars filled with % complete, baseline bars underneath, summary brackets and a status-date line. Zoom by week, month, quarter or year, and collapse to components. In the live view, click a bar to edit it. |
| **Budget** | CAPEX reserved, allocated (contract value), reserve − allocation, due to date by payment terms, balance to pay, optional *actually paid by finance*, and the utilisation trend vs work progress. Add or edit scopes with: tender no., component, goods / services allocation, payment terms (% at GRN / PAC / FAC), and the milestone tasks that release each payment. |
| **Status Updates** | Upload this month's `.pod` / `.xml` / `.xlsx` (several files at once work too). Choose *Replace plan* or *Progress only*. See the history of all updates and download PDF, Excel or XML for any month. Rename, restore or delete updates. Edit the **Planned progress (input)** per month. |
| **Project Settings** | Name, project no., manager, currency, status date, description; delete the project. |

The **Status as of** selector in the header switches every page (Dashboard, Timeline, Budget and the downloads) between recorded months and the live plan.

---

## 3. Calculation rules (same as the Excel tracker)

| Item | Rule |
|---|---|
| Counted task | not a summary row **and** Text1 blank. Text1 = 1 excludes it: *not required*, *not applicable*, FAC duration, Stage Gates 3–6… |
| Actual progress | simple average of % complete of counted tasks (each counted task has equal weight) |
| Baseline planned % | for each counted task, the baseline working days (Mon–Fri) elapsed at the status date ÷ its baseline working days |
| Planned (input) | your own monthly target %, like the *Planned Progress* column in Excel |
| Task status | Completed (100 %) · Overdue (planned finish before status date and < 100 %) · In Progress · Not Started · Excluded · Summary |
| Component status | Completed · Delayed (any overdue task or > 30 days slip vs baseline) · Minor Slip · On Track |
| Due to date | Goods allocation × (GRN % + PAC % + FAC % of the milestones at 100 %) + the same for Services |
| Budget status | Not Allocated · Over Reserve (allocation > reserve) · Allocated – No Payment · Payment in Progress · Fully Paid |
| Health | Schedule: Green if within 5 pp of plan (or baseline when no plan is set) and nothing overdue, Amber within 15 pp, otherwise Red. Budget: Red if any scope is over its reserve, Amber if scopes are not yet allocated. |

Summary rows are rolled up like ProjectLibre: earliest start, latest finish, duration-weighted %. They are only for display; they are never counted.

---

## 4. File formats

| Format | Import | Export |
|---|---|---|
| **.pod** (ProjectLibre) | ✔ reads the project XML that ProjectLibre 1.9+ stores inside every `.pod` | ✔ the last uploaded `.pod` with its XML section replaced by the current plan |
| **.xml** (MS Project / ProjectLibre XML) | ✔ tasks, outline, dates, % complete, baseline, Text1/Text2, links | ✔ opens in ProjectLibre (*File › Open*) and MS Project |
| **.xlsx** | ✔ this app's export, or tracker workbooks with month sheets (`ID`, `Task Name`, `% Complete` …), a `Budget` sheet and Summary *Planned Progress* | ✔ Summary, Budget register, Trend, *Current Plan* and one sheet per month, with live formulas (counted %, baseline %, status, change vs previous month) |
| **PDF** | – | ✔ one-page A4 status report: health, main figures, progress trend, component status, budget chart and table, issues and the next 45 days |

> **About .pod export.** A `.pod` has two parts: a ProjectLibre binary block, and an XML copy of the plan. The app rewrites the XML copy. Independent readers such as MPXJ read the updated plan from it. Whether ProjectLibre itself shows your web edits when it opens the `.pod` depends on its version. The safe round trip is: **Download › XML**, open it in ProjectLibre (*File › Open*), then *Save As* `.pod`.

---

## 5. Monthly routine

1. Update the plan in ProjectLibre as usual and save `250702-YYYY-MM-….pod`.
2. Go to **Status Updates**, drop the file in and click **Upload & update**. The status date comes from `YYYY-MM` in the file name, so it is the month end.
3. Check **Dashboard** and **Budget**, and update allocations or payment terms when a contract is awarded.
4. **Download › PDF** for the one-page status report, or **Excel** for the full tracker.

Alternatively, update % complete directly in **Tasks & Progress** and press **Record status update**.

---

## 6. Project structure

```
app.py                Flask app: REST API, import / update / export
models.py             SQLAlchemy models: Project, Task, Snapshot, Scope, PlannedPoint
services/mspdi.py     .pod / .xml reader and writer (MS Project XML)
services/calc.py      progress, schedule, baseline and budget calculations
services/excel_io.py  Excel import / export
services/pdf_report.py  one-page PDF (ReportLab + Matplotlib)
static/               single-page web UI (vanilla JS + Chart.js, no build step)
tests/test_smoke.py   end-to-end test: python tests/test_smoke.py <folder with sample .pod/.xlsx>
```

The REST API (JSON) is under `/api/…`, for example `GET /api/projects/<id>/dashboard?view=latest`. Other tools such as PulseKPI can read it.
