# Retail Customer Segmentation & Sales Analysis

**Capstone project — IBM SkillsBuild Data Analytics with AI Academic Internship Programme**  
Conducted by BharatCares in association with AICTE  
Submitted by **Aryan Verma**

An end-to-end analysis of 1,067,371 transaction lines from a UK online gift retailer,
answering one question: **which customers is this business actually built on, and which
of them are leaving?**

The analysis segments the customer base with RFM scoring, cross-checks that segmentation
against unsupervised K-Means clustering, estimates customer lifetime value, and measures
retention with an acquisition-cohort matrix — then names the 683 accounts worth
£1.69M that have stopped ordering.

---

## Dataset

| | |
|---|---|
| **Dataset** | **Online Retail II** |
| **Link** | **https://archive.ics.uci.edu/dataset/502/online+retail+ii** |
| Direct download | https://archive.ics.uci.edu/static/public/502/online+retail+ii.zip |
| Publisher | UCI Machine Learning Repository (dataset 502) |
| Donated by | Dr Daqing Chen, London South Bank University |
| Licence | Creative Commons Attribution 4.0 International (CC BY 4.0) |
| Size | 45.6 MB zip holding one `.xlsx` workbook of the same size (stored, not compressed) |
| Sheets | `Year 2009-2010`, `Year 2010-2011` |
| Rows | 1,067,371 across both sheets |
| Period | 1 December 2009 to 9 December 2011 |
| Granularity | One row per product line per invoice |

**The dataset is not committed to this repository.** The notebook downloads it from UCI on
first run, extracts it to `data/raw/`, and never modifies it afterwards.  
UCI's throughput varies a lot — the same download measured 31 seconds on one run and several minutes on
another, so allow time for it.

Columns: `Invoice`, `StockCode`, `Description`, `Quantity`, `InvoiceDate`, `Price`,
`Customer ID`, `Country`. The notebook adds `SourceSheet` and the engineered
`Revenue = Quantity × Price`.

---

## What the project does

1. **Acquires and profiles the raw workbook** — row counts, missing values, duplicates,
   invoice-prefix distribution, non-product stock codes, negative quantities.
2. **Cleans it with a logged audit trail** — nine steps, each recording the rows and
   revenue it removed and the reason. 96.4% of rows are retained.
3. **Explores the business** — revenue trend and seasonality, new vs returning revenue,
   order processing patterns, products, geography, returns and cancellations, and
   customer concentration.
4. **Segments customers with RFM** — rank-based quintile scoring into nine named
   segments, each with its revenue contribution.
5. **Cross-checks with K-Means** — the same features `log1p`-transformed and
   standardised, `k` scanned from 2 to 8 with inertia and silhouette recorded.
6. **Estimates customer lifetime value** — historical and projected, with the projection
   tested against actual revenue rather than published unchallenged.
7. **Measures retention by cohort** — a monthly acquisition matrix, corrected for
   censoring at both ends of the observation window.
8. **Writes findings, recommendations and limitations** — every number quoted from a
   computed value, and an explicit list of what the data cannot support.

---

## Key results

| Result | Number |
|---|---|
| Clean revenue analysed | **£20,517,554** across 39,675 orders |
| Identified customers | 5,854 (85.4% of clean revenue is attributable) |
| Revenue from the top 20% of customers | **77.2%** (1,170 customers) |
| Revenue from the top 1% of customers | **31.9%** (58 customers) |
| Champions segment | 1,555 customers (26.6%) → **71.5% of revenue** |
| High-value customers who have stopped ordering | **683 accounts, £1,689,620** |
| Second-year revenue from previously-acquired customers | **91.6%** |
| Returns | £1,523,788 = **7.43%** of gross, spiking to 19.05% in January 2011 |
| Largest line in the dataset | £168,469.60 — **cancelled 12 minutes later**, net £0.00 |
| Products whose rank changes net of returns | **8 of the top 10** |
| Christmas-acquired retention at month 3 | **9.3%** vs 21.9% for every other month |
| RFM segments that map onto a single K-Means cluster (≥60%) | **9 of 9** |

Three of these contradict what a first pass at this dataset produces: the largest "sale"
was a cancellation, so gross product rankings are wrong at the top; peak-season customers
retain at less than half the rate of everyone else; and `InvoiceDate` records back-office
keying rather than customer intent, which invalidates any "best day to advertise"
conclusion. Each was found by checking a claim against the data instead of accepting a
plausible number.

---

## Technologies used

| Tool | Version verified | Role |
|---|---|---|
| Python | 3.14.6 | Language for the entire analysis |
| Jupyter Notebook | 7.6.2 | Executable document holding code, charts and narrative |
| pandas | 3.0.5 | Loading, cleaning, joining, grouping, cohort pivots |
| NumPy | 2.5.3 | Vectorised arithmetic, `log1p` transform |
| Matplotlib | 3.11.2 | All 14 figures, on one explicit chart theme |
| scikit-learn | 1.9.1 | `StandardScaler`, `KMeans`, `silhouette_score` |
| openpyxl | 3.1.5 | Reads the two sheets of the source workbook |
| python-docx | 1.2.0 | Generates the report from `outputs/facts.json` |

Lower bounds are in [requirements.txt](requirements.txt). Matplotlib is used rather than an
interactive charting library because an interactive chart does not survive being committed
to a notebook and read by someone else — every figure here is a static image that renders
identically on any machine.

---

## Setup and run

Developed and run on **Python 3.14.6** (Windows 11). The version bounds in
`requirements.txt` are the oldest releases whose APIs the code relies on; only 3.14.6 was
actually exercised.

```bash
git clone https://github.com/BU1lDR/IBM-BharatCares---Internship-CAPSTONE-PROJECT-.git
cd IBM-BharatCares---Internship-CAPSTONE-PROJECT-

python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
jupyter notebook AryanVerma_RetailCustomerSegmentationAnalysis.ipynb
```

Then run **Kernel → Restart Kernel and Run All Cells**.

Measured run times on the development machine (Windows 11, Python 3.14.6), excluding the
download itself:

| Run | Time |
|---|---|
| First run — parses both `.xlsx` sheets and writes a CSV cache | **1 min 28 s** |
| Every later run — reads the CSV cache | **33 s** |

Deleting `data/online_retail_II.csv` forces the workbook to be re-parsed; doing that
reproduced a byte-identical cache (same SHA-256), so the cache is not a source of drift.
No cell depends on a later cell, so *Run All* from a fresh kernel reproduces every number
and every figure.

To regenerate the Word report from the notebook's outputs:

```bash
python tools/build_report.py
```

### Troubleshooting

* **The download fails or produces a corrupt zip.** Download
  `https://archive.ics.uci.edu/static/public/502/online+retail+ii.zip` manually, extract
  `online_retail_II.xlsx`, and place it in `data/raw/`. The notebook detects the cached
  file and skips the download. Do not resume a partial download — a truncated file raises
  `zipfile.BadZipFile`.
* **`ModuleNotFoundError: sklearn`.** The package is `scikit-learn`, not `sklearn` —
  `pip install -r requirements.txt` installs the right name.
* **`jupyter: command not found`.** The `notebook` package provides the command and is in
  `requirements.txt`; make sure the virtual environment is activated.

---

## Repository contents

| Path | What it is |
|---|---|
| [AryanVerma_RetailCustomerSegmentationAnalysis.ipynb](AryanVerma_RetailCustomerSegmentationAnalysis.ipynb) | **The analysis.** 75 cells, all executed, no errors, 14 figures. |
| [requirements.txt](requirements.txt) | Dependencies, with the version each was verified on. |
| [AryanVerma_ProjectReport.docx](AryanVerma_ProjectReport.docx) | **The report.** 17 sections, 14 figures, 15 tables. |
| [README.md](README.md) | This file. |
| [tools/build_notebook.py](tools/build_notebook.py) | Holds the notebook's cell sources and writes the `.ipynb`. |
| [tools/run_notebook.py](tools/run_notebook.py) | Executes the notebook headlessly; exits non-zero if any cell errors. |
| [tools/build_report.py](tools/build_report.py) | Generates the `.docx` from `outputs/facts.json` + `outputs/figures/`. |
| [outputs/facts.json](outputs/facts.json) | Every computed number the notebook produced, plus the segment, cluster, cohort, product and country tables. |
| `outputs/figures/` | The 14 exported PNGs. Regenerated by a notebook run, so not committed. |
| `data/raw/` | The downloaded zip and workbook. Not committed. |
| `data/online_retail_II.csv` | Both sheets concatenated, written once to make later runs fast. Not committed. |

### Why the report is generated rather than written

`tools/build_report.py` reads `outputs/facts.json` and `outputs/figures/` and writes
`AryanVerma_ProjectReport.docx`. Nothing in the report is transcribed by hand, so the
document cannot drift away from the analysis: change a cleaning rule, re-run the notebook,
re-run the script, and every affected number in the report updates. The script fails loudly
if a fact or a figure it needs is missing rather than emitting a blank.

---

## Method notes worth knowing

* **RFM scoring is rank-based.** Frequency is heavily tied — thousands of customers have
  exactly one or two orders — and `pd.qcut` on raw values fails on ties. The notebook ranks
  first (`pd.qcut(x.rank(method="first"), 5, ...)`), which guarantees five populated bands.
* **Recency is measured from a fixed reference date** (the day after the last transaction
  in the file), not from today, so the analysis is reproducible in any future year.
* **Cancellations and write-offs are set aside, not deleted.** A `C`-prefixed invoice is a
  customer return; a negative quantity on an ordinary zero-priced invoice is a stock
  adjustment. Merging the two would overstate the customer return rate. Six `A`-prefixed
  invoices carry −£147,614 of bad debt that the common "drop invoices starting with C"
  recipe leaves in the revenue line.
* **Two analysis bases are carried deliberately.** `sales` (all valid revenue lines,
  including guest checkouts) is used for revenue, products, geography and returns;
  `sales_id` (known customer ID only) is used for RFM, K-Means, CLV and cohorts, because
  all four require an identified customer.
* **`k = 4` disagrees with the silhouette metric, and the notebook says so.** The metric
  preferred `k = 2` (0.4376 against 0.3649). Two clusters cannot be marketed to
  differently, so interpretability was chosen over the metric and the cost is reported
  rather than hidden.
* **The 12-month CLV projection is an upper bound, not a forecast.** Summed across the
  base it is 2.43× actual second-year revenue, and it ranks a dormant segment above
  Champions. Every prioritisation decision in the report therefore rests on *historical*
  value, which is measured.

## What this analysis does not support

* **Any pricing or product-mix decision** — the dataset has no cost of goods, so every
  figure is revenue, not margin.
* **Any claim that contacting a customer *causes* a return** — there is no campaign data
  and no experiment in the file, which is why the win-back recommendation specifies a 20%
  untreated control group.
* **Any conclusion about when customers prefer to shop** — the timestamp is a processing
  time, not a purchase time.
* **Using the CLV projection as a forecast.**

The report states all nine limitations in full, including the attribution gap: 22.9% of
clean sales lines, carrying 14.6% of clean revenue, have no customer ID and are invisible
to every customer-level model here.

---

## References

* Chen, D. (2019). *Online Retail II* [Data set]. UCI Machine Learning Repository.
  https://archive.ics.uci.edu/dataset/502/online+retail+ii (CC BY 4.0)
* Hughes, A. M. (1994). *Strategic Database Marketing*. Probus Publishing.
* Fader, P. S., Hardie, B. G. S., & Lee, K. L. (2005). RFM and CLV: Using iso-value curves
  for customer base analysis. *Journal of Marketing Research*, 42(4), 415–430.
* Rousseeuw, P. J. (1987). Silhouettes: a graphical aid to the interpretation and
  validation of cluster analysis. *Journal of Computational and Applied Mathematics*, 20,
  53–65.
* Pedregosa, F. et al. (2011). Scikit-learn: Machine learning in Python. *Journal of
  Machine Learning Research*, 12, 2825–2830.
