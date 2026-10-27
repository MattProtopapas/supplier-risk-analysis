# Supplier Scorecards & Risk Analysis

A small pipeline that generates a fake supplier dataset, validates it, scores every supplier
on price, lead-time variability and defect rate, rates the risk of each purchased category,
and works out when dual-sourcing pays for itself — writing the input data and the results
into a single Excel workbook.

It is the inbound mirror of *OTIF* (on-time-in-full), which covers outbound delivery: here the
question is **which suppliers should you trust?**

- **Weighted scorecard** — each supplier is scored 0–100 on three criteria, then weighted (default 40 / 30 / 30):
  - *Price* — `100 × cheapest price in the category / this supplier's price`
  - *Lead-time variability* — `100 × (1 − min(CV / 0.5, 1))`, where CV = lead-time standard deviation / mean lead time (an unpredictable supplier is worse than a slow one)
  - *Defects* — `100 × (1 − min(defect rate / 8%, 1))`
  - Tiers: **A** ≥ 75 (preferred), **B** ≥ 60 (approved), **C** below (watch list)
- **Category risk** — incumbent vs. best alternative, expected yearly disruption loss, a High / Medium / Low rating and a recommended action (qualify a backup, review the incumbent, mitigate, maintain).
- **Dual-sourcing economics** — total annual cost of a sourcing set-up =
  purchase cost + defect cost + safety-stock holding cost (`z × daily demand × lead-time std dev × price × holding rate`) + expected disruption loss (`outage probability × daily demand × outage days × price × shortage multiplier`) + a fixed yearly cost for managing a second supplier.
  With two suppliers the lead-time variability diversifies (`sqrt(w₁²σ₁² + w₂²σ₂²)`) and a full loss needs *both* to fail; if only one fails, the survivor covers part of the gap with surge capacity.
  The single-source baseline is the cheapest supplier in the category, so any saving comes from **spreading risk**, not from just switching supplier. The tool searches every alternate and a grid of volume splits for the best set-up, and reports the **break-even outage probability** — the primary supplier's yearly disruption probability above which dual-sourcing wins.

> The disruption, defect-cost and surge-capacity inputs are assumptions you should replace with
> your own numbers (all are command-line flags below) — dual-sourcing economics are very sensitive to them.

## How it works

Three scripts, run in order, each consuming the previous one's output:

1. **`generate_data.py`** — uses [Faker](https://faker.readthedocs.io/) to create a fake
   dataset of suppliers grouped by purchased category (price, lead-time mean and variability, defect
   rate, yearly disruption probability, country, incumbent flag) and writes it to
   `data/suppliers_raw.csv`. Cheaper suppliers are made a little less reliable on average, so price and
   reliability genuinely trade off. A small percentage of rows are deliberately corrupted (negative
   prices, out-of-range rates, blanks, a duplicate ID) so the next step has real errors to catch.
2. **`validate_data.py`** — checks every row in `data/suppliers_raw.csv` against the schema in
   `supplier_common.py` (via [pydantic](https://docs.pydantic.dev/)), drops invalid or
   duplicate rows, and writes:
   - `data/suppliers_validated.csv` — the clean rows
   - `data/validation_report.txt` — what was rejected and why
3. **`analyze_suppliers.py`** — reads `data/suppliers_validated.csv`, builds the scorecard, rates category risk,
   evaluates dual-sourcing, and writes a new `output/supplier_analysis.xlsx` workbook with four sheets:
   - **Input Data** — the validated records the analysis ran on
   - **Supplier Scorecard** — component scores, total score, tier and rank per supplier (tier C highlighted)
   - **Category Risk** — incumbent vs. best alternative, expected disruption loss, rating and action (High risk highlighted)
   - **Dual-Sourcing** — single vs. dual annual cost, saving, break-even outage probability and a recommendation (categories where it pays off highlighted)

`supplier_common.py` holds the shared data model and file paths used by all three scripts.

> **Example Results** : [sample_results.xlsx](./sample_results.xlsx)

## Setup

### 1. Get the code

```powershell
git clone <your-repo-url>
cd supplier-risk-analysis
```

### 2. Create and activate a virtual environment

**Windows (PowerShell):**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**macOS / Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

Requires Python 3.10+ (uses `statistics.NormalDist` and PEP 604 type hints — no `scipy`
or `numpy` needed).

## Usage

Run the three scripts in order from the project root, with the virtual environment
activated:

```powershell
python generate_data.py --categories 6 --suppliers-per-category 4 --error-rate 0.05 --seed 42
python validate_data.py
python analyze_suppliers.py
```

Then open `output/supplier_analysis.xlsx`.

### Options

| Script | Flag | Default | Description |
| --- | --- | --- | --- |
| `generate_data.py` | `--categories` | `6` | Number of purchased categories (1–8) |
| | `--suppliers-per-category` | `4` | Qualified suppliers per category |
| | `--error-rate` | `0.05` | Fraction of rows to deliberately corrupt (`0` for a perfectly clean dataset) |
| | `--seed` | none | Random seed, for reproducible output |
| | `--output` | `data/suppliers_raw.csv` | Where to write the raw data |
| `validate_data.py` | `--input` | `data/suppliers_raw.csv` | File to validate |
| | `--output` | `data/suppliers_validated.csv` | Where to write the cleaned data |
| | `--report` | `data/validation_report.txt` | Where to write the rejection report |
| `analyze_suppliers.py` | `--input` | `data/suppliers_validated.csv` | File to analyze |
| | `--output` | `output/supplier_analysis.xlsx` | Where to write the results workbook |
| | `--weight-price` / `--weight-lead-time` / `--weight-defects` | `0.40` / `0.30` / `0.30` | Scorecard weights |
| | `--defect-cost-multiplier` | `2.0` | Cost of a defective unit as a multiple of its price |
| | `--holding-rate` | `0.25` | Yearly inventory holding cost as a share of unit value |
| | `--service-level` | `0.95` | Service level used to size safety stock |
| | `--outage-days` | `30` | Length of a major supplier outage |
| | `--shortage-multiplier` | `1.5` | Cost of an undelivered unit as a multiple of its price |
| | `--surge-capacity` | `0.5` | Share of a failed supplier's volume the other can absorb |
| | `--dual-fixed-cost` | `15000` | Yearly cost of qualifying and managing a second supplier |

Every script can also point at a custom input/output path, so you can drop in your own real
supplier CSV (matching the columns in `supplier_common.py`) and run it straight through
`validate_data.py` and `analyze_suppliers.py`, skipping `generate_data.py`.

## Project layout

```
supplier-risk-analysis/
├── supplier_common.py      # shared schema + file paths
├── generate_data.py        # step 1: fake data generation
├── validate_data.py        # step 2: schema validation / cleaning
├── analyze_suppliers.py    # step 3: scorecard, category risk, dual-sourcing -> Excel
├── requirements.txt
├── data/                    # generated CSVs + validation report (not committed)
└── output/                  # generated Excel workbook (not committed)
```

`data/` and `output/` are gitignored (apart from `.gitkeep` placeholders) since their
contents are reproducible by re-running the scripts — as is `.venv/`, so the virtual
environment itself is never pushed to the repo.


> **Disclaimer** — This software is provided for educational and demonstration purposes only. See the full [Disclaimer & EULA](./DISCLAIMER.md) for details.
