# B2B Cohort CPRScore Approach

## Project Location

```
Workspace: e:\Development\Shashi\B2B-Pricing-Model\
Functions: repo\python\functions\B2B_functions_cohort_score.py
Notebook:  repo\python\B2B_Reporting_cohort_score.ipynb
Excel out: Cohort_Analysis\B2B_reporting_raw\B2B_cohort_score_202601.xlsx
DB server: VMC2VMDBRADC01P  (Analytics_Rpt, ADM databases, Windows Auth)
Python:    3.11.9, kernel name "B2B"
```

---

## What the New Cohort Approach Does

The notebook replaces the old `ModelCPR` benchmark (a pricing-model score stored in `B2BPricingHist.ModelCPR`) with a **peer-cohort actual CPR benchmark** built from `Analytics_Rpt.md.B2B_revamp_cohorts`.

**Key table:** `Analytics_Rpt.md.B2B_revamp_cohorts`  
One row per `LoanNum × AsofDt`. Loans are assigned a `Cohort_ID` based on their characteristics (FICO band, LTV, IntRate, InvestorCatgy, LoanAge, Rule_Bucket). Cohorts only exist from Jan 2025 onward; earlier months have NULL Cohort_ID (PLS loans and older data are excluded from scoring).

---

## CPRScore Algorithm (`_cohort_cpr_score_by_keys` in the functions file)

```
For each (MetricGroup × Cohort_ID) cell where LoanCt_gc ≥ 10:
    ActualCPR_gc = 1 − (ΣEndPrinBal / ΣBeginPrinBal)^12

For each Cohort_ID (pooled across ALL groups in that cohort):
    ModelCPR_cohort = Σ(ActualCPR_gc × LoanCt_gc) / Σ(LoanCt_gc)

For each MetricGroup:
    ActualCPR = Σ(ActualCPR_gc  × LoanCt_gc) / Σ(LoanCt_gc)
    ModelCPR  = Σ(ModelCPR_cohort × LoanCt_gc) / Σ(LoanCt_gc)
    CPRScore  = ActualCPR / ModelCPR
```

**Rules:**
- `MIN_COHORT_LOANS = 10` — cells below threshold are excluded from CPRScore computation
- PLS / `Cohort_ID = NULL` → default `CPRScore = 1.0`
- Groups with no valid cohort data → `CPRScore = 1.0`
- RecapScore logic is **unchanged** from the old approach

---

## `build_perform_summary_all` — Key SQL Change vs. Old Approach

Old approach: `INNER JOIN ADM.ODS.vwOrigLoan` only  
New approach: adds `LEFT JOIN Analytics_Rpt.md.B2B_revamp_cohorts c ON a.LoanNum = c.LoanNum AND a.AsofDt = c.AsofDt` to bring `Cohort_ID` into every row. **`Cohort_ID` is now a group-by column** in `perform_all`, so each pre-aggregated row carries its cohort assignment.

---

## `load()` — Simplified API

```python
reporter = B2BPricingReporter(engine).load(report_as_of_dt=date(2026, 1, 31))
# Automatically derives:
#   begin_fund_dt  = 2025-02-01  (13 months prior)
#   end_fund_dt    = 2026-01-31
#   end_score_dt   = 2026-01-31  (all history up to report date)
#   begin_score_dt = None        (all history)
```

`perform_all` has **8,133,104 rows × 36 months** for the Jan 2026 run.

---

## Notebook Structure (`B2B_Reporting_cohort_score.ipynb`)

| Cell | Purpose |
|------|---------|
| User Inputs | `REPORT_AS_OF_DT = date(2026, 1, 31)` |
| Imports | Loads `B2B_functions_cohort_score.py` |
| Load | `reporter.load(report_as_of_dt=...)` |
| DQ | Cohort coverage for report month (99.7% of non-PLS loans have Cohort_ID) |
| All Metrics | `reporter.run_all_metrics(['Total','State','FICO','LTV','DTI','Interest Rate','Product','Avg Loan Bal','Channel','Client'])` |
| Product loop | Loops `['All','Conventional','FHA','VA','USDA','Govt']` |
| GroupBy tabs | `reporter.build_groupby_summary(['Client','PropState'])` and `['Client','PropState','ProductCategory']` |
| Excel export | 8 sheets: rawTotal, rawConv, rawFHA, rawVA, rawUSDA, rawGovt, rawClientState, rawClientStateProduct |

---

## Verified Results (NORTHWEST BANK, Jan 2026)

- CPRScore (new cohort) = **1.136364** — 13.6% faster than peer cohort
- CPRScore (old pricing-model) = **0.788331** — appeared 21% slower than model
- The two benchmarks give opposite conclusions because old ModelCPR (~6.5%) was higher than peer actual CPR (~4.1%)

---

## Old Approach vs. New Approach Comparison

| Dimension | Old Approach (`B2B_functions_monthly_report.py`) | New Approach (`B2B_functions_cohort_score.py`) |
|-----------|--------------------------------------------------|------------------------------------------------|
| **Benchmark (ModelCPR)** | Pricing model's predicted CPR — a loan-level score stored in `B2BPricingHist.ModelCPR` | Peer-group actual CPR — weighted avg of other clients in the same characteristic cohort |
| **Benchmark source** | `B2BPricingHist.ModelCPR` column | `B2B_revamp_cohorts.Cohort_ID` + peer clients' actual prepay |
| **Threshold** | Total loan-months ≥ 10 (aggregate across all grouped cells) | Per-cohort loan count ≥ 10 |
| **Scope** | All `B2BPricingHist` history up to `end_score_dt` | All 36 months of `B2BPricingHist` |
| **Exclusions** | XFER/Bulk loans removed; inner join to `vwOrigLoan`; no Reverse | PLS/unmapped loans (no `Cohort_ID`) default CPRScore = 1.0 |

---

## Related Files

| File | Purpose |
|------|---------|
| `functions/B2B_functions_cohort_score.py` | Core functions library — cohort CPRScore logic |
| `functions/B2B_functions_monthly_report.py` | Original functions library — old pricing-model approach |
| `B2B_Reporting_cohort_score.ipynb` | Main reporting notebook (new approach) |
| `B2B_Reporting_monthly_report.ipynb` | Main reporting notebook (old approach) |
| `B2B_CPRScore_Audit.ipynb` | Step-by-step audit notebook for cohort approach |
| `B2B_OldApproach_CPRScore_Audit.ipynb` | Step-by-step audit notebook for old approach + side-by-side comparison |
