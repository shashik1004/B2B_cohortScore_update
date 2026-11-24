# Cohort Assignment Guide (schema version 4)

Use these files to tag **unseen loans** with a prepayment cohort. Do **not** retrain, re-merge or re-band on new data; the files are fixed lookups learned from the 2026-06-30 and 2026-07-31 training tapes.

## Files

| File | Purpose |
|---|---|
| `cohort_bands_config.json` | Band edges, labels, LoanType rules, missing-value defaults and cohort key order. |
| `prinbal_map.csv` | Maps `LoanType` + `RemainingTermBand` + `FICOOrigBand` + `PrinBalBand` to `PrinBalGroup`. |
| `spread_map.csv` | Maps the above + `PrinBalGroup` + `RateSpreadBand` to `RateSpreadGroup`. |
| `cohort_definitions.csv` | Final cohort table: `Rule`, `N`, `UPB`, `Mean_CPR`, `Std_CPR`, `SE_Mean`, `CV` and flags. |

`cohort_holdout_metrics.csv` and `remainingterm_tree_ablation.csv` are evidence only and are not needed for scoring.

There is no FICO map and no term map: FICO and RemainingTerm bands are fixed and never merged.

## Required input columns

`LoanType`, `RemainingTerm` (months), `FICOOrig`, `PrinBal`, `IntRate`, `Market_Rate`.

`Market_Rate` must be the market rate in effect for **that loan's own snapshot**, taken from its own tape. Do not use a single constant across snapshots.

Rows with a missing or non-numeric `PrinBal`, `IntRate` or `Market_Rate` cannot be scored. Remove or repair them first.

## Step 1: Normalize LoanType

1. Trim whitespace and convert to upper case.
2. `CONV-PMI` becomes `CONV`.
3. Keep `CONV`, `FHA` and `VA` as they are.
4. Anything else, including blank, null and values such as `USDA`, becomes `OTHERS`.

| Raw LoanType | Normalized |
|---|---|
| `CONV`, `conv`, `CONV-PMI` | `CONV` |
| `FHA` | `FHA` |
| `VA` | `VA` |
| `USDA`, `""`, null, anything else | `OTHERS` |

The same rules are stored in `loan_type_normalization` in the config.

## Step 2: Create bands

All numeric intervals are **right-inclusive** (for example `240` falls in `<=240mo`), and the lowest edge is included. In the config, `"-inf"` and `"inf"` are strings and must be converted to floating infinities.

| Band column | Source | Rule |
|---|---|---|
| `RemainingTermBand` | `RemainingTerm` | `<=240mo` or `>240mo`. Missing, non-numeric or negative becomes `Missing`. |
| `FICOOrigBand` | `FICOOrig` | `<=750` or `>750`. Missing, non-numeric, `<=300` or `>=850` becomes `<=750`. |
| `PrinBalBand` | `PrinBal` | `<=150k`, `150-250k`, `250-400k`, `>400k` (edges 150000, 250000, 400000). |
| `RateSpreadBand` | `IntRate - Market_Rate` | 18 bands using `spread_edges` and `spread_labels` from the config. |

Use the labels exactly as written in the config. They are string join keys.

## Step 3: Join the lookups

1. Join `prinbal_map.csv` on `LoanType, RemainingTermBand, FICOOrigBand, PrinBalBand` to get `PrinBalGroup`.
2. Join `spread_map.csv` on `LoanType, RemainingTermBand, FICOOrigBand, PrinBalGroup, RateSpreadBand` to get `RateSpreadGroup`.
3. Join `cohort_definitions.csv` on `LoanType, RemainingTermBand, FICOOrigBand, PrinBalGroup, RateSpreadGroup` (the `cohort_keys` in the config) to get the cohort statistics.

Read the CSVs as strings with default NA parsing disabled so labels such as `OTHERS` are not altered: `dtype=str, keep_default_na=False`.

Every combination of the bands above has a mapping, so every scorable loan matches exactly one cohort. If a row does not match, the input is bad (for example a missing balance); do not invent a cohort.

## Reference implementation (tested)

Run from the folder that contains `cohort_artifacts_FHA_VA_Oct26`. Tested on the 2026-07-31 tape: all 601,706 eligible loans matched, using 109 cohorts.

```python
import json
from pathlib import Path

import numpy as np
import pandas as pd

ART = Path("cohort_artifacts_FHA_VA_Oct26")
cfg = json.loads((ART / "cohort_bands_config.json").read_text())


def edges(values):
    return [-np.inf if v == "-inf" else np.inf if v == "inf" else float(v) for v in values]


def normalize_loan_type(values):
    values = values.astype("string").str.strip().str.upper().replace({"CONV-PMI": "CONV"})
    return values.where(values.isin(["CONV", "FHA", "VA"]), "OTHERS")


def band(values, edge_key, label_key):
    return pd.cut(values, edges(cfg[edge_key]), labels=cfg[label_key], include_lowest=True).astype("object")


def assign_cohorts(df):
    out = df.copy()
    out["LoanType"] = normalize_loan_type(out["LoanType"]).astype(object)

    term = pd.to_numeric(out["RemainingTerm"], errors="coerce")
    term = term.where(term >= 0)
    out["RemainingTermBand"] = band(term, "remainingterm_edges", "remainingterm_labels").fillna(cfg["remainingterm_missing_label"])

    fico = pd.to_numeric(out["FICOOrig"], errors="coerce")
    fico = fico.where((fico > 300) & (fico < 850))
    out["FICOOrigBand"] = band(fico, "ficoorig_edges", "ficoorig_labels").fillna(cfg["ficoorig_missing_label"])

    out["RateSpread"] = out["IntRate"] - out["Market_Rate"]
    out["PrinBalBand"] = band(out["PrinBal"], "prinbal_edges", "prinbal_labels")
    out["RateSpreadBand"] = band(out["RateSpread"], "spread_edges", "spread_labels")

    read = lambda name: pd.read_csv(ART / name, dtype=str, keep_default_na=False)
    prinbal_map, spread_map, definitions = read("prinbal_map.csv"), read("spread_map.csv"), read("cohort_definitions.csv")

    out = out.merge(prinbal_map, on=["LoanType", "RemainingTermBand", "FICOOrigBand", "PrinBalBand"],
                    how="left", validate="many_to_one")
    out = out.merge(spread_map, on=["LoanType", "RemainingTermBand", "FICOOrigBand", "PrinBalGroup", "RateSpreadBand"],
                    how="left", validate="many_to_one")
    out = out.merge(definitions.drop(columns=["PrinBalRange"]),
                    on=["LoanType", "RemainingTermBand", "FICOOrigBand", "PrinBalGroup", "RateSpreadGroup"],
                    how="left", validate="many_to_one")
    unmatched = out["Rule"].isna()
    if unmatched.any():
        raise ValueError(f"{unmatched.sum()} rows did not match a cohort (check PrinBal, IntRate, Market_Rate).")
    return out
```

Output columns added: `LoanType` (normalized), the four band columns, `PrinBalGroup`, `RateSpreadGroup`, `Rule`, `N`, `UPB`, `Mean_CPR`, `Std_CPR`, `SE_Mean`, `CV`, `Reliable`, `BelowCountFloor`, `Extrapolated`. Because definitions are read as strings, convert `Mean_CPR`, `Std_CPR`, `SE_Mean`, `CV` and `N` to numeric before using them. Flags are the strings `True` and `False`.

## Reading the result

- `Mean_CPR` is the cohort's typical `ModelCPR` and `Std_CPR` is its spread, both from training data.
- `Reliable = False` or `BelowCountFloor = True` means the estimate is based on few observations. Use it with caution.
- `Extrapolated = True` means the segment had no training loans. The rule exists for coverage, but `Mean_CPR` is blank. Do not substitute a value silently; flag it.
- `Rule` is a readable description. Merged bands appear as a combined range, for example `RateSpread (-Inf, -1.75%]`.

## Eligibility used in training

Training used only active loans (`PaidOff_Flag_ADM == 0`), excluded `InvestorCatgy == PLS`, and excluded sellers `OCWEN`, `OCWEN REO`, `UNKNOWN` and blank or null. Apply the same filters if you want like-for-like comparisons. `InvestorCatgy`, `LoanAge`, `LTVOrig` and `ArmFlag` are **not** cohort inputs.

## Known limitations

- Training used two monthly snapshots (June and July 2026), so cohort statistics are not out-of-time validated.
- `SE_Mean` treats repeated monthly observations of the same loan as independent, so it understates uncertainty.
- Sparse LoanType, term and FICO segments (for example `OTHERS`, `Missing` term and small FHA/VA segments) can remain below the 1500-observation floor. They are flagged in `cohort_definitions.csv`.
- The earlier `*_v2` files use a different schema (investor, loan age) and are **not** interchangeable with these files.
