"""
FUN (Favourable / Unfavourable / Neutral) client classification for the B2B rawTotal output.

TotalScore, higher is favourable (lower cutoff inclusive, matching the Excel Client sheet):
    Z            = TotalScore on the ProductCategory='All', MetricName='Total', MetricValue='Total' row
    F (Favourable)   : TotalScore >= Z * 1.25
    U (Unfavourable) : TotalScore <= Z * 0.75
    N (Neutral)      : otherwise
"""
from __future__ import annotations

import numpy as np
import pandas as pd

FAVOURABLE_MULTIPLIER: float = 1.25
UNFAVOURABLE_MULTIPLIER: float = 0.75
FUN_COLUMN: str = 'FUN_Classification'


def derive_score_thresholds(raw_total: pd.DataFrame,
                            score_column: str = 'TotalScore',
                            favourable_multiplier: float = FAVOURABLE_MULTIPLIER,
                            unfavourable_multiplier: float = UNFAVOURABLE_MULTIPLIER) -> dict:
    totals = raw_total.loc[
        raw_total['MetricName'].eq('Total')
        & raw_total['MetricValue'].eq('Total')
        & raw_total['ProductCategory'].eq('All'),
        score_column,
    ]
    if len(totals) != 1:
        raise ValueError(f'Expected one All/Total/Total portfolio row, found {len(totals)}')
    baseline = pd.to_numeric(totals, errors='coerce').iloc[0]
    if pd.isna(baseline) or not np.isfinite(baseline) or baseline <= 0:
        raise ValueError(f'Portfolio {score_column} baseline must be finite and positive')
    return {
        'Score column': score_column,
        'Z': float(baseline),
        'Upper threshold': float(baseline) * favourable_multiplier,
        'Lower threshold': float(baseline) * unfavourable_multiplier,
    }


def add_fun_classification(raw_total: pd.DataFrame,
                           score_column: str = 'TotalScore',
                           favourable_multiplier: float = FAVOURABLE_MULTIPLIER,
                           unfavourable_multiplier: float = UNFAVOURABLE_MULTIPLIER):
    """Return (copy of raw_total with FUN_Classification on Client rows only, thresholds)."""
    thresholds = derive_score_thresholds(raw_total, score_column,
                                         favourable_multiplier, unfavourable_multiplier)
    out = raw_total.copy()
    is_client = out['MetricName'].eq('Client')
    scores = pd.to_numeric(out.loc[is_client, score_column], errors='coerce')
    if not np.isfinite(scores.to_numpy(dtype=float)).all():
        raise ValueError(f'Client rows have missing or non-finite {score_column}')

    out[FUN_COLUMN] = pd.Series(np.nan, index=out.index, dtype='object')
    out.loc[is_client, FUN_COLUMN] = np.select(
        [scores.ge(thresholds['Upper threshold']), scores.le(thresholds['Lower threshold'])],
        ['F', 'U'],
        default='N',
    )
    return out, thresholds


def client_score_table(raw_total_fun: pd.DataFrame) -> pd.DataFrame:
    """Client rows of rawTotal, Excel columns D:X (MetricName..TotalScore) plus FUN_Classification."""
    cols = list(raw_total_fun.columns[3:24])
    if cols[0] != 'MetricName' or cols[-1] != 'TotalScore':
        raise ValueError(f'Unexpected rawTotal layout for columns D:X: {cols[0]} .. {cols[-1]}')
    return (raw_total_fun.loc[raw_total_fun['MetricName'].eq('Client'), cols + [FUN_COLUMN]]
            .reset_index(drop=True))
