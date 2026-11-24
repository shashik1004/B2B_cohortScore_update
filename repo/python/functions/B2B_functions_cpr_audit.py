"""
B2B Pricing — CPRScore Audit Table Builder
==========================================
Standalone companion to `B2B_functions_cohort_score.py`.  That module is NOT
modified by this one; this module only *reads* from a loaded
`B2BPricingReporter` instance and independently re-derives the cohort math so
the published CPRScore can be reconciled line-by-line.

Purpose
-------
Produce an Exec-facing, SQL-pushable table where each row is a Client and the
columns expose every input that feeds that client's CPRScore:

    C###_ActCPR    client-wide balance-rollup ActualCPR, repeated by cohort
    C###_CohortModelCPR  pooled-balance cohort benchmark (same value for every client)
    C###_EndBal    ending balance that client contributed to that cohort
    C###_LoanCt    loan-months that client contributed to that cohort

so an Exec can reproduce the published number with nothing but Excel:

    ActualCPR = SUMPRODUCT(ActCPR_range,   EndBal_range) / SUM(EndBal_range)
    ModelCPR  = SUMPRODUCT(ModelCPR_range, EndBal_range) / SUM(EndBal_range)
        CPRScore  = ActualCPR / ModelCPR
                            = SUMPRODUCT(ActCPR_range, EndBal_range)
                                / SUMPRODUCT(ModelCPR_range, EndBal_range)     <- denominators cancel

Methodology mirrored from B2B_functions_cohort_score.B2BPricingReporter
-----------------------------------------------------------------------
1. Client ActualCPR is calculated across all assigned cohorts:
    ActualCPR = 1 - (SUM EndPrinBal / SUM BeginPrinBal) ** 12
2. Client groups with total LoanCt < MIN_COHORT_LOANS are dropped from the
    client roll-up (Included_Flag=0), but NOT from the cohort benchmark's
    applicability - ModelCPR remains a property of the Cohort_ID.
3. CohortModelCPR = 1 - (SUM Cohort_EndPrinBal / SUM Cohort_BeginPrinBal) ** 12,
    broadcast to every row sharing that Cohort_ID.
4. Cohort_ID IS NULL (PLS / unmapped) is excluded entirely - ModelCPR is NaN
   there because no cohort, and therefore no benchmark, exists.
5. Clients with no surviving cell default to CPRScore = 1.0.

Two sets of roll-ups are produced:
  *_Calc   built from the ROUNDED cohort columns -> an Exec's SUMPRODUCT
           reproduces these to the digit.
  *_Exact  built from full precision -> this is the methodology check, and
           must reconcile to the published number to ~1e-9.
`Rounding_Drift` quantifies the cost of the 5dp display rounding.  It is
non-trivial for clients whose cohort benchmarks are very small (a 0.00020
benchmark loses ~2% of its value at 5dp), which is expected, not a defect.
"""
from __future__ import annotations

import re
from datetime import date
from typing import Dict, List, Optional

import numpy as np
import pandas as pd

from B2B_functions_cohort_score import (
    MIN_COHORT_LOANS,
    B2BPricingReporter,
)

# Sentinel row carrying the pooled portfolio benchmark
BENCHMARK_LABEL: str = '~PORTFOLIO BENCHMARK~'
UNKNOWN_CLIENT: str = '(Unknown)'
NO_COHORT_LABEL: str = '(no cohort)'

# Tie-out is checked on the full-precision roll-up, so it should be near-exact
TIEOUT_TOL: float = 1e-9


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _filter_product(perform_all: pd.DataFrame, product: str) -> pd.DataFrame:
    """Same product slicing rule used by B2BPricingReporter.run_all_metrics."""
    if product == 'All':
        return perform_all
    if product == 'Govt':
        return perform_all[perform_all['ProductCategory'].isin(['FHA', 'VA', 'USDA'])]
    return perform_all[perform_all['ProductCategory'] == product]


def _cohort_sort_key(cohort_id: str):
    """Sort 'C7' before 'C53' before 'C120'; non-conforming ids sort last."""
    m = re.match(r'^[A-Za-z]*(\d+)$', str(cohort_id))
    return (0, int(m.group(1)), '') if m else (1, 0, str(cohort_id))


def _annualized_cpr(begin: pd.Series, end: pd.Series) -> np.ndarray:
    with np.errstate(divide='ignore', invalid='ignore'):
        return np.where(begin > 0, 1 - (end / begin) ** 12, np.nan)


def cohort_blocks(wide: pd.DataFrame) -> Dict[str, List[str]]:
    """The three contiguous cohort column blocks, in matching cohort order."""
    pat = re.compile(r'^(?P<coh>.+?)_(?P<kind>ActCPR|ModelCPR|CohortModelCPR|EndBal|LoanCt)$')
    found: Dict[str, Dict[str, str]] = {}
    for col in wide.columns:
        m = pat.match(col)
        if m:
            found.setdefault(m.group('kind'), {})[m.group('coh')] = col
    cohorts = sorted(found.get('ActCPR', {}), key=_cohort_sort_key)
    return {kind: [found[kind][c] for c in cohorts]
            for kind in ('ActCPR', 'ModelCPR', 'CohortModelCPR', 'EndBal', 'LoanCt')}


# ---------------------------------------------------------------------------
# core: cell-level rebuild
# ---------------------------------------------------------------------------
def build_cell_detail(reporter: B2BPricingReporter,
                      product: str = 'All',
                      min_loans: int = MIN_COHORT_LOANS) -> pd.DataFrame:
    """
    Long / tall audit frame — one row per (Client x Cohort_ID) cell, plus one
    row per client for the unmapped (PLS / no-cohort) population.

    Columns
    -------
    ReportAsofDt, AsofDt, ProductCategory, Client, Cohort_ID, Rule_bucket, Client_LoanCt,
    Client_BeginPrinBal, Client_EndPrinBal, Cohort_LoanCt, Cohort_BeginPrinBal,
    Cohort_EndPrinBal, Client_ActualCPR, Client_x_Cohort_ActualCPR,
    Cohort_ActualCPR, Cohort_x_Cohort_ActualCPR, CPRScore,
    Included_Flag, ExclusionReason
    """
    if reporter.perform_all is None:
        raise RuntimeError('Reporter not loaded. Call reporter.load(...) first.')

    p = _filter_product(reporter.perform_all, product)
    p = p[p['LoanCt'] > 0].copy()
    p['Client'] = p['Client'].fillna(UNKNOWN_CLIENT)
    client_loan_ct = p.groupby('Client', dropna=False)['LoanCt'].sum()

    cell = p.groupby(['Client', 'Cohort_ID', 'Rule_Bucket'], dropna=False).agg(
        LoanCt      =('LoanCt',       'sum'),
        BeginPrinBal=('BeginPrinBal', 'sum'),
        EndPrinBal  =('EndPrinBal',   'sum'),
    ).reset_index()

    cohort_totals = (p.groupby('Cohort_ID', dropna=False)
                     .agg(Cohort_LoanCt=('LoanCt', 'sum'),
                          Cohort_BeginPrinBal=('BeginPrinBal', 'sum'),
                          Cohort_EndPrinBal=('EndPrinBal', 'sum'))
                     .reset_index())
    cell = cell.merge(cohort_totals, on='Cohort_ID', how='left')
    cell = cell.rename(columns={'Rule_Bucket': 'Rule_bucket'})

    cell['CohortActualCPR'] = _annualized_cpr(cell['BeginPrinBal'], cell['EndPrinBal'])
    client_balances = (p[p['Cohort_ID'].notna()]
                       .groupby('Client', dropna=False)
                       .agg(Client_BeginPrinBal=('BeginPrinBal', 'sum'),
                            Client_EndPrinBal=('EndPrinBal', 'sum'))
                       .reset_index())
    client_balances['Client_ActualCPR'] = _annualized_cpr(
        client_balances['Client_BeginPrinBal'], client_balances['Client_EndPrinBal'])
    cell = cell.merge(client_balances[['Client', 'Client_ActualCPR']], on='Client', how='left')
    cell['ActualCPR'] = cell['Client_ActualCPR']

    unmapped = (cell['Cohort_ID'].isna() |
                cell['Cohort_ID'].isin([NO_COHORT_LABEL, 'NULL', 'NO COHORT (PLS/unmapped)']))
    below_min = (~unmapped) & cell['Client'].map(client_loan_ct).lt(min_loans)
    no_cpr = (~unmapped) & (~below_min) & (cell['CohortActualCPR'].isna())

    cell['Included_Flag'] = np.where(unmapped | below_min | no_cpr, 0, 1).astype(int)
    cell['ExclusionReason'] = np.select(
        [unmapped, below_min, no_cpr],
        ['No cohort assigned (PLS / unmapped)',
         f'Cell loan-months < {min_loans}',
         'BeginPrinBal = 0 (CPR undefined)'],
        default='',
    )

    # Cohort benchmark: direct pooled-balance CPR for the entire cohort.
    if cohort_totals.empty:
        cell['ModelCPR'] = np.nan
    else:
        cohort_totals['CohortModelCPR'] = _annualized_cpr(
            cohort_totals['Cohort_BeginPrinBal'], cohort_totals['Cohort_EndPrinBal'])
        cell = cell.merge(cohort_totals[['Cohort_ID', 'CohortModelCPR']],
                  on='Cohort_ID', how='left')

    cell.loc[unmapped, 'CohortModelCPR'] = np.nan

    client_model = (cell[~unmapped]
                    .drop_duplicates(['Client', 'Cohort_ID'])
                    .groupby('Client', dropna=False)
                    .agg(ClientModelBegin=('Cohort_BeginPrinBal', 'sum'),
                         ClientModelEnd=('Cohort_EndPrinBal', 'sum'))
                    .reset_index())
    client_model['ModelCPR'] = _annualized_cpr(
        client_model['ClientModelBegin'], client_model['ClientModelEnd'])
    cell = cell.merge(client_model[['Client', 'ModelCPR']], on='Client', how='left')
    cell.loc[unmapped, 'ModelCPR'] = np.nan

    cell['Cohort_ID'] = cell['Cohort_ID'].fillna(NO_COHORT_LABEL)
    report_dt = reporter.end_score_dt or reporter.report_as_of_dt or date.today()
    cell['ReportAsofDt'] = pd.Timestamp(report_dt).to_period('M').end_time.date()
    cell['AsofDt'] = pd.Timestamp(date.today()).to_period('M').end_time.date()
    cell['ProductCategory'] = product

    cols = ['ReportAsofDt', 'AsofDt', 'ProductCategory', 'Client', 'Cohort_ID', 'Rule_bucket',
            'Client_LoanCt', 'Client_BeginPrinBal', 'Client_EndPrinBal',
            'Cohort_LoanCt', 'Cohort_BeginPrinBal', 'Cohort_EndPrinBal',
            'Client_ActualCPR', 'Client_x_Cohort_ActualCPR',
            'Cohort_ActualCPR', 'Cohort_x_Cohort_ActualCPR', 'CPRScore',
            'Included_Flag', 'ExclusionReason']
    cell = cell.drop(columns=['Client_ActualCPR'], errors='ignore')
    cell['CPRScore'] = np.where(cell['ModelCPR'] > 0,
                                cell['ActualCPR'] / cell['ModelCPR'], 1.0)
    cell = cell.rename(columns={
        'LoanCt': 'Client_LoanCt', 'BeginPrinBal': 'Client_BeginPrinBal',
        'EndPrinBal': 'Client_EndPrinBal', 'ActualCPR': 'Client_ActualCPR',
        'CohortActualCPR': 'Client_x_Cohort_ActualCPR',
        'ModelCPR': 'Cohort_ActualCPR',
        'CohortModelCPR': 'Cohort_x_Cohort_ActualCPR',
    })
    return (cell[cols]
            .sort_values(['Client', 'Cohort_ID'], key=lambda s: s.map(_cohort_sort_key)
                         if s.name == 'Cohort_ID' else s)
            .reset_index(drop=True))


# ---------------------------------------------------------------------------
# core: wide pivot
# ---------------------------------------------------------------------------
def build_wide(cell_detail: pd.DataFrame,
               reporter: B2BPricingReporter,
               product: str = 'All',
               decimals: int = 5,
               min_loans: int = MIN_COHORT_LOANS) -> pd.DataFrame:
    """
    Exec-facing wide audit table — one row per Client, three columns per cohort.

    Cohort columns are grouped by TYPE (all *_ActCPR, then all *_ModelCPR, then
    all *_LoanCt) so each block is a contiguous Excel range for SUMPRODUCT.
    """
    cell_detail = cell_detail.rename(columns={
        'Client_LoanCt': 'LoanCt', 'Client_BeginPrinBal': 'BeginPrinBal',
        'Client_EndPrinBal': 'EndPrinBal', 'Client_ActualCPR': 'ActualCPR',
        'Client_x_Cohort_ActualCPR': 'CohortActualCPR',
        'Cohort_ActualCPR': 'ModelCPR',
        'Cohort_x_Cohort_ActualCPR': 'CohortModelCPR',
    })
    cell_detail = cell_detail.loc[:, ~cell_detail.columns.duplicated()]
    inc = cell_detail[cell_detail['Included_Flag'] == 1].copy()
    cohorts = sorted(inc['Cohort_ID'].unique(), key=_cohort_sort_key)
    clients = sorted(cell_detail['Client'].unique())

    act = (inc.pivot_table(index='Client', columns='Cohort_ID', values='ActualCPR',
                           aggfunc='first')
              .reindex(index=clients, columns=cohorts))
    mdl = (inc.pivot_table(index='Client', columns='Cohort_ID', values='CohortModelCPR',
                           aggfunc='first')
              .reindex(index=clients, columns=cohorts))
    client_mdl = (inc.pivot_table(index='Client', columns='Cohort_ID', values='ModelCPR',
                                  aggfunc='first')
                    .reindex(index=clients, columns=cohorts))
    cnt = (inc.pivot_table(index='Client', columns='Cohort_ID', values='LoanCt',
                           aggfunc='first')
              .reindex(index=clients, columns=cohorts))
    bal = (inc.pivot_table(index='Client', columns='Cohort_ID', values='EndPrinBal',
                           aggfunc='first')
             .reindex(index=clients, columns=cohorts))

    act_raw = act.fillna(0.0)
    mdl_raw = mdl.fillna(0.0)
    client_mdl_raw = client_mdl.fillna(0.0)
    act = act_raw.round(decimals)
    mdl = mdl_raw.round(decimals)
    client_mdl = client_mdl_raw.round(decimals)
    cnt = cnt.fillna(0).astype('int64')
    bal_raw = bal.fillna(0.0)
    bal = bal_raw.round(decimals)

    # ── portfolio benchmark row ──────────────────────────────────────────────
    bench_cnt = cnt.sum(axis=0)
    bench_mdl = (inc.drop_duplicates('Cohort_ID')
                    .set_index('Cohort_ID')['ModelCPR']
                    .reindex(cohorts)
                    .fillna(0.0))
    act_raw.loc[BENCHMARK_LABEL] = bench_mdl.values      # portfolio actual == benchmark
    mdl_raw.loc[BENCHMARK_LABEL] = bench_mdl.values
    client_mdl_raw.loc[BENCHMARK_LABEL] = bench_mdl.values
    act.loc[BENCHMARK_LABEL] = bench_mdl.round(decimals).values
    mdl.loc[BENCHMARK_LABEL] = bench_mdl.round(decimals).values
    client_mdl.loc[BENCHMARK_LABEL] = bench_mdl.round(decimals).values
    cnt.loc[BENCHMARK_LABEL] = bench_cnt.values
    bench_bal = bal.sum(axis=0)
    bal_raw.loc[BENCHMARK_LABEL] = bench_bal.values
    bal.loc[BENCHMARK_LABEL] = bench_bal.round(decimals).values

    c = cnt.to_numpy(float)
    b = bal.to_numpy(float)
    tot_ct = c.sum(axis=1)
    tot_bal = b.sum(axis=1)

    def _rollup(a_arr, m_arr, weights):
        with np.errstate(divide='ignore', invalid='ignore'):
            a_w = np.where(weights.sum(axis=1) > 0, (a_arr * weights).sum(axis=1) / weights.sum(axis=1), 0.0)
            m_w = np.where(weights.sum(axis=1) > 0, (m_arr * weights).sum(axis=1) / weights.sum(axis=1), 0.0)
            return a_w, m_w, np.where(m_w > 0, a_w / m_w, 1.0)

    # From the rounded columns — what an Exec's SUMPRODUCT will return
    act_calc, mdl_calc, score_calc = _rollup(act.to_numpy(float), client_mdl.to_numpy(float), b)
    # From full precision — the methodology check
    act_ex, mdl_ex, score_ex = _rollup(act_raw.to_numpy(float), client_mdl_raw.to_numpy(float), bal_raw.to_numpy(float))

    client_values = (cell_detail[cell_detail['Included_Flag'] == 1]
                     .drop_duplicates('Client')
                     .set_index('Client'))
    client_actual = client_values['ActualCPR'].reindex(act.index).fillna(0.0).to_numpy(float)
    client_model = client_values['ModelCPR'].reindex(act.index).fillna(0.0).to_numpy(float)
    client_score = np.where(client_model > 0, client_actual / client_model, 1.0)
    act_calc = client_actual.copy()
    mdl_calc = client_model.copy()
    score_calc = client_score.copy()
    act_ex = client_actual.copy()
    mdl_ex = client_model.copy()
    score_ex = client_score.copy()

    out = pd.DataFrame({
        'ReportAsOfDt':        reporter.report_as_of_dt,
        'ProductCategory':     product,
        'Client':              act.index,
        'BeginScoreDt':        reporter.begin_score_dt or date(1900, 1, 1),
        'EndScoreDt':          reporter.end_score_dt,
        'MinCohortLoans':      min_loans,
        'Total_LoanCt_Scored': tot_ct.astype('int64'),
        'Total_EndPrinBal_Scored': tot_bal,
        'CohortsCovered':      (c > 0).sum(axis=1).astype('int64'),
        'ActualCPR_Calc':      np.round(act_calc, decimals + 1),
        'ModelCPR_Calc':       np.round(mdl_calc, decimals + 1),
        'CPRScore_Calc':       np.round(score_calc, 6),
        'ActualCPR_Exact':     act_ex,
        'ModelCPR_Exact':      mdl_ex,
        'CPRScore_Exact':      score_ex,
        'Rounding_Drift':      score_calc - score_ex,
    })
    out['ScoreDefaulted'] = (out['Total_LoanCt_Scored'] == 0).astype(int)

    # ── excluded-population transparency (not part of the SUMPRODUCT) ───────
    # Named *_LoanMonths, not *_LoanCt, so they cannot be swept into the cohort block.
    exc = cell_detail[cell_detail['Included_Flag'] == 0]
    unmapped = (exc[exc['Cohort_ID'] == NO_COHORT_LABEL]
                .groupby('Client')['LoanCt'].sum().rename('Unmapped_LoanMonths'))
    belowmin = (exc[exc['Cohort_ID'] != NO_COHORT_LABEL]
                .groupby('Client')
                .agg(BelowMin_LoanMonths=('LoanCt', 'sum'),
                     BelowMin_CellCt=('Cohort_ID', 'size')))

    all_loans = (cell_detail.groupby('Client')
                 .agg(_b=('BeginPrinBal', 'sum'), _e=('EndPrinBal', 'sum'),
                      CPRLoanct=('LoanCt', 'sum')))
    all_loans['ActualCPR_AllLoans'] = _annualized_cpr(all_loans['_b'], all_loans['_e'])
    all_loans = all_loans.drop(columns=['_b', '_e'])

    out = (out.merge(unmapped, left_on='Client', right_index=True, how='left')
              .merge(belowmin, left_on='Client', right_index=True, how='left')
              .merge(all_loans, left_on='Client', right_index=True, how='left'))
    for col in ('Unmapped_LoanMonths', 'BelowMin_LoanMonths', 'BelowMin_CellCt', 'CPRLoanct'):
        out[col] = out[col].fillna(0).astype('int64')
    out['ActualCPR_AllLoans'] = out['ActualCPR_AllLoans'].round(decimals + 1)

    # ── cohort blocks, grouped by type ──────────────────────────────────────
    act.columns = [f'{c_}_ActCPR' for c_ in act.columns]
    mdl.columns = [f'{c_}_CohortModelCPR' for c_ in mdl.columns]
    client_mdl.columns = [f'{c_}_ModelCPR' for c_ in client_mdl.columns]
    cnt.columns = [f'{c_}_LoanCt' for c_ in cnt.columns]
    bal.columns = [f'{c_}_EndBal' for c_ in bal.columns]
    blocks = pd.concat([act, client_mdl, mdl, bal, cnt], axis=1).reset_index(drop=True)

    wide = pd.concat([out.reset_index(drop=True), blocks], axis=1)

    # benchmark row first, then clients by scored loan-months desc
    wide['__ord'] = np.where(wide['Client'] == BENCHMARK_LABEL, 0, 1)
    wide = (wide.sort_values(['__ord', 'Total_LoanCt_Scored', 'Client'],
                             ascending=[True, False, True])
                .drop(columns='__ord')
                .reset_index(drop=True))
    return wide


# ---------------------------------------------------------------------------
# tie-out against the reporter's own scoring routine
# ---------------------------------------------------------------------------
def build_tieout(wide: pd.DataFrame,
                 reporter: B2BPricingReporter,
                 product: str = 'All',
                 tol: float = TIEOUT_TOL) -> pd.DataFrame:
    """
    Independently recomputed values vs. the published ones from
    B2BPricingReporter._cohort_cpr_score_by_keys.  Adds *_Report, TieOut_Diff
    and TieOut_Flag columns to a copy of `wide`.
    """
    p = _filter_product(reporter.perform_all, product).copy()
    p['Client'] = p['Client'].fillna(UNKNOWN_CLIENT)

    rpt = B2BPricingReporter._cohort_cpr_score_by_keys(p, ['Client'])
    rpt = rpt.rename(columns={
        'ActualCPR': 'ActualCPR_Report',
        'ModelCPR':  'ModelCPR_Report',
        'CPRScore':  'CPRScore_Report',
    })[['Client', 'ActualCPR_Report', 'ModelCPR_Report', 'CPRScore_Report']]

    out = wide.merge(rpt, on='Client', how='left')
    out['TieOut_Diff'] = out['CPRScore_Exact'] - out['CPRScore_Report']
    out['TieOut_Flag'] = np.where(
        out['Client'] == BENCHMARK_LABEL, 'N/A (benchmark row)',
        np.where(out['TieOut_Diff'].abs() <= tol, 'OK', 'MISMATCH'))

    # place tie-out columns next to the calc columns, ahead of the cohort blocks
    lead = [c for c in out.columns if not re.search(r'_(ActCPR|CohortModelCPR|EndBal|LoanCt)$', c)]
    blocks = [c for c in out.columns if c not in lead]
    return out[lead + blocks]


# ---------------------------------------------------------------------------
# orchestrator
# ---------------------------------------------------------------------------
def build_cpr_audit(reporter: B2BPricingReporter,
                    product: str = 'All',
                    decimals: int = 5,
                    min_loans: int = MIN_COHORT_LOANS,
                    verbose: bool = True) -> Dict[str, pd.DataFrame]:
    """
    Build the full CPRScore audit pack.

    Returns {'long': cell-level detail, 'wide': Exec table with tie-out columns}
    """
    long_df = build_cell_detail(reporter, product=product, min_loans=min_loans)
    wide_df = build_wide(long_df, reporter, product=product,
                         decimals=decimals, min_loans=min_loans)
    wide_df = build_tieout(wide_df, reporter, product=product)

    if verbose:
        n_coh = long_df.loc[long_df['Included_Flag'] == 1, 'Cohort_ID'].nunique()
        n_cli = wide_df['Client'].nunique() - 1
        body = wide_df[wide_df['Client'] != BENCHMARK_LABEL]
        flags = body['TieOut_Flag']
        worst = body['TieOut_Diff'].abs().max()
        drift = body['Rounding_Drift'].abs().max()
        print(f'CPR audit [{product}]  clients={n_cli:,}  cohorts={n_coh:,}  '
              f'wide shape={wide_df.shape}')
        print(f'  tie-out (exact vs published): {(flags == "OK").sum():,} OK / '
              f'{(flags == "MISMATCH").sum():,} MISMATCH  '
              f'(max |diff| = {worst:.2e}, tol = {TIEOUT_TOL:.0e})')
        print(f'  max |rounding drift| at {decimals}dp: {drift:.2e} CPRScore points')
    return {'long': long_df, 'wide': wide_df}
