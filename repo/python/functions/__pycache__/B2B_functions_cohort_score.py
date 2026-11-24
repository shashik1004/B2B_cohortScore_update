"""
B2B Pricing Reporting — Cohort-Based CPR Scoring
=================================================
Extends B2B_functions_monthly_report.py with a cohort-aware CPRScore methodology.

Key change: CPRScore now uses cohort benchmarks from
`Analytics_Rpt.md.B2B_revamp_cohorts` instead of the pre-stored ModelCPR
column in B2BPricingHist.

New CPRScore algorithm
----------------------
For each MetricGroup:
    ActualCPR = 1 − (ΣEndPrinBal / ΣBeginPrinBal)^12

For each Cohort_ID (pooled across ALL groups in that cohort):
    CohortModelCPR = 1 − (ΣEndPrinBal / ΣBeginPrinBal)^12

For each MetricGroup (rolled up across all its cohorts):
    ModelCPR   = 1 − (ΣCohortEndBal / ΣCohortBeginBal)^12
    CPRScore   = ActualCPR / ModelCPR

Rules:
- PLS / unmapped loans (Cohort_ID IS NULL) are excluded from CPRScore.
- Cells with < MIN_COHORT_LOANS (10) default to CPRScore = 1.0.
- RecapScore methodology is unchanged.

New load() parameter
--------------------
    report_as_of_dt : date
        Pass a month-end date to produce a historical report.  All date
        windows (funding, score) are derived automatically.  Leave None
        to run for the current latest available month.

Created from: B2B_functions_monthly_report.py (Sotiris Adamakis, 05/30/2026)
"""
from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Optional, List, Dict

import numpy as np
import pandas as pd
import sqlalchemy
from sqlalchemy import text
from sqlalchemy.engine import Engine


# ---------------------------------------------------------------------------
# Minimum loan count per cohort cell — cells below this default to CPRScore=1
# ---------------------------------------------------------------------------
MIN_COHORT_LOANS: int = 10


# ---------------------------------------------------------------------------
# SQL connection helper
# ---------------------------------------------------------------------------
def sql_connection(server: str = "VMC2VMDBRADC01P",
                   database: str = "Analytics_Rpt",
                   driver: str = "ODBC Driver 17 for SQL Server") -> Engine:
    """Return a SQLAlchemy engine connected to the Analytics_Rpt database."""
    sql_string = (
        f"mssql+pyodbc:///?odbc_connect="
        f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};Trusted_Connection=yes"
    )
    return sqlalchemy.create_engine(sql_string, fast_executemany=True)


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------
class DateHelpers:
    """SQL EOMONTH / DATEADD equivalents."""

    @staticmethod
    def eomonth(d: date, months_offset: int = 0) -> date:
        year = d.year + (d.month - 1 + months_offset) // 12
        month = (d.month - 1 + months_offset) % 12 + 1
        last_day = calendar.monthrange(year, month)[1]
        return date(year, month, last_day)

    @staticmethod
    def add_days(d: date, days: int) -> date:
        return d + timedelta(days=days)


# ---------------------------------------------------------------------------
# Band classifier  (unchanged)
# ---------------------------------------------------------------------------
class BandClassifier:
    """Bucket loan-level numeric values into report bands."""

    @staticmethod
    def _band_fico(s: pd.Series) -> np.ndarray:
        v = pd.to_numeric(s, errors='coerce').to_numpy(dtype='float64')
        return np.where(
            np.isnan(v) | (v < 580) | (v > 850), 'zUnknown',
            np.select(
                [v >= 780, v >= 740, v >= 700, v >= 660, v >= 620, v >= 580],
                ['a_780+', 'b_740', 'c_700', 'd_660', 'e_620', 'f_580'],
                default='zUnknown',
            ),
        )

    @staticmethod
    def _band_ltv(s: pd.Series) -> np.ndarray:
        v = pd.to_numeric(s, errors='coerce').to_numpy(dtype='float64')
        return np.where(
            np.isnan(v) | (v < 0.10) | (v > 0.99), 'zUnknown',
            np.select(
                [v >= 0.95, v >= 0.90, v >= 0.85, v >= 0.80, v >= 0.70, v >= 0.50],
                ['a_95+', 'b_90', 'c_85', 'd_80', 'e_70', 'f_50'],
                default='g_<50',
            ),
        )

    @staticmethod
    def _band_dti(s: pd.Series) -> np.ndarray:
        v = pd.to_numeric(s, errors='coerce').to_numpy(dtype='float64')
        return np.where(
            np.isnan(v) | (v < 0.05) | (v > 0.65), 'zUnknown',
            np.select(
                [v >= 0.50, v >= 0.45, v >= 0.40, v >= 0.35, v >= 0.30, v >= 0.25],
                ['a_50+', 'b_45', 'c_40', 'd_35', 'e_30', 'f_25'],
                default='g_<25',
            ),
        )

    @staticmethod
    def _band_rate(s: pd.Series) -> np.ndarray:
        v = pd.to_numeric(s, errors='coerce').to_numpy(dtype='float64')
        return np.where(
            np.isnan(v) | (v < 0.02) | (v > 0.10), 'zUnknown',
            np.select(
                [v >= 0.0800, v >= 0.0750, v >= 0.0700, v >= 0.0675,
                 v >= 0.0650, v >= 0.0625, v >= 0.0600, v >= 0.0575,
                 v >= 0.0550, v >= 0.0500, v >= 0.0450, v >= 0.0400,
                 v >= 0.0350, v >= 0.0300],
                ['a_8.00+', 'b_7.50', 'c_7.00', 'd_6.75',
                 'e_6.50', 'f_6.25', 'g_6.00', 'h_5.75',
                 'i_5.50', 'j_5.00', 'k_4.50', 'l_4.00',
                 'm_3.50', 'n_3.00'],
                default='o_<3%',
            ),
        )

    @staticmethod
    def _band_alb(s: pd.Series) -> np.ndarray:
        v = pd.to_numeric(s, errors='coerce').to_numpy(dtype='float64')
        nan_mask = np.isnan(v)
        v_safe = np.where(nan_mask, 0.0, v)
        rounded = (np.ceil(v_safe / 50000.0) * 50000).astype('int64').astype(str)
        return np.where(
            nan_mask | (v_safe <= 100000), '100000',
            np.where(v_safe > 1000000, '1000001', rounded),
        )

    @classmethod
    def add_bands(cls, df: pd.DataFrame, *,
                  fico_col: str = 'FICO', ltv_col: str = 'LTV',
                  dti_col: str = 'DTI', rate_col: str = 'Rate',
                  alb_col: str = 'FundedLoanAmt') -> pd.DataFrame:
        out = df.copy()
        out['FICOBand']    = cls._band_fico(out[fico_col])
        out['LTVBand']     = cls._band_ltv(out[ltv_col])
        out['DTIBand']     = cls._band_dti(out[dti_col])
        out['IntRateBand'] = cls._band_rate(out[rate_col])
        out['ALB']         = cls._band_alb(out[alb_col])
        return out


# ---------------------------------------------------------------------------
# Weighted stats  (unchanged)
# ---------------------------------------------------------------------------
class WeightedStats:
    @staticmethod
    def wavg(values: pd.Series, weights: pd.Series, lo: float, hi: float):
        mask  = values.fillna(0).between(lo, hi)
        w     = weights.where(mask, 0).fillna(0)
        v     = values.where(mask, 0).fillna(0)
        denom = w.sum()
        return np.nan if denom == 0 else (v * w).sum() / denom

    @staticmethod
    def safe_div(num, den):
        if den is None or den == 0 or pd.isna(den):
            return np.nan
        return num / den


# ---------------------------------------------------------------------------
# Constants  (unchanged)
# ---------------------------------------------------------------------------
DEFAULT_METRIC_ORDER: List[str] = [
    'Total', 'State', 'FICO', 'LTV', 'DTI', 'Interest Rate',
    'Product', 'Avg Loan Bal', 'Channel', 'Client',
]

PRODUCT_LOOP: List[str] = ['All', 'Conventional', 'FHA', 'VA', 'USDA', 'Govt']

SELLER_COLS: List[str] = [
    'Account Manager_East', 'Account Manager_West', 'Account Manager_Central',
    'Account Manager_NULL',
    'ClientID_Delegated - Best Efforts', 'ClientID_Delegated - Mandatory',
    'ClientID_Direct', 'ClientID_FHLMC CRX', 'ClientID_FNMA SMP',
    'ClientID_GNMA PIIT', 'ClientID_Non-delegated',
]

METRIC_COLUMN_MAP: Dict[str, str] = {
    'Total':         '__TOTAL__',
    'State':         'PropState',
    'FICO':          'FICOBand',
    'LTV':           'LTVBand',
    'DTI':           'DTIBand',
    'Interest Rate': 'IntRateBand',
    'Product':       'ProductCategory',
    'Avg Loan Bal':  'ALB',
    'Channel':       'Channel',
    'Client':        'Client',
    'PropState':     'PropState',
    'ProductCategory': 'ProductCategory',
}


def _metric_column(df: pd.DataFrame, metric: str) -> pd.Series:
    if metric not in METRIC_COLUMN_MAP:
        raise KeyError(f"Unknown metric '{metric}'. Valid: {sorted(METRIC_COLUMN_MAP)}")
    col = METRIC_COLUMN_MAP[metric]
    if col == '__TOTAL__':
        return pd.Series(['Total'] * len(df), index=df.index)
    if col not in df.columns:
        raise KeyError(f"Column '{col}' (for metric '{metric}') not found in dataframe")
    return df[col]


# ---------------------------------------------------------------------------
# Result container  (unchanged)
# ---------------------------------------------------------------------------
@dataclass
class B2BPricingResults:
    all_metrics_by_product: Dict[str, pd.DataFrame] = field(default_factory=dict)
    client_state:           Optional[pd.DataFrame]  = None
    client_state_product:   Optional[pd.DataFrame]  = None
    client_metrics_snapshot: Optional[pd.DataFrame] = None


# ---------------------------------------------------------------------------
# Main reporter class
# ---------------------------------------------------------------------------
class B2BPricingReporter:
    """
    End-to-end builder for the B2B Pricing report with cohort-based CPRScore.

    Typical usage:
        rep = B2BPricingReporter(engine)
        rep.load(report_as_of_dt=date(2026, 1, 31))   # historical run
        # -- or --
        rep.load()                                      # current month

        all_metrics = rep.run_all_metrics(METRIC_ORDER, product='All')
        client_state = rep.build_groupby_summary(['Client', 'PropState'])
    """

    def __init__(self, engine: Optional[Engine] = None):
        self.engine:          Optional[Engine]       = engine
        self.client_seller:   Optional[pd.DataFrame] = None
        self.output_template: Optional[pd.DataFrame] = None
        self.perform_all:     Optional[pd.DataFrame] = None
        self.volume_all:      Optional[pd.DataFrame] = None
        self.begin_fund_dt:   Optional[date]         = None
        self.end_fund_dt:     Optional[date]         = None
        self.begin_score_dt:  Optional[date]         = None
        self.end_score_dt:    Optional[date]         = None
        self.report_as_of_dt: Optional[date]         = None

    # ── step 1: client/seller mapping  (unchanged) ───────────────────────────
    def build_client_seller_id_ae(self, begin_fund_dt: date, end_fund_dt: date) -> pd.DataFrame:
        sql = text("""
            SELECT o.Client, o.Region, o.AEName, o.SubChannel, o.SourceSellerID, o.CurrentStatusDt
            FROM ADM.ODS.vwOrigLoan o
            WHERE o.Channel IN ('Correspondent', 'Co-Issue')
              AND o.Client IS NOT NULL
              AND o.FundDt IS NOT NULL
              AND o.FundedLoanAmt > 0
              AND o.FundDt BETWEEN :begin_dt AND :end_dt
              AND o.MortgageType <> 'Reverse'
              AND NOT EXISTS (
                  SELECT 1 FROM Analytics_Rpt.md.B2BPricingHist h
                  WHERE h.LoanNum = o.LoanNum
                    AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
              )
        """)
        df = pd.read_sql(sql, self.engine, params={'begin_dt': begin_fund_dt, 'end_dt': end_fund_dt})

        region_ae = (
            df.groupby(['Client', 'Region', 'AEName'], dropna=False)['CurrentStatusDt']
              .max().reset_index()
              .sort_values('CurrentStatusDt', ascending=False)
        )

        def _agg_region(g, region_value):
            sub  = g[g['Region'].isna()] if pd.isna(region_value) else g[g['Region'] == region_value]
            names = sub['AEName'].dropna().tolist()
            return ', '.join(names) if names else None

        ae_rows = []
        for client, g in region_ae.groupby('Client'):
            ae_rows.append({
                'Client':                  client,
                'Account Manager_East':    _agg_region(g, 'East'),
                'Account Manager_West':    _agg_region(g, 'West'),
                'Account Manager_Central': _agg_region(g, 'Central'),
                'Account Manager_NULL':    _agg_region(g, np.nan),
            })
        ae_df = pd.DataFrame(ae_rows)

        sc_seller = (
            df.groupby(['Client', 'SubChannel', 'SourceSellerID'], dropna=False)['CurrentStatusDt']
              .max().reset_index()
              .sort_values('CurrentStatusDt', ascending=False)
        )
        sub_channels = [
            'Delegated - Best Efforts', 'Delegated - Mandatory', 'Direct',
            'FHLMC CRX', 'FNMA SMP', 'GNMA PIIT', 'Non-delegated',
        ]
        sc_rows = []
        for client, g in sc_seller.groupby('Client'):
            row = {'Client': client}
            for sc in sub_channels:
                sub = g[g['SubChannel'] == sc]
                ids = sub['SourceSellerID'].dropna().astype(str).tolist()
                row[f'ClientID_{sc}'] = ', '.join(ids) if ids else None
            sc_rows.append(row)
        sc_df = pd.DataFrame(sc_rows)

        return ae_df.merge(sc_df, on='Client', how='outer')

    # ── step 2: output template  (unchanged) ─────────────────────────────────
    def build_output_template(self, begin_fund_dt: date, end_fund_dt: date) -> pd.DataFrame:
        sql = text("""
            SELECT a.PropState, a.FICO, a.LTV, a.DTI, a.Rate, a.FundedLoanAmt,
                   a.ProductCategory, a.Channel
            FROM adm.ODS.vwOrigLoan a
            WHERE a.Channel IN ('Correspondent', 'Co-Issue')
              AND a.Client IS NOT NULL
              AND a.FundDt IS NOT NULL
              AND a.FundedLoanAmt > 0
              AND a.MortgageType <> 'Reverse'
              AND NOT EXISTS (
                  SELECT 1 FROM Analytics_Rpt.md.B2BPricingHist h
                  WHERE h.LoanNum = a.LoanNum
                    AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
              )
        """)
        df = pd.read_sql(sql, self.engine)
        df = BandClassifier.add_bands(df)

        rows  = [{'MetricName': 'Total', 'MetricValue': 'Total'}]
        rows += [{'MetricName': 'State',         'MetricValue': v} for v in df['PropState'].dropna().unique()]
        rows += [{'MetricName': 'FICO',          'MetricValue': v} for v in pd.unique(df['FICOBand'])]
        rows += [{'MetricName': 'LTV',           'MetricValue': v} for v in pd.unique(df['LTVBand'])]
        rows += [{'MetricName': 'DTI',           'MetricValue': v} for v in pd.unique(df['DTIBand'])]
        rows += [{'MetricName': 'Interest Rate', 'MetricValue': v} for v in pd.unique(df['IntRateBand'])]
        rows += [{'MetricName': 'Avg Loan Bal',  'MetricValue': v} for v in pd.unique(df['ALB'])]
        rows += [{'MetricName': 'Product',       'MetricValue': v} for v in df['ProductCategory'].dropna().unique()]
        rows += [{'MetricName': 'Channel',       'MetricValue': v} for v in df['Channel'].dropna().unique()]

        out = pd.DataFrame(rows).drop_duplicates()
        out['BeginFundDt'] = begin_fund_dt
        out['EndFundDt']   = end_fund_dt
        return out[['BeginFundDt', 'EndFundDt', 'MetricName', 'MetricValue']]

    # ── step 3: performance summary  ─────────────────────────────────────────
    # CHANGED: adds cohort metadata via LEFT JOIN to B2B_revamp_cohorts
    def build_perform_summary_all(self,
                                  begin_score_dt: Optional[date],
                                  end_score_dt:   Optional[date]) -> pd.DataFrame:
        begin = begin_score_dt or date(1900, 1, 1)
        end   = end_score_dt   or date(9999, 12, 31)

        sql = text("""
            SELECT a.Client, a.PropState, a.AsofDt, a.ProductCategory, a.LoanNum,
                   a.FundedLoanAmt, a.ServAdjPTI, a.CDAdjPTI, a.ServGAAPPTI,
                   a.EquityAmt, a.BeginPrinBal, a.EndPrinBal,
                   a.PIFReason, a.RecapBal, a.PIFBal, a.CohortRecap,
                   b.FICO, b.LTV, b.DTI, b.Rate, b.FundedLoanAmt AS b_FundedLoanAmt,
                   b.Channel,
                   c.Cohort_unique_ID as Cohort_ID,
                   c.Rule_Bucket
            FROM Analytics_Rpt.md.B2BPricingHist2 a
            INNER JOIN ADM.ODS.vwOrigLoan b ON a.LoanNum = b.LoanNum
            LEFT  JOIN Analytics_Rpt.md.B2B_revamp_cohorts c
                   ON  a.LoanNum = c.LoanNum
                   AND a.AsofDt  = c.AsofDt
            WHERE a.FundedLoanAmt > 0
              AND a.AsofDt BETWEEN :begin_dt AND :end_dt
                              AND (a.PropState NOT IN ('PR', 'VI', 'GU'))
              AND b.MortgageType <> 'Reverse'
              AND NOT EXISTS (
                                    SELECT 1 FROM Analytics_Rpt.md.B2BPricingHist2 x
                  WHERE x.LoanNum = a.LoanNum
                    AND (x.ResolType = 'XFER' OR x.Channel = 'Bulk')
              )
        """)
        df = pd.read_sql(sql, self.engine, params={'begin_dt': begin, 'end_dt': end})

        df = BandClassifier.add_bands(df, alb_col='b_FundedLoanAmt')
        df['IsGovt'] = df['ProductCategory'].isin(['FHA', 'VA', 'USDA']).astype('int64')

        # Pre-compute weighted-average numerators (vectorized)
        bw = df['b_FundedLoanAmt'].astype('float64').fillna(0)
        for col, lo, hi in (('FICO', 580, 850), ('LTV', 0.10, 0.99),
                            ('DTI', 0.05, 0.65), ('Rate', 0.02, 0.10)):
            vals         = pd.to_numeric(df[col], errors='coerce').fillna(0)
            mask         = vals.between(lo, hi)
            df[f'_w_{col}'] = bw.where(mask, 0)
            df[f'_n_{col}'] = (vals * bw).where(mask, 0)

        df['_AdjPTIAnnual']  = df['ServAdjPTI'].fillna(0) * 12.0 + df['CDAdjPTI'].fillna(0)
        df['_GAAPPTIAnnual'] = df['ServGAAPPTI'].fillna(0) * 12.0 + df['CDAdjPTI'].fillna(0)
        df['_CohortRecap_w'] = df['CohortRecap'].fillna(0) * df['PIFBal'].fillna(0)

        # Cohort_ID and Rule_Bucket carry the monthly cohort assignment into the audit.
        group_cols = ['Client', 'PropState', 'AsofDt', 'ProductCategory',
                      'FICOBand', 'LTVBand', 'DTIBand', 'IntRateBand', 'ALB',
                      'PIFReason', 'Channel', 'Cohort_ID', 'Rule_Bucket']

        agg = df.groupby(group_cols, dropna=False, sort=False, observed=True).agg(
            LoanCt        =('FundedLoanAmt', 'size'),
            UniqueLoanCt  =('LoanNum',       'nunique'),
            FundedLoanAmt =('FundedLoanAmt', 'sum'),
            _IsGovt       =('IsGovt',        'sum'),
            _n_FICO=('_n_FICO', 'sum'), _w_FICO=('_w_FICO', 'sum'),
            _n_LTV =('_n_LTV',  'sum'), _w_LTV =('_w_LTV',  'sum'),
            _n_DTI =('_n_DTI',  'sum'), _w_DTI =('_w_DTI',  'sum'),
            _n_Rate=('_n_Rate', 'sum'), _w_Rate=('_w_Rate', 'sum'),
            AdjPTIAnnual  =('_AdjPTIAnnual',  'sum'),
            GAAPPTIAnnual =('_GAAPPTIAnnual', 'sum'),
            EquityAmt     =('EquityAmt',      'sum'),
            BeginPrinBal  =('BeginPrinBal',   'sum'),
            EndPrinBal    =('EndPrinBal',     'sum'),
            RecapBal      =('RecapBal',       'sum'),
            PIFBal        =('PIFBal',         'sum'),
            _CohortRecap_w=('_CohortRecap_w', 'sum'),
        ).reset_index()

        with np.errstate(divide='ignore', invalid='ignore'):
            agg['GovtPct']   = np.where(agg['LoanCt'] > 0, agg['_IsGovt'] / agg['LoanCt'], 0.0)
            agg['WAFICO']    = np.where(agg['_w_FICO'] > 0, agg['_n_FICO'] / agg['_w_FICO'], np.nan)
            agg['WALTV']     = np.where(agg['_w_LTV']  > 0, agg['_n_LTV']  / agg['_w_LTV'],  np.nan)
            agg['WADTI']     = np.where(agg['_w_DTI']  > 0, agg['_n_DTI']  / agg['_w_DTI'],  np.nan)
            agg['WAC']       = np.where(agg['_w_Rate'] > 0, agg['_n_Rate'] / agg['_w_Rate'], np.nan)
            agg['CohortRecap'] = np.where(agg['PIFBal'] > 0,
                                          agg['_CohortRecap_w'] / agg['PIFBal'], np.nan)

        keep = group_cols + [
            'GovtPct', 'WAFICO', 'WALTV', 'WADTI', 'WAC',
            'LoanCt', 'UniqueLoanCt', 'FundedLoanAmt',
            'AdjPTIAnnual', 'GAAPPTIAnnual', 'EquityAmt',
            'BeginPrinBal', 'EndPrinBal',
            'RecapBal', 'PIFBal', 'CohortRecap',
        ]
        return agg[keep]

    # ── step 4: volume summary  (unchanged) ──────────────────────────────────
    def build_volume_summary_all(self, begin_fund_dt: date, end_fund_dt: date) -> pd.DataFrame:
        sql = text("""
            SELECT a.Client, a.PropState, a.ProductCategory, a.FICO, a.LTV, a.DTI,
                   a.Rate, a.FundedLoanAmt, a.Channel
            FROM adm.ODS.vwOrigLoan a
            WHERE a.Channel IN ('Correspondent', 'Co-Issue')
              AND a.Client IS NOT NULL
              AND a.FundDt BETWEEN :begin_dt AND :end_dt
              AND a.FundedLoanAmt > 0
              AND a.MortgageType <> 'Reverse'
              AND NOT EXISTS (
                  SELECT 1 FROM Analytics_Rpt.md.B2BPricingHist h
                  WHERE h.LoanNum = a.LoanNum
                    AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
              )
        """)
        df = pd.read_sql(sql, self.engine, params={'begin_dt': begin_fund_dt, 'end_dt': end_fund_dt})

        df = BandClassifier.add_bands(df)
        df['IsGovt'] = df['ProductCategory'].isin(['FHA', 'VA', 'USDA']).astype('int64')

        fl = df['FundedLoanAmt'].astype('float64').fillna(0)
        for col, lo, hi in (('FICO', 580, 850), ('LTV', 0.10, 0.99),
                            ('DTI', 0.05, 0.65), ('Rate', 0.02, 0.10)):
            vals         = pd.to_numeric(df[col], errors='coerce').fillna(0)
            mask         = vals.between(lo, hi)
            df[f'_w_{col}'] = fl.where(mask, 0)
            df[f'_n_{col}'] = (vals * fl).where(mask, 0)

        group_cols = ['Client', 'PropState', 'ProductCategory',
                      'FICOBand', 'LTVBand', 'DTIBand', 'IntRateBand', 'ALB', 'Channel']

        agg = df.groupby(group_cols, dropna=False, sort=False, observed=True).agg(
            LoanCt       =('FundedLoanAmt', 'size'),
            FundedLoanAmt=('FundedLoanAmt', 'sum'),
            _IsGovt      =('IsGovt',        'sum'),
            _n_FICO=('_n_FICO','sum'), _w_FICO=('_w_FICO','sum'),
            _n_LTV =('_n_LTV', 'sum'), _w_LTV =('_w_LTV', 'sum'),
            _n_DTI =('_n_DTI', 'sum'), _w_DTI =('_w_DTI', 'sum'),
            _n_Rate=('_n_Rate','sum'), _w_Rate=('_w_Rate','sum'),
        ).reset_index()

        with np.errstate(divide='ignore', invalid='ignore'):
            agg['GovtPct'] = np.where(agg['LoanCt'] > 0, agg['_IsGovt'] / agg['LoanCt'], 0.0)
            agg['WAFICO']  = np.where(agg['_w_FICO'] > 0, agg['_n_FICO'] / agg['_w_FICO'], np.nan)
            agg['WALTV']   = np.where(agg['_w_LTV']  > 0, agg['_n_LTV']  / agg['_w_LTV'],  np.nan)
            agg['WADTI']   = np.where(agg['_w_DTI']  > 0, agg['_n_DTI']  / agg['_w_DTI'],  np.nan)
            agg['WAC']     = np.where(agg['_w_Rate'] > 0, agg['_n_Rate'] / agg['_w_Rate'], np.nan)

        out = agg[group_cols + ['GovtPct', 'WAFICO', 'WALTV', 'WADTI', 'WAC',
                                'LoanCt', 'FundedLoanAmt']].copy()
        out['BeginFundDt'] = begin_fund_dt
        out['EndFundDt']   = end_fund_dt
        return out

    # ── step 5: load all source frames ───────────────────────────────────────
    def load(self,
             report_as_of_dt: Optional[date] = None,
             begin_score_dt:  Optional[date] = None,
             end_score_dt:    Optional[date] = None,
             begin_fund_dt:   Optional[date] = None,
             end_fund_dt:     Optional[date] = None,
             *,
             reference_today: Optional[date] = None) -> 'B2BPricingReporter':
        """
        Load all source data frames.

        Parameters
        ----------
        report_as_of_dt : date, optional
            Pass a month-end date (e.g. date(2026, 1, 31)) to produce a
            historical report.  All date windows are derived automatically:
              - begin_fund_dt = first day of 13 months prior to report_as_of_dt
              - end_fund_dt   = report_as_of_dt
              - end_score_dt  = report_as_of_dt  (caps performance data)
            Leave None to run for the current latest month (default behaviour).

        All other date parameters can still override the auto-derived values.
        """
        self.report_as_of_dt = report_as_of_dt

        if report_as_of_dt is not None:
            # Historical mode: treat the day after report date as "today"
            ref           = DateHelpers.add_days(report_as_of_dt, 1)
            end_score_dt  = end_score_dt  or report_as_of_dt
            end_fund_dt   = end_fund_dt   or report_as_of_dt
        else:
            ref = reference_today or date.today()

        self.begin_fund_dt  = begin_fund_dt or DateHelpers.add_days(DateHelpers.eomonth(ref, -13), 1)
        self.end_fund_dt    = end_fund_dt   or DateHelpers.eomonth(ref, -1)
        self.begin_score_dt = begin_score_dt
        self.end_score_dt   = end_score_dt

        print(f"Report as of  : {report_as_of_dt or 'current'}")
        print(f"Funding window: {self.begin_fund_dt}  →  {self.end_fund_dt}")
        print(f"Score window  : {self.begin_score_dt or '1900-01-01'}  →  {self.end_score_dt or 'all history'}")

        self.client_seller   = self.build_client_seller_id_ae(self.begin_fund_dt, self.end_fund_dt)
        self.output_template = self.build_output_template(self.begin_fund_dt, self.end_fund_dt)
        self.perform_all     = self.build_perform_summary_all(begin_score_dt, end_score_dt)
        self.volume_all      = self.build_volume_summary_all(self.begin_fund_dt, self.end_fund_dt)
        return self

    # ── cohort CPR scoring helpers  ───────────────────────────────────────────
    @staticmethod
    def _cohort_cpr_score_by_keys(perform_df: pd.DataFrame,
                                  keys: List[str],
                                  min_loans: int = MIN_COHORT_LOANS) -> pd.DataFrame:
        """
        Cohort-based CPR scoring grouped by arbitrary column keys.

        Algorithm
        ---------
         1. For each keys-group, roll all cohort balances together:
             ActualCPR = 1 - (ΣEndBal / ΣBeginBal)^12

         2. For each Cohort_ID (pooled across ALL groups):
             CohortModelCPR = 1 - (ΣEndBal / ΣBeginBal)^12

           3. For each keys-group:
               ModelCPR = 1 - (ΣCohortEndBal / ΣCohortBeginBal)^12
               CPRScore  = ActualCPR / ModelCPR

        PLS / Cohort_ID=NULL are excluded.  Groups with no valid cohort data
        receive CPRScore = 1.0 (neutral default).

        Returns DataFrame with columns: keys + [CPRLoanct, ActualCPR, ModelCPR, CPRScore]
        """
        p = perform_df[perform_df['LoanCt'] > 0].copy()

        # Full-group ActualCPR across all assigned cohorts.
        all_g = p.groupby(keys, dropna=False).agg(
            LoanCt=('LoanCt', 'sum'),
            CPRLoanct=('UniqueLoanCt', 'sum'),
            BeginBal =('BeginPrinBal',  'sum'),
            EndBal   =('EndPrinBal',    'sum'),
        ).reset_index()
        all_g['ActualCPR_all'] = np.where(
            all_g['BeginBal'] > 0,
            1 - (all_g['EndBal'] / all_g['BeginBal']) ** 12,
            np.nan,
        )

        # Restrict to cohort-assigned loans for cohort benchmarking.
        p_c = p[p['Cohort_ID'].notna()].copy()
        if p_c.empty:
            all_g['ActualCPR'] = all_g['ActualCPR_all']
            all_g['ModelCPR']  = np.nan
            all_g['CPRScore']  = 1.0
            return all_g[keys + ['CPRLoanct', 'ActualCPR', 'ModelCPR', 'CPRScore']]

        client_balances = (p_c.groupby(keys, dropna=False)
                            .agg(ClientBeginBal=('BeginPrinBal', 'sum'),
                                 ClientEndBal=('EndPrinBal', 'sum'))
                            .reset_index())
        client_balances['ActualCPR_c'] = np.where(
            client_balances['ClientBeginBal'] > 0,
            1 - (client_balances['ClientEndBal'] / client_balances['ClientBeginBal']) ** 12,
            np.nan,
        )

        # ── Step 1: aggregate each key/cohort cell's balances ────────────────
        gc_keys = keys + ['Cohort_ID']
        gc = p_c.groupby(gc_keys, dropna=False).agg(
            LoanCt_gc=('LoanCt',       'sum'),
            BeginBal_gc=('BeginPrinBal', 'sum'),
            EndBal_gc  =('EndPrinBal',   'sum'),
        ).reset_index()

        # ── Step 2: pooled-balance ModelCPR per cohort ────────────────────────
        cohort_totals = (gc.groupby('Cohort_ID', dropna=False)
                           .agg(CohortBeginBal=('BeginBal_gc', 'sum'),
                                CohortEndBal=('EndBal_gc', 'sum'))
                           .reset_index())
        cohort_totals['CohortModelCPR_cohort'] = np.where(
            cohort_totals['CohortBeginBal'] > 0,
            1 - (cohort_totals['CohortEndBal'] / cohort_totals['CohortBeginBal']) ** 12,
            np.nan,
        )
        gc = gc.merge(cohort_totals[['Cohort_ID', 'CohortModelCPR_cohort']], on='Cohort_ID', how='left')
        valid = gc.dropna(subset=['CohortModelCPR_cohort'])
        if valid.empty:
            all_g['ActualCPR'] = all_g['ActualCPR_all']
            all_g['ModelCPR']  = np.nan
            all_g['CPRScore']  = 1.0
            return all_g[keys + ['CPRLoanct', 'ActualCPR', 'ModelCPR', 'CPRScore']]

        # ── Step 3: roll the cohort benchmarks up to the keys level ──────────
        gc = gc.merge(
            cohort_totals[['Cohort_ID', 'CohortBeginBal', 'CohortEndBal']],
            on='Cohort_ID', how='left')
        score = (gc.groupby(keys, dropna=False, as_index=False)
                   .agg(ModelBeginBal=('CohortBeginBal', 'sum'),
                        ModelEndBal=('CohortEndBal', 'sum')))
        score['ModelCPR_c'] = np.where(
            score['ModelBeginBal'] > 0,
            1 - (score['ModelEndBal'] / score['ModelBeginBal']) ** 12,
            np.nan,
        )
        score = score.merge(client_balances[keys + ['ActualCPR_c']], on=keys, how='left')
        score['CPRScore_c'] = np.where(
            (score['ModelCPR_c'] > 0) & score['ModelCPR_c'].notna(),
            score['ActualCPR_c'] / score['ModelCPR_c'],
            1.0,
        )

        # Merge back — groups with no cohort data get ActualCPR_all and CPRScore=1.0
        out = all_g.merge(
            score[keys + ['ActualCPR_c', 'ModelCPR_c', 'CPRScore_c']],
            on=keys, how='left',
        )
        out['ActualCPR'] = out['ActualCPR_c'].combine_first(out['ActualCPR_all'])
        out['ModelCPR']  = out['ModelCPR_c']
        out['CPRScore']  = np.where(
            out['LoanCt'] < min_loans,
            1.0,
            out['CPRScore_c'].fillna(1.0),
        )
        return out[keys + ['CPRLoanct', 'ActualCPR', 'ModelCPR', 'CPRScore']]

    @staticmethod
    def _cohort_cpr_score(perform_df: pd.DataFrame,
                          metric: str,
                          min_loans: int = MIN_COHORT_LOANS) -> pd.DataFrame:
        """Wrapper: derives MetricValue from `metric` then calls _cohort_cpr_score_by_keys."""
        p = perform_df.copy()
        p['MetricValue'] = _metric_column(p, metric)
        result = B2BPricingReporter._cohort_cpr_score_by_keys(p, ['MetricValue'], min_loans)
        return result  # columns: MetricValue, CPRLoanct, ActualCPR, ModelCPR, CPRScore

    # ── internal helpers  (unchanged) ────────────────────────────────────────
    @staticmethod
    def _total_score(row) -> float:
        if row['CPRScore'] == 0:
            return row['RecapScore']
        if row['RecapScore'] == 0:
            return (1 / row['CPRScore']) * 0.5
        return (1 / row['CPRScore']) * row['RecapScore']

    @staticmethod
    def _round_sql_like(df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        round_map = {
            'FundedLoanAmt': 2, 'GovtPct': 6, 'WAFICO': 6, 'WALTV': 6,
            'WADTI': 6, 'WAC': 6, 'ModelCPR': 6, 'ActualCPR': 6,
            'CPRScore': 6, 'CohortRecap': 6, 'ActualRecap': 6,
            'RecapScore': 6, 'TotalScore': 6,
        }
        for col, digits in round_map.items():
            if col in out.columns:
                out[col] = pd.to_numeric(out[col], errors='coerce').round(digits)
        return out

    def _ensure_loaded(self):
        if any(x is None for x in (self.volume_all, self.perform_all,
                                   self.output_template, self.client_seller)):
            raise RuntimeError("Reporter not loaded. Call .load(...) first.")

    # ── per-metric summary  ───────────────────────────────────────────────────
    def _build_metric_summary(self,
                              volume_df: pd.DataFrame,
                              perform_df: pd.DataFrame,
                              metric: str,
                              product: str) -> pd.DataFrame:
        # ── Volume side (unchanged) ───────────────────────────────────────────
        v = volume_df.copy()
        v['MetricValue'] = _metric_column(v, metric)

        def _agg_vol(g):
            funded = g['FundedLoanAmt'].sum()
            return pd.Series({
                'LoanCt':       g['LoanCt'].sum(),
                'FundedLoanAmt': funded,
                'GovtPct':  WeightedStats.safe_div(
                    (g['GovtPct'] * g['FundedLoanAmt']).sum(), funded) if funded else 0,
                'WAFICO': WeightedStats.wavg(g['WAFICO'], g['FundedLoanAmt'], 580,  850),
                'WALTV':  WeightedStats.wavg(g['WALTV'],  g['FundedLoanAmt'], 0.10, 0.99),
                'WADTI':  WeightedStats.wavg(g['WADTI'],  g['FundedLoanAmt'], 0.05, 0.65),
                'WAC':    WeightedStats.wavg(g['WAC'],    g['FundedLoanAmt'], 0.02, 0.10),
            })

        vol_cols = ['LoanCt', 'FundedLoanAmt', 'GovtPct', 'WAFICO', 'WALTV', 'WADTI', 'WAC']
        ms = (v.groupby(['MetricValue', 'BeginFundDt', 'EndFundDt'], dropna=False)[vol_cols]
              .apply(_agg_vol).reset_index())
        ms = ms.reindex(columns=['MetricValue', 'BeginFundDt', 'EndFundDt'] + vol_cols)
        ms['ProductCategory'] = product
        ms['MetricName']      = metric
        ms['BeginFundDt']     = self.begin_fund_dt
        ms['EndFundDt']       = self.end_fund_dt
        ms['VolumeRank']      = ms['FundedLoanAmt'].rank(method='first', ascending=False).astype(int)

        # ── CPR — cohort-based (NEW) ──────────────────────────────────────────
        cpr_df = self._cohort_cpr_score(perform_df, metric)

        # ── Recap — unchanged ─────────────────────────────────────────────────
        r = perform_df[
            (perform_df['PIFBal'] > 0) &
            (perform_df['PIFReason'].isin(
                ['Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance']))
        ].copy()
        r['MetricValue'] = _metric_column(r, metric)
        recap_cols = ['LoanCt', 'UniqueLoanCt', 'CohortRecap', 'PIFBal', 'RecapBal']
        recap = r.groupby('MetricValue', dropna=False)[recap_cols].apply(lambda g: pd.Series({
            'LoanCt':      g['LoanCt'].sum(),
            'UniqueLoanCt': g['UniqueLoanCt'].sum(),
            'CohortRecap': WeightedStats.safe_div(
                (g['CohortRecap'] * g['PIFBal']).sum(), g['PIFBal'].sum()),
            'ActualRecap': WeightedStats.safe_div(g['RecapBal'].sum(), g['PIFBal'].sum()),
        })).reset_index()
        recap = recap.reindex(columns=['MetricValue', 'LoanCt', 'UniqueLoanCt', 'CohortRecap', 'ActualRecap'])

        # ── Merge and assemble output ─────────────────────────────────────────
        ms = ms.merge(cpr_df, on='MetricValue', how='left')
        ms = ms.merge(
            recap.add_prefix('r_').rename(columns={'r_MetricValue': 'MetricValue'}),
            on='MetricValue', how='left',
        )

        ms['CPRLoanct'] = ms['CPRLoanct'].fillna(0).astype(int)
        ms['CPRScore']  = ms['CPRScore'].fillna(1.0)

        ms['RecapLoanct'] = ms['r_UniqueLoanCt'].fillna(0).astype(int)
        ms['CohortRecap'] = ms['r_CohortRecap']
        ms['ActualRecap'] = ms['r_ActualRecap']
        ms['RecapScore']  = np.where(
            (ms['r_CohortRecap'].fillna(0) == 0) |
            (ms['r_LoanCt'].fillna(0) < 5)       |
            ms['r_ActualRecap'].isna(),
            1.0,
            ms['r_ActualRecap'] / ms['r_CohortRecap'],
        )
        ms['TotalScore'] = ms.apply(self._total_score, axis=1)

        keep = ['ProductCategory', 'MetricName', 'MetricValue', 'BeginFundDt', 'EndFundDt',
                'LoanCt', 'FundedLoanAmt', 'VolumeRank', 'GovtPct', 'WAFICO', 'WALTV',
                'WADTI', 'WAC', 'CPRLoanct', 'ModelCPR', 'ActualCPR', 'CPRScore',
                'RecapLoanct', 'CohortRecap', 'ActualRecap', 'RecapScore', 'TotalScore']
        return (ms[keep]
            .rename(columns={'ModelCPR': 'Cohort_ActualCPR',
                     'ActualCPR': 'Client_ActualCPR'}))

    # ── run all metrics  (unchanged logic, uses new _build_metric_summary) ────
    def run_all_metrics(self,
                        metric_order: Optional[List[str]] = None,
                        product: str = 'All') -> pd.DataFrame:
        self._ensure_loaded()
        metric_order = list(metric_order) if metric_order else list(DEFAULT_METRIC_ORDER)

        if product == 'All':
            v, p = self.volume_all, self.perform_all
        elif product == 'Govt':
            v = self.volume_all[self.volume_all['ProductCategory'].isin(['FHA', 'VA', 'USDA'])]
            p = self.perform_all[self.perform_all['ProductCategory'].isin(['FHA', 'VA', 'USDA'])]
        else:
            v = self.volume_all[self.volume_all['ProductCategory'] == product]
            p = self.perform_all[self.perform_all['ProductCategory'] == product]

        all_rows = []
        for loop_ct, metric in enumerate(metric_order, start=1):
            ms = self._build_metric_summary(v, p, metric, product)

            if metric == 'Client':
                ms = ms.sort_values('VolumeRank').reset_index(drop=True)
                ms['SortOrder'] = np.arange(1, len(ms) + 1)
                ms = ms.merge(self.client_seller, left_on='MetricValue',
                              right_on='Client', how='left')
                ms = ms.drop(columns=['Client'], errors='ignore')
            else:
                tmpl = self.output_template[self.output_template['MetricName'] == metric].copy()
                ms   = tmpl.merge(ms, on=['MetricName', 'MetricValue', 'BeginFundDt', 'EndFundDt'],
                                  how='left')
                ms['ProductCategory']  = product
                ms['LoanCt']           = ms['LoanCt'].fillna(0)
                ms['FundedLoanAmt']    = ms['FundedLoanAmt'].fillna(0)

                if metric == 'Product':
                    order = {'Conventional': 1, 'FHA': 2, 'VA': 3, 'USDA': 4}
                    ms['__sort'] = ms['MetricValue'].map(order).fillna(5)
                    ms = ms.sort_values('__sort').drop(columns='__sort').reset_index(drop=True)
                elif metric == 'Avg Loan Bal':
                    ms['__sort'] = pd.to_numeric(ms['MetricValue'], errors='coerce')
                    ms = ms.sort_values('__sort').drop(columns='__sort').reset_index(drop=True)
                else:
                    ms = ms.sort_values('MetricValue').reset_index(drop=True)

                ms['SortOrder'] = np.arange(1, len(ms) + 1)
                for c in SELLER_COLS:
                    ms[c] = None

            ms['MetricOrder'] = loop_ct
            all_rows.append(ms)

        final = pd.concat(all_rows, ignore_index=True)
        cols  = ['MetricOrder', 'SortOrder', 'ProductCategory', 'MetricName', 'MetricValue',
                 'BeginFundDt', 'EndFundDt', 'LoanCt', 'FundedLoanAmt', 'VolumeRank',
                 'GovtPct', 'WAFICO', 'WALTV', 'WADTI', 'WAC',
                 'CPRLoanct', 'Cohort_ActualCPR', 'Client_ActualCPR', 'CPRScore',
                 'RecapLoanct', 'CohortRecap', 'ActualRecap', 'RecapScore',
                 'TotalScore'] + SELLER_COLS
        final = (final.reindex(columns=cols)
                      .sort_values(['MetricOrder', 'SortOrder'])
                      .reset_index(drop=True))
        return self._round_sql_like(final)

    # ── generic group-by summary  ─────────────────────────────────────────────
    def build_groupby_summary(self, metric_order: List[str]) -> pd.DataFrame:
        """
        Build a multi-key summary (e.g. Client × State, Client × State × Product).
        CPRScore uses the new cohort-based methodology.  RecapScore is unchanged.
        """
        self._ensure_loaded()
        if not metric_order:
            raise ValueError("metric_order must contain at least one metric")

        keys: List[str] = []
        for m in metric_order:
            if m not in METRIC_COLUMN_MAP:
                raise KeyError(f"Unknown metric '{m}'. Valid: {sorted(METRIC_COLUMN_MAP)}")
            col = METRIC_COLUMN_MAP[m]
            if col == '__TOTAL__':
                raise ValueError("'Total' is not valid in build_groupby_summary.")
            keys.append(col)

        # ── Volume side ───────────────────────────────────────────────────────
        def _agg_vol(g):
            funded = g['FundedLoanAmt'].sum()
            return pd.Series({
                'BeginFundDt':  g['BeginFundDt'].iloc[0],
                'EndFundDt':    g['EndFundDt'].iloc[0],
                'LoanCt':       g['LoanCt'].sum(),
                'FundedLoanAmt': funded,
                'GovtPct': WeightedStats.safe_div(
                    (g['GovtPct'] * g['FundedLoanAmt']).sum(), funded) if funded else 0,
                'WAFICO': WeightedStats.wavg(g['WAFICO'], g['FundedLoanAmt'], 580,  850),
                'WALTV':  WeightedStats.wavg(g['WALTV'],  g['FundedLoanAmt'], 0.10, 0.99),
                'WADTI':  WeightedStats.wavg(g['WADTI'],  g['FundedLoanAmt'], 0.05, 0.65),
                'WAC':    WeightedStats.wavg(g['WAC'],    g['FundedLoanAmt'], 0.02, 0.10),
            })

        vol_cols = ['BeginFundDt', 'EndFundDt', 'LoanCt', 'FundedLoanAmt', 'GovtPct', 'WAFICO', 'WALTV', 'WADTI', 'WAC']
        base = (self.volume_all.groupby(keys, dropna=False)[vol_cols]
                .apply(_agg_vol).reset_index())
        base = base.reindex(columns=keys + vol_cols)
        base['VolumeRank'] = base['FundedLoanAmt'].rank(method='first', ascending=False).astype(int)

        # ── CPR — cohort-based (NEW) ──────────────────────────────────────────
        cpr_df = self._cohort_cpr_score_by_keys(self.perform_all, keys)
        out    = base.merge(cpr_df, on=keys, how='left')
        out['CPRLoanct'] = out['CPRLoanct'].fillna(0).astype(int)
        out['CPRScore']  = out['CPRScore'].fillna(1.0)

        # ── Recap — unchanged ─────────────────────────────────────────────────
        r = self.perform_all[
            (self.perform_all['PIFBal'] > 0) &
            (self.perform_all['PIFReason'].isin(
                ['Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance']))
        ]
        recap_cols = ['LoanCt', 'UniqueLoanCt', 'CohortRecap', 'PIFBal', 'RecapBal']
        recap = r.groupby(keys, dropna=False)[recap_cols].apply(lambda g: pd.Series({
            'LoanCt':       g['LoanCt'].sum(),
            'UniqueLoanCt': g['UniqueLoanCt'].sum(),
            'CohortRecap':  WeightedStats.safe_div(
                (g['CohortRecap'] * g['PIFBal']).sum(), g['PIFBal'].sum()),
            'ActualRecap':  WeightedStats.safe_div(g['RecapBal'].sum(), g['PIFBal'].sum()),
        })).reset_index()
        recap = recap.reindex(columns=keys + ['LoanCt', 'UniqueLoanCt', 'CohortRecap', 'ActualRecap'])

        recap_pref = recap.add_prefix('r_')
        if keys == ['Client', 'PropState', 'ProductCategory']:
            out = out.merge(recap_pref,
                            left_on =['Client', 'PropState', 'ProductCategory'],
                            right_on=['r_Client', 'r_PropState', 'r_ProductCategory'],
                            how='left')
        else:
            out = out.merge(recap_pref,
                            left_on =keys,
                            right_on=[f'r_{k}' for k in keys],
                            how='left')

        out['RecapLoanct'] = out['r_UniqueLoanCt'].fillna(0).astype(int)
        out['CohortRecap'] = out['r_CohortRecap']
        out['ActualRecap'] = out['r_ActualRecap']
        out['RecapScore']  = np.where(
            (out['r_CohortRecap'].fillna(0) == 0) | (out['r_LoanCt'].fillna(0) < 5),
            1.0, out['r_ActualRecap'] / out['r_CohortRecap'],
        )
        out['TotalScore'] = out.apply(self._total_score, axis=1)

        # Attach seller info when Client is a key
        if 'Client' in keys and self.client_seller is not None:
            cs = self.client_seller.copy()
            cs['__key']  = cs['Client'].astype(str).str.strip().str.upper()
            out['__key'] = out['Client'].astype(str).str.strip().str.upper()
            out = out.merge(cs[['__key'] + SELLER_COLS], on='__key', how='left').drop(columns='__key')

        drop_cols = [c for c in out.columns if c.startswith('r_')]
        out = out.drop(columns=drop_cols).reset_index(drop=True)

        ordered_cols = keys + [
            'BeginFundDt', 'EndFundDt', 'LoanCt', 'FundedLoanAmt', 'VolumeRank',
            'GovtPct', 'WAFICO', 'WALTV', 'WADTI', 'WAC',
            'CPRLoanct', 'Cohort_ActualCPR', 'Client_ActualCPR', 'CPRScore',
            'RecapLoanct', 'CohortRecap', 'ActualRecap', 'RecapScore', 'TotalScore',
        ] + SELLER_COLS
        out = out.rename(columns={'ModelCPR': 'Cohort_ActualCPR',
                      'ActualCPR': 'Client_ActualCPR'})
        out = out.reindex(columns=[c for c in ordered_cols if c in out.columns])
        return self._round_sql_like(out)
