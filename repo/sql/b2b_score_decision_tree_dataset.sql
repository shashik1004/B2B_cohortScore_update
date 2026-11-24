--DROP TABLE IF EXISTS Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT;
DROP TABLE IF EXISTS Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT_extended;

DECLARE @BeginRefFundDt DATE, @EndRefFundDt DATE
IF @BeginRefFundDt IS NULL SET @BeginRefFundDt = DATEADD(DAY, 1, EOMONTH(GETDATE(), -13))
IF @EndRefFundDt IS NULL SET @EndRefFundDt = EOMONTH(GETDATE(), -1)

select distinct
	o.asofdt
	, o.FundDt
	, loanid
	, InvestorLoanNum
	, o.Client
	, s.TotalScore as B2B_score
	, s.RecapScore
	, s.CPRScore
	, s.CohortRecap
	, s.ActualRecap
	, s.ModelCPR
	, s.ActualCPR
	, Channel
	, SubChannel
	, CASE
		WHEN AmortizationTerm IS NULL THEN 'z_Unknown'
		WHEN AmortizationTerm <= 0 THEN 'z_Unknown'
		WHEN AmortizationTerm < 180 THEN '<15yr'
		WHEN AmortizationTerm = 180 THEN '15yr'
		WHEN AmortizationTerm = 240 THEN '20yr'
		WHEN AmortizationTerm = 300 THEN '25yr'
		WHEN AmortizationTerm = 360 THEN '30yr'
		WHEN AmortizationTerm > 360 THEN '>30yr'
		ELSE 'z_Unknown'
		END AS AmortizationTerm_Band
	, AmortizationType
	, case when ARMIndex is null then 'z_Unknown'
		else 'Known'
		end as ARMIndex_Band
	, BaseLoanAmt
	, TotalLoanAmt
	, CLTV 
	, DTI
	, FICO 
    -- Loan investor (credit / guarantor)
	,  CASE
        WHEN Investor IS NULL THEN 'Unknown'
        WHEN Investor IN ('FANNIE MAE', 'FNMA', 'F N M A')
             OR Investor LIKE '%FANNIE%'
             OR Investor LIKE '%FNMA%'
        THEN 'Fannie Mae'
        WHEN Investor LIKE '%FREDDIE%'
             OR Investor LIKE '%FHLMC%'
        THEN 'Freddie Mac'
        WHEN Investor LIKE '%GINNIE%'
             OR Investor LIKE '%GNMA%'
        THEN 'Ginnie Mae'
        ELSE 'Non-Agency'
    END AS LoanInvestor,
    -- MSR owner
    CASE
        WHEN Investor IS NULL THEN 'Unknown'
        WHEN Investor LIKE '%NRZ%' THEN 'NRZ (Rithm)'
        WHEN Investor LIKE '%ONITY%' THEN 'Onity'
        WHEN Investor LIKE '%PHH%' THEN 'PHH'
        WHEN Investor LIKE '%OCWEN%' THEN 'Ocwen'
        WHEN Investor LIKE '%WELLS%' THEN 'Wells Fargo'
        ELSE 'Other / Unknown'
    END AS MSROwner,
	CASE
		WHEN LoanPurpose IS NULL OR LTRIM(RTRIM(LoanPurpose)) = '' THEN 'Other / Unknown'
		-- Purchase
		WHEN LoanPurpose = 'Purchase' THEN 'Purchase'
		-- Cash-Out Refinance
		WHEN LoanPurpose IN ('Cash-Out Refi', 'Cashout Refinance') OR LoanPurpose LIKE '%EQUITY TAKEOUT%' THEN 'Cash-Out Refi'
		-- Rate/Term Refinance (no cash out)
		WHEN LoanPurpose IN ('Rate/Term Refi', 'No Cash-Out Refi', 'REFINANCE/NO CASH OUT') OR LoanPurpose = 'Refinance' THEN 'Rate/Term Refi'
		-- Reverse
		WHEN LoanPurpose LIKE '%REVERSE%' THEN 'Reverse'
		-- Construction / Home Improvement
		WHEN LoanPurpose IN ('Construction', 'Home Improvement') THEN 'Construction / Improvement'
		-- Everything else
		ELSE 'Other / Unknown'
	END AS LoanPurposeGroup,
	CASE
		WHEN LoanType IS NULL OR LTRIM(RTRIM(LoanType)) = '' THEN 'Other / Unknown'
		-- Conventional
		WHEN LoanType IN ('Conv', 'Conv PMI') THEN 'Conventional' 
		-- FHA
		WHEN LoanType = 'FHA' THEN 'FHA'
		-- VA
		WHEN LoanType = 'VA' THEN 'VA'
		-- USDA
		WHEN LoanType = 'USDA' THEN 'USDA'
		-- Reverse / HECM (all variants)
		WHEN LoanType LIKE '%HECM%' OR LoanType IN ('Home Keeper', 'Simple 60', 'REM', 'Non-Insured HECM') THEN 'Reverse (HECM)'
		-- HELOC / Equity products
		WHEN LoanType IN ('HELOC', 'Equity IQ - Fixed', 'Equity Plus Advantage', 'Cash Account') THEN 'HELOC / Equity'
		-- Everything else
		ELSE 'Other / Unknown'
	END AS LoanTypeGroup
	, LockDuration
	, LTV
	, CASE
		WHEN OccupancyType IS NULL OR LTRIM(RTRIM(OccupancyType)) = '' THEN 'Other / Unknown'
		-- Primary residence
		WHEN OccupancyType IN ('Primary Residence', 'PRIMARY') THEN 'Primary Residence'
		-- Second home
		WHEN OccupancyType IN ('Second Home', 'Secondary') THEN 'Second Home'
		-- Investment / Non-owner occupied
		WHEN OccupancyType IN ('Investment', 'Non-owner Occupied') THEN 'Investment / Non-Owner'
		-- Unknown or other
		WHEN OccupancyType = 'Unknown' THEN 'Other / Unknown'
		ELSE 'Other / Unknown'
	END AS OccupancyGroup
	, o.ProductCategory
	, CASE
		WHEN PropertyType IS NULL OR LTRIM(RTRIM(PropertyType)) = '' THEN 'Other / Unknown'
		-- Single Family
		WHEN PropertyType IN ('SINGLE FAMILY DETACHED', 'Single Family', '1-4 Family') THEN 'Single Family'
		-- Townhouse / PUD
		WHEN PropertyType IN ('TOWNHOUSE', 'PUD Detached', 'PUD Attached' ) THEN 'Townhouse / PUD'
		-- Condominium / Co-op
		WHEN PropertyType IN ('CONDOMINIUM', 'Condo', 'CO-OP', 'Manufactured Home/Condo/PUD/Coop') THEN 'Condominium / Co-op'
		-- Multi-Family (2–4)
		WHEN PropertyType IN ('MULTI-FAMILY UNIT', '2-4 Family') THEN 'Multi-Family (2–4)'
		-- Manufactured / Mobile
		WHEN PropertyType IN ('MOBILE HOME', 'PREFABRICATED', 'Manufactured Housing') THEN 'Manufactured / Mobile'
		-- Land / Special use
		WHEN PropertyType IN ('LAND AND DEVELOPMENT LOTS', 'Unique Property') THEN 'Land / Special Use'
		-- Everything else
		ELSE 'Other / Unknown'
	END AS PropertyTypeGroup
	, CASE
		WHEN o.PropState IS NULL THEN 'Unknown'
		-- West
		WHEN o.PropState IN ('CA','AZ','CO','NV','UT','WA','OR','ID','MT','WY','AK','HI') THEN 'West'
		-- South
		WHEN o.PropState IN ('TX','FL','GA','NC','VA','TN','SC','AL','OK','AR','LA','MS','KY','WV') THEN 'South'
		-- Midwest
		WHEN o.PropState IN ('IL','OH','MI','MO','IN','WI','MN','IA','KS','NE','SD','ND') THEN 'Midwest'
		-- Northeast
		WHEN o.PropState IN ('NY','NJ','PA','MA','CT','RI','NH','VT','ME') THEN 'Northeast'
		-- Mid-Atlantic
		WHEN o.PropState IN ('MD','DE','DC') THEN 'Mid-Atlantic'
		-- Territory
		WHEN o.PropState = 'PR' THEN 'Territory'
		ELSE 'Unknown'
	END AS StateRegion
	, Region
	, o.PropState
	, SelfEmployedFlag
	, ServicingType
	, sms.balance_bin
	, sms.category
	, sms.origination_category
	, sms.servicing_category
	, sms.servicing_subcategory
	, sms.agency
	, sms.investor_type
	, CASE
		WHEN INVESTOR_NAME IS NULL THEN 'Unknown / Other'
		-- Fannie Mae
		WHEN INVESTOR_NAME LIKE '%FANNIE%' OR INVESTOR_NAME LIKE '%FNMA%' THEN 'Fannie Mae'
		-- Freddie Mac
		WHEN INVESTOR_NAME LIKE '%FREDDIE%' OR INVESTOR_NAME LIKE '%FHLMC%' THEN 'Freddie Mac'
		-- Ginnie Mae
		WHEN INVESTOR_NAME LIKE '%GINNIE%' OR INVESTOR_NAME LIKE '%GNMA%' THEN 'Ginnie Mae'
		-- Agency but economically managed (NRZ, Pingora, USAA, ARC, etc.)
		WHEN INVESTOR_NAME LIKE '%NRZ%' OR INVESTOR_NAME LIKE '%PINGORA%' OR INVESTOR_NAME LIKE '%USAA%' OR INVESTOR_NAME LIKE '%ARC%' OR INVESTOR_NAME LIKE '%CALSTERS%' THEN 'Agency (Blended / MSR-Managed)'
		-- Large banks / aggregators
		WHEN INVESTOR_NAME LIKE '%WELLS%'
			 OR INVESTOR_NAME LIKE '%BANK OF AMERICA%'
			 OR INVESTOR_NAME LIKE '%CITIBANK%'
			 OR INVESTOR_NAME LIKE '%JPMORGAN%'
			 OR INVESTOR_NAME LIKE '%CHASE%'
			 OR INVESTOR_NAME LIKE '%U.S. BANK%'
			 OR INVESTOR_NAME LIKE '%HSBC%'
			 OR INVESTOR_NAME LIKE '%PNC%'
			 OR INVESTOR_NAME LIKE '%SANTANDER%'
			 OR INVESTOR_NAME LIKE '%GOLDMAN%'
			 OR INVESTOR_NAME LIKE '%MORGAN STANLEY%'
		THEN 'Large Bank / Aggregator'
		-- Credit unions & community banks
		WHEN INVESTOR_NAME LIKE '%CREDIT UNION%' OR INVESTOR_NAME LIKE '%FCU%' OR INVESTOR_NAME LIKE '%CU' THEN 'Credit Union / Community Bank'
		-- Everything else = private capital / non-agency
		ELSE 'Private / Non-Agency'
	END AS InvestorGroup
	, sms.buydown_flag
	, sms.buydown_type
	, sms.lien_position
	, sms.interest_type
	, sms.occupany
	, rgn.StateFIPS
	, rgn.RegionName
	, rgn.DivisionName
	, s.productcategory_bin
	, s.BaseLoanAmtBand
	, s.division
	, s.loanct

--INTO Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT
INTO Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT_extended

FROM ADM.ODS.vwOrigLoan as o

--left join Analytics_Rpt.RPT.B2BPricing_ClientMetrics_Snapshot as s
--on o.Client = s.MetricValue
left join Analytics_Rpt.RPT.B2BPricing_client_product_amt_div_Metrics_Snapshot as s
on o.Client = s.client and 
(case when o.ProductCategory in ('FHA', 'VA', 'USDA', 'Govt') then 'Gov' else 'Conv' end) = s.productcategory_bin and 
(case when o.BaseLoanAmt>500000 then '>500K' else '<=500K' end) = s.BaseLoanAmtBand and 
(case when o.Region = 'East' or o.PropState in ('NY', 'SC', 'NH', 'MN') then 'East_SelectedStates' else 'Other' end) = s.division

left join finance_DM.[ODS].[SASModelScores] as sms
on o.LoanNum = sms.LOAN_NUMBER

left join [ADM].[lkp].USCensusRegionDivision as rgn
on rgn.PropState = o.PropState

				WHERE o.Channel IN ('Correspondent', 'Co-Issue')
				AND o.Client IS NOT NULL
				AND o.FundDt IS NOT NULL
				AND o.FundedLoanAmt > 0
				AND o.FundDt BETWEEN @BeginRefFundDt and @EndRefFundDt
				AND o.MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = o.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )
			  and cast(sms.enddt as date) = '12/31/9999';


select
Client
, count(*) as cnt
, LTRIM(STR(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 6, 1)) + '%' AS pct
, COUNT(*) as contribution_size
, CASE 
    WHEN 1.0 / COUNT(*) < 1.0 / 200 THEN COUNT(*)
    ELSE 200
END AS contribution_size_cap
, CAST(SQRT(count(*)) as DECIMAL(12,1)) as contribution_sqrt_size

, CAST(1.0 / COUNT(*) AS DECIMAL(12,5)) AS weight_size
, CASE 
    WHEN 1.0 / COUNT(*) < 1.0 / 200 THEN CAST(1.0 / COUNT(*) AS DECIMAL(12,5))
    ELSE CAST(1.0 / 200 AS DECIMAL(12,5))
END AS weight_size_cap
, CAST(1/SQRT(count(*)) AS DECIMAL(12,5)) as weight_sqrt_size

from Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT
group by Client
order by count(*) desc;

DROP TABLE IF EXISTS Analytics_Rpt.RPT.B2BPricingHist_backup_20260512;
DROP TABLE IF EXISTS Analytics_Rpt.RPT.B2BPricingHist_left_join;

select top 10 * from Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT;


select 
case when FundDt<='2026-03-31' then 'Train sample' else 'Test sample' end as sample
, SubChannel
--, StateRegion
, count(*) as cnt
,concat(
        cast(
            round(count(*) * 100.0 
                  / sum(count(*)) over (
                        partition by case when FundDt <= '2026-03-31' then 'Train sample' else 'Test sample' end
                  ), 2
            ) as decimal(10,2)
        ),
        '%'
    ) as pct
from Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT
group by case when FundDt<='2026-03-31' then 'Train sample' else 'Test sample' end, SubChannel
order by case when FundDt<='2026-03-31' then 'Train sample' else 'Test sample' end, SubChannel;


select 
    DivisionName,
    case when FundDt <= '2026-03-31' then 'Train sample' else 'Test sample' end as sample,
    count(*) as cnt,
    concat(
        cast(
            round(count(*) * 100.0 
                  / sum(count(*)) over (
                        partition by case when FundDt <= '2026-03-31' then 'Train sample' else 'Test sample' end
                  ), 2
            ) as decimal(10,2)
        ),
        '%'
    ) as pct
from Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT_extended
where LoanCt>30
group by 
    DivisionName,
    case when FundDt <= '2026-03-31' then 'Train sample' else 'Test sample' end
order by 
    sample,
    DivisionName;







select count(*) from Analytics_Rpt.RPT.B2BPricing_ClientMetrics_DT_extended
where loanct>30;


select * from Analytics_Rpt.RPT.B2BPricing_ClientMetrics_Snapshot;
select 
	client
	, productcategory_bin
	, baseloanamtband
	, division
	, loanct
	, totalscore as B2B_score
from Analytics_Rpt.RPT.B2BPricing_client_product_amt_div_Metrics_Snapshot
order by totalscore desc; 

select count(distinct client) from Analytics_Rpt.RPT.B2BPricing_client_product_amt_div_Metrics_Snapshot



