USE [Analytics_Rpt]
GO
/****** Object:  StoredProcedure [RPT].[usp_md_B2BPricingReporting]    Script Date: 4/27/2026 6:37:31 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER PROC  [RPT].[usp_md_B2BPricingReporting] (@pi_Product VARCHAR(50) = NULL, @pi_BeginScoreDt DATE = NULL, @pi_EndScoreDt DATE = NULL, @pi_BeginRefFundDt DATE = NULL, @pi_EndRefFundDt DATE = NULL)

AS

/******************************************************************************************************************************************************************

	Origination B2B Pricing
	Created by: Jack Cavanagh 
	Created: 01/13/2026
	Sample: 

		EXEC Rpt.[usp_md_B2BPricingReporting]  @pi_Product =  'All'
        
		UpdateDt		UpdatedBy			UpdateDesc
		--------		------------		-------------------------------------------------------------------------
		1/14/26			Baer				Consolidated metrics into single table, fixed e_30 DTI FICO definition
		1/15/26			Cavanagh			Restructured outputs and added product breakout

******************************************************************************************************************************************************************/

-- NOTE:
-- Analytics_Rpt.md.B2BPricingHist.AsofDt represents the REPORTING month,
-- not the servicing source month.
-- Date parameters passed to this procedure should use reporting month-end dates.

DECLARE @BeginScoreDt DATE, @EndScoreDt DATE, @BeginRefFundDt DATE, @EndRefFundDt DATE, @Metric VARCHAR(50), @Product VARCHAR(50), @LoopCt INT = 1


	SET @Product = @pi_Product

		IF @Product IS NULL SET @Product = 'All'	

	SET @BeginScoreDt = @pi_BeginScoreDt

	SET @EndScoreDt = @pi_EndScoreDt
	
	SET @BeginRefFundDt = @pi_BeginRefFundDt

		IF @BeginRefFundDt IS NULL SET @BeginRefFundDt = DATEADD(DAY, 1, EOMONTH(GETDATE(), -13))

	SET @EndRefFundDt = @pi_EndRefFundDt

		IF @EndRefFundDt IS NULL SET @EndRefFundDt = EOMONTH(GETDATE(), -1)

---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
-- Create a table with one row per client, SourceSellerId by SubChannel, and AEName by Region
		DROP TABLE IF EXISTS #ClientSellerIDAE

-- STEP 1: Select AEName per Client and pivot by Region
;WITH DedupedRegionAE AS (
    SELECT
        o.Client,
        o.Region,
        o.AEName,
        MAX(o.CurrentStatusDt) AS CurrentStatusDt
FROM 
       ADM.ODS.vwOrigLoan as o
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
    GROUP BY
        o.Client,
        o.Region,
        o.AEName
),
OneRegionAE AS (
SELECT
    Client,

    STRING_AGG(
        CASE WHEN Region = 'East' THEN AEName END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS AEName_East,

    STRING_AGG(
        CASE WHEN Region = 'West' THEN AEName END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS AEName_West,

    STRING_AGG(
        CASE WHEN Region = 'Central' THEN AEName END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS AEName_Central,

    STRING_AGG(
        CASE WHEN Region IS NULL THEN AEName END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS AEName_NULL

FROM DedupedRegionAE
GROUP BY Client
)

-- STEP 2: Select SourceSellerID per Client and pivot by SubChannel
,DedupedSubChannelSellerID AS (
    SELECT
        o.Client,
        o.SubChannel,
        o.SourceSellerID,
        MAX(o.CurrentStatusDt) AS CurrentStatusDt
FROM 
       ADM.ODS.vwOrigLoan as o
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
    GROUP BY
        o.Client,
        o.SubChannel,
        o.SourceSellerID
),
OneSubChannelSellerID AS (
SELECT
    Client,

    STRING_AGG(
        CASE WHEN SubChannel = 'Delegated - Best Efforts' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_Delegated - Best Efforts],

    STRING_AGG(
        CASE WHEN SubChannel = 'Delegated - Mandatory' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_Delegated - Mandatory],

    STRING_AGG(
        CASE WHEN SubChannel = 'Direct' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_Direct],

    STRING_AGG(
        CASE WHEN SubChannel = 'FHLMC CRX' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_FHLMC CRX],

    STRING_AGG(
        CASE WHEN SubChannel = 'FNMA SMP' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_FNMA SMP],

    STRING_AGG(
        CASE WHEN SubChannel = 'GNMA PIIT' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_GNMA PIIT],

    STRING_AGG(
        CASE WHEN SubChannel = 'Non-delegated' THEN SourceSellerID END, ', '
    ) WITHIN GROUP (ORDER BY CurrentStatusDt DESC)
        AS [SourceSellerID_Non-delegated]

FROM DedupedSubChannelSellerID
GROUP BY Client
)

-- STEP 3: Join AEName and SourceSellerID in one table
SELECT
	COALESCE(t1.Client, t2.Client) AS Client,
	t1.AEName_East as [Account Manager_East], 
	t1.AEName_West as [Account Manager_West], 
	t1.AEName_Central as [Account Manager_Central], 
	t1.AEName_NULL as [Account Manager_NULL], 
	t2.[SourceSellerID_Delegated - Best Efforts] as [ClientID_Delegated - Best Efforts], 
	t2.[SourceSellerID_Delegated - Mandatory] as [ClientID_Delegated - Mandatory], 
	t2.[SourceSellerID_Direct] as [ClientID_Direct], 
	t2.[SourceSellerID_FHLMC CRX] as [ClientID_FHLMC CRX], 
	t2.[SourceSellerID_FNMA SMP] as [ClientID_FNMA SMP], 
	t2.[SourceSellerID_GNMA PIIT] as [ClientID_GNMA PIIT], 
	t2.[SourceSellerID_Non-delegated] as [ClientID_Non-delegated]
	
INTO #ClientSellerIDAE
FROM OneRegionAE as t1
full outer JOIN OneSubChannelSellerID as t2
    ON t1.Client = t2.Client
;

---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------

		DROP TABLE IF EXISTS #OutputTemplate


				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'Total',
						MetricValue = 'Total'

				INTO #OutputTemplate

				UNION
				
				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'State',
						MetricValue = PropState

				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )
						
				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'FICO',
						MetricValue = CASE WHEN ISNULL(a.FICO, 0) NOT BETWEEN 580 AND 850 THEN 'zUnknown'
										WHEN a.FICO >= 780 THEN 'a_780+'
										WHEN a.FICO >= 740 THEN 'b_740'
										WHEN a.FICO >= 700 THEN 'c_700'
										WHEN a.FICO >= 660 THEN 'd_660'
										WHEN a.FICO >= 620 THEN 'e_620'
										WHEN a.FICO >= 580 THEN 'f_580'
										END

				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )
						
				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'LTV',
						MetricValue = CASE WHEN ISNULL(a.LTV, 0) NOT BETWEEN .10 AND .99 THEN 'zUnknown'
										WHEN a.LTV >= .95 THEN 'a_95+'
										WHEN a.LTV >= .90 THEN 'b_90'
										WHEN a.LTV >= .85 THEN 'c_85'
										WHEN a.LTV >= .80 THEN 'd_80'
										WHEN a.LTV >= .70 THEN 'e_70'
										WHEN a.LTV >= .50 THEN 'f_50'
										ELSE 'g_<50'
										END
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'DTI',
						MetricValue = CASE WHEN ISNULL(a.DTI, 0) NOT BETWEEN .05 AND .65 THEN 'zUnknown'
										WHEN a.DTI >= .50 THEN 'a_50+'
										WHEN a.DTI >= .45 THEN 'b_45'
										WHEN a.DTI >= .40 THEN 'c_40'
										WHEN a.DTI >= .35 THEN 'd_35'
										WHEN a.DTI >= .30 THEN 'e_30'
										WHEN a.DTI >= .25 THEN 'f_25'
										ELSE 'g_<25'
										END
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'Interest Rate',
						MetricValue = CASE WHEN ISNULL(a.Rate, 0) NOT BETWEEN .02 AND .10 THEN 'zUnknown'
										WHEN a.Rate >= .0800 THEN 'a_8.00+'
										WHEN a.Rate >= .0750 THEN 'b_7.50'
										WHEN a.Rate >= .0700 THEN 'c_7.00'
										WHEN a.Rate >= .0675 THEN 'd_6.75'
										WHEN a.Rate >= .0650 THEN 'e_6.50'
										WHEN a.Rate >= .0625 THEN 'f_6.25'
										WHEN a.Rate >= .0600 THEN 'g_6.00'
										WHEN a.Rate >= .0575 THEN 'h_5.75'
										WHEN a.Rate >= .0550 THEN 'i_5.50'
										WHEN a.Rate >= .0500 THEN 'j_5.00'
										WHEN a.Rate >= .0450 THEN 'k_4.50'
										WHEN a.Rate >= .0400 THEN 'l_4.00'
										WHEN a.Rate >= .0350 THEN 'm_3.50'
										WHEN a.Rate >= .0300 THEN 'n_3.00'
										ELSE 'o_<3%'
										END
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'Avg Loan Bal',
						MetricValue = CASE WHEN a.FundedLoanAmt <= 100000 THEN '100000' WHEN a.FundedLoanAmt > 1000000 THEN '1000001' ELSE  CONVERT(VARCHAR(50), CEILING(a.FundedLoanAmt / 50000.0) * 50000.0) END
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'Product',
						MetricValue = a.ProductCategory
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

				UNION

				SELECT DISTINCT	
						BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						MetricName = 'Channel',
						MetricValue = a.Channel
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				AND a.Client IS NOT NULL
				AND a.FundDt IS NOT NULL
				AND a.FundedLoanAmt > 0
				AND MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

--				select * from #OutputTemplate;

---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------



			DROP TABLE IF EXISTS #PerformSummaryAll

			DROP TABLE IF EXISTS #VolumeSummaryAll



				SELECT	a.Client, 
						a.PropState,
						a.AsofDt,
						a.ProductCategory,
						GovtPct = SUM(case when a.ProductCategory IN ('FHA', 'VA', 'USDA') then 1 else 0 END) / SUM(1.0),
						FICOBand = CASE WHEN ISNULL(b.FICO, 0) NOT BETWEEN 580 AND 850 THEN 'zUnknown'
										WHEN b.FICO >= 780 THEN 'a_780+'
										WHEN b.FICO >= 740 THEN 'b_740'
										WHEN b.FICO >= 700 THEN 'c_700'
										WHEN b.FICO >= 660 THEN 'd_660'
										WHEN b.FICO >= 620 THEN 'e_620'
										WHEN b.FICO >= 580 THEN 'f_580'
										END,
						WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(b.FICO, 0) BETWEEN 580 AND 850 THEN b.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(b.FICO, 0) BETWEEN 580 AND 850 THEN b.FundedLoanAmt * b.FICO ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(b.FICO, 0) BETWEEN 580 AND 850 THEN b.FundedLoanAmt ELSE 0 END) END,
							
						LTVBand = CASE WHEN ISNULL(b.LTV, 0) NOT BETWEEN .10 AND .99 THEN 'zUnknown'
										WHEN b.LTV >= .95 THEN 'a_95+'
										WHEN b.LTV >= .90 THEN 'b_90'
										WHEN b.LTV >= .85 THEN 'c_85'
										WHEN b.LTV >= .80 THEN 'd_80'
										WHEN b.LTV >= .70 THEN 'e_70'
										WHEN b.LTV >= .50 THEN 'f_50'
										ELSE 'g_<50'
										END,

						WALTV = CASE WHEN SUM(CASE WHEN ISNULL(b.LTV, 0) BETWEEN .10 AND .99 THEN b.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(b.LTV, 0) BETWEEN .10 AND .99 THEN b.FundedLoanAmt * b.LTV ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(b.LTV, 0) BETWEEN .10 AND .99 THEN b.FundedLoanAmt ELSE 0 END) END,

						DTIBand = CASE WHEN ISNULL(b.DTI, 0) NOT BETWEEN .05 AND .65 THEN 'zUnknown'
										WHEN b.DTI >= .50 THEN 'a_50+'
										WHEN b.DTI >= .45 THEN 'b_45'
										WHEN b.DTI >= .40 THEN 'c_40'
										WHEN b.DTI >= .35 THEN 'd_35'
										WHEN b.DTI >= .30 THEN 'e_30'
										WHEN b.DTI >= .25 THEN 'f_25'
										ELSE 'g_<25'
										END,

						WADTI = CASE WHEN SUM(CASE WHEN ISNULL(b.DTI, 0) BETWEEN .05 AND .65 THEN b.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(b.DTI, 0) BETWEEN .05 AND .65 THEN b.FundedLoanAmt * b.DTI ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(b.DTI, 0) BETWEEN .05 AND .65 THEN b.FundedLoanAmt ELSE 0 END) END,

						IntRateBand = CASE WHEN ISNULL(b.Rate, 0) NOT BETWEEN .02 AND .10 THEN 'zUnknown'
										WHEN b.Rate >= .0800 THEN 'a_8.00+'
										WHEN b.Rate >= .0750 THEN 'b_7.50'
										WHEN b.Rate >= .0700 THEN 'c_7.00'
										WHEN b.Rate >= .0675 THEN 'd_6.75'
										WHEN b.Rate >= .0650 THEN 'e_6.50'
										WHEN b.Rate >= .0625 THEN 'f_6.25'
										WHEN b.Rate >= .0600 THEN 'g_6.00'
										WHEN b.Rate >= .0575 THEN 'h_5.75'
										WHEN b.Rate >= .0550 THEN 'i_5.50'
										WHEN b.Rate >= .0500 THEN 'j_5.00'
										WHEN b.Rate >= .0450 THEN 'k_4.50'
										WHEN b.Rate >= .0400 THEN 'l_4.00'
										WHEN b.Rate >= .0350 THEN 'm_3.50'
										WHEN b.Rate >= .0300 THEN 'n_3.00'
										ELSE 'o_<3%'
										END,

						WAC = CASE WHEN SUM(CASE WHEN ISNULL(b.Rate, 0) BETWEEN .02 AND .10 THEN b.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(b.Rate, 0) BETWEEN .02 AND .10 THEN b.FundedLoanAmt * b.Rate ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(b.Rate, 0) BETWEEN .02 AND .10 THEN b.FundedLoanAmt ELSE 0 END) END,

						ALB = CASE WHEN b.FundedLoanAmt <= 100000 THEN '100000' WHEN b.FundedLoanAmt > 1000000 THEN '1000001' ELSE  CONVERT(VARCHAR(50), CEILING(b.FundedLoanAmt / 50000.0) * 50000.0) end,

				
					  LoanCt = COUNT(*),
					  UniqueLoanCt = COUNT(DISTINCT a.LoanNum), 
					  FundedLoanAmt = SUM(a.FundedLoanAmt),

					  AdjPTIAnnual = SUM(a.ServAdjPTI * 12.0 + a.CDAdjPTI),
					  GAAPPTIAnnual = SUM(a.ServGAAPPTI * 12.0 + a.CDAdjPTI),
					  EquityAmt = SUM(a.EquityAmt),

					  BeginPrinBal = SUM(a.BeginPrinBal),
					  EndPrinBal = SUM(a.EndPrinBal),
					  ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal),


					  a.PIFReason,
					  RecapBal = SUM(a.RecapBal),
					  PIFBal = SUM(a.PIFBal),
					  CohortRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal) END,

						b.Channel	


				INTO #PerformSummaryAll
				FROM Analytics_Rpt.md.B2BPricingHist a
		
					INNER JOIN ADM.ODS.vwOrigLoan b ON a.LoanNum = b.LoanNum

				Where a.FundedLoanAmt > 0
				and a.AsofDt BETWEEN ISNULL(@BeginScoreDt, '1/1/1900') and ISNULL(@EndScoreDt, '12/31/9999')
				AND b.MortgageType <> 'Reverse'
			  -- EXCLUDE any loan that EVER XFERed
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist x
					WHERE x.LoanNum = a.LoanNum
					  AND (x.ResolType = 'XFER' OR x.Channel = 'Bulk')
			  )

				GROUP BY 
						a.Client, 
						a.PropState,
						a.AsofDt,
						a.ProductCategory,
						CASE WHEN ISNULL(b.FICO, 0) NOT BETWEEN 580 AND 850 THEN 'zUnknown'
										WHEN b.FICO >= 780 THEN 'a_780+'
										WHEN b.FICO >= 740 THEN 'b_740'
										WHEN b.FICO >= 700 THEN 'c_700'
										WHEN b.FICO >= 660 THEN 'd_660'
										WHEN b.FICO >= 620 THEN 'e_620'
										WHEN b.FICO >= 580 THEN 'f_580'
										END,

						CASE WHEN ISNULL(b.LTV, 0) NOT BETWEEN .10 AND .99 THEN 'zUnknown'
										WHEN b.LTV >= .95 THEN 'a_95+'
										WHEN b.LTV >= .90 THEN 'b_90'
										WHEN b.LTV >= .85 THEN 'c_85'
										WHEN b.LTV >= .80 THEN 'd_80'
										WHEN b.LTV >= .70 THEN 'e_70'
										WHEN b.LTV >= .50 THEN 'f_50'
										ELSE 'g_<50'
										END,
						CASE WHEN ISNULL(b.DTI, 0) NOT BETWEEN .05 AND .65 THEN 'zUnknown'
										WHEN b.DTI >= .50 THEN 'a_50+'
										WHEN b.DTI >= .45 THEN 'b_45'
										WHEN b.DTI >= .40 THEN 'c_40'
										WHEN b.DTI >= .35 THEN 'd_35'
										WHEN b.DTI >= .30 THEN 'e_30'
										WHEN b.DTI >= .25 THEN 'f_25'
										ELSE 'g_<25'
										END,

						CASE WHEN ISNULL(b.Rate, 0) NOT BETWEEN .02 AND .10 THEN 'zUnknown'
										WHEN b.Rate >= .0800 THEN 'a_8.00+'
										WHEN b.Rate >= .0750 THEN 'b_7.50'
										WHEN b.Rate >= .0700 THEN 'c_7.00'
										WHEN b.Rate >= .0675 THEN 'd_6.75'
										WHEN b.Rate >= .0650 THEN 'e_6.50'
										WHEN b.Rate >= .0625 THEN 'f_6.25'
										WHEN b.Rate >= .0600 THEN 'g_6.00'
										WHEN b.Rate >= .0575 THEN 'h_5.75'
										WHEN b.Rate >= .0550 THEN 'i_5.50'
										WHEN b.Rate >= .0500 THEN 'j_5.00'
										WHEN b.Rate >= .0450 THEN 'k_4.50'
										WHEN b.Rate >= .0400 THEN 'l_4.00'
										WHEN b.Rate >= .0350 THEN 'm_3.50'
										WHEN b.Rate >= .0300 THEN 'n_3.00'
										ELSE 'o_<3%'
										END,
						case when b.FundedLoanAmt <= 100000 THEN '100000' WHEN b.FundedLoanAmt > 1000000 THEN '1000001' else  CONVERT(VARCHAR(50), CEILING(b.FundedLoanAmt / 50000.0) * 50000.0) end,
						a.PIFReason, b.Channel

				SELECT	BeginFundDt = @BeginRefFundDt,
						EndFundDt = @EndRefFundDt,
						a.Client, 
						a.PropState,
						a.ProductCategory,
						GovtPct = SUM(case when a.ProductCategory IN ('FHA', 'VA', 'USDA') then 1 else 0 END) / SUM(1.0),
						FICOBand = CASE WHEN ISNULL(a.FICO, 0) NOT BETWEEN 580 AND 850 THEN 'zUnknown'
										WHEN a.FICO >= 780 THEN 'a_780+'
										WHEN a.FICO >= 740 THEN 'b_740'
										WHEN a.FICO >= 700 THEN 'c_700'
										WHEN a.FICO >= 660 THEN 'd_660'
										WHEN a.FICO >= 620 THEN 'e_620'
										WHEN a.FICO >= 580 THEN 'f_580'
										END,
						WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.FICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(a.FICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.FICO ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(a.FICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							
						LTVBand = CASE WHEN ISNULL(a.LTV, 0) NOT BETWEEN .10 AND .99 THEN 'zUnknown'
										WHEN a.LTV >= .95 THEN 'a_95+'
										WHEN a.LTV >= .90 THEN 'b_90'
										WHEN a.LTV >= .85 THEN 'c_85'
										WHEN a.LTV >= .80 THEN 'd_80'
										WHEN a.LTV >= .70 THEN 'e_70'
										WHEN a.LTV >= .50 THEN 'f_50'
										ELSE 'g_<50'
										END,

						WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.LTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(a.LTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.LTV ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(a.LTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

						DTIBand = CASE WHEN ISNULL(a.DTI, 0) NOT BETWEEN .05 AND .65 THEN 'zUnknown'
										WHEN a.DTI >= .50 THEN 'a_50+'
										WHEN a.DTI >= .45 THEN 'b_45'
										WHEN a.DTI >= .40 THEN 'c_40'
										WHEN a.DTI >= .35 THEN 'd_35'
										WHEN a.DTI >= .30 THEN 'e_30'
										WHEN a.DTI >= .25 THEN 'f_25'
										ELSE 'g_<25'
										END,

						WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.DTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(a.DTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.DTI ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(a.DTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,

						IntRateBand = CASE WHEN ISNULL(a.Rate, 0) NOT BETWEEN .02 AND .10 THEN 'zUnknown'
										WHEN a.Rate >= .0800 THEN 'a_8.00+'
										WHEN a.Rate >= .0750 THEN 'b_7.50'
										WHEN a.Rate >= .0700 THEN 'c_7.00'
										WHEN a.Rate >= .0675 THEN 'd_6.75'
										WHEN a.Rate >= .0650 THEN 'e_6.50'
										WHEN a.Rate >= .0625 THEN 'f_6.25'
										WHEN a.Rate >= .0600 THEN 'g_6.00'
										WHEN a.Rate >= .0575 THEN 'h_5.75'
										WHEN a.Rate >= .0550 THEN 'i_5.50'
										WHEN a.Rate >= .0500 THEN 'j_5.00'
										WHEN a.Rate >= .0450 THEN 'k_4.50'
										WHEN a.Rate >= .0400 THEN 'l_4.00'
										WHEN a.Rate >= .0350 THEN 'm_3.50'
										WHEN a.Rate >= .0300 THEN 'n_3.00'
										ELSE 'o_<3%'
										END,

						WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.Rate, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
										SUM(CASE WHEN ISNULL(a.Rate, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.Rate ELSE 0 END) 
											/ SUM(CASE WHEN ISNULL(a.Rate, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,

						ALB = case when a.FundedLoanAmt <= 100000 THEN '100000' WHEN a.FundedLoanAmt > 1000000 THEN '1000001' else  CONVERT(VARCHAR(50), CEILING(a.FundedLoanAmt / 50000.0) * 50000.0) END,

					  LoanCt = COUNT(*),
					  FundedLoanAmt = SUM(a.FundedLoanAmt), 
					  a.Channel

				INTO #VolumeSummaryAll
				FROM adm.ODS.vwOrigLoan a
				WHERE a.Channel IN ('Correspondent', 'Co-Issue')
				and a.Client is not null
				AND a.FundDt BETWEEN @BeginRefFundDt and @EndRefFundDt
				and a.FundedLoanAmt > 0
				AND a.MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

				GROUP BY a.Client, 
						a.PropState,
						a.ProductCategory,
						CASE WHEN ISNULL(a.FICO, 0) NOT BETWEEN 580 AND 850 THEN 'zUnknown'
										WHEN a.FICO >= 780 THEN 'a_780+'
										WHEN a.FICO >= 740 THEN 'b_740'
										WHEN a.FICO >= 700 THEN 'c_700'
										WHEN a.FICO >= 660 THEN 'd_660'
										WHEN a.FICO >= 620 THEN 'e_620'
										WHEN a.FICO >= 580 THEN 'f_580'
										END,

						CASE WHEN ISNULL(a.LTV, 0) NOT BETWEEN .10 AND .99 THEN 'zUnknown'
										WHEN a.LTV >= .95 THEN 'a_95+'
										WHEN a.LTV >= .90 THEN 'b_90'
										WHEN a.LTV >= .85 THEN 'c_85'
										WHEN a.LTV >= .80 THEN 'd_80'
										WHEN a.LTV >= .70 THEN 'e_70'
										WHEN a.LTV >= .50 THEN 'f_50'
										ELSE 'g_<50'
										END,
						CASE WHEN ISNULL(a.DTI, 0) NOT BETWEEN .05 AND .65 THEN 'zUnknown'
										WHEN a.DTI >= .50 THEN 'a_50+'
										WHEN a.DTI >= .45 THEN 'b_45'
										WHEN a.DTI >= .40 THEN 'c_40'
										WHEN a.DTI >= .35 THEN 'd_35'
										WHEN a.DTI >= .30 THEN 'e_30'
										WHEN a.DTI >= .25 THEN 'f_25'
										ELSE 'g_<25'
										END,

						CASE WHEN ISNULL(a.Rate, 0) NOT BETWEEN .02 AND .10 THEN 'zUnknown'
										WHEN a.Rate >= .0800 THEN 'a_8.00+'
										WHEN a.Rate >= .0750 THEN 'b_7.50'
										WHEN a.Rate >= .0700 THEN 'c_7.00'
										WHEN a.Rate >= .0675 THEN 'd_6.75'
										WHEN a.Rate >= .0650 THEN 'e_6.50'
										WHEN a.Rate >= .0625 THEN 'f_6.25'
										WHEN a.Rate >= .0600 THEN 'g_6.00'
										WHEN a.Rate >= .0575 THEN 'h_5.75'
										WHEN a.Rate >= .0550 THEN 'i_5.50'
										WHEN a.Rate >= .0500 THEN 'j_5.00'
										WHEN a.Rate >= .0450 THEN 'k_4.50'
										WHEN a.Rate >= .0400 THEN 'l_4.00'
										WHEN a.Rate >= .0350 THEN 'm_3.50'
										WHEN a.Rate >= .0300 THEN 'n_3.00'
										ELSE 'o_<3%'
										END,
					case when a.FundedLoanAmt <= 100000 THEN '100000' WHEN a.FundedLoanAmt > 1000000 THEN '1000001' else  CONVERT(VARCHAR(50), CEILING(a.FundedLoanAmt / 50000.0) * 50000.0) END, 
					a.Channel

--select top 1000 * from #PerformSummaryAll;

--select * from #VolumeSummaryAll;


--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------


	WHILE @Product <> 'Exit'

		BEGIN


			DROP TABLE IF EXISTS #PerformSummary

			DROP TABLE IF EXISTS #VolumeSummary
			

			SELECT	a.*

			INTO #PerformSummary
			FROM #PerformSummaryAll a
			Where (@Product = 'All' or a.ProductCategory = @Product)


--			select top 1000 * from #PerformSummary;



			SELECT	a.*

			INTO #VolumeSummary
			FROM #VolumeSummaryAll a
			WHERE (@Product = 'All' or a.ProductCategory = @Product)


--			select top 1000 * from #VolumeSummary;



			DROP TABLE IF EXISTS #AllMetrics

			CREATE TABLE #AllMetrics (MetricOrder INT, SortOrder INT, ProductCategory VARCHAR(50), MetricName varchar(50), MetricValue VARCHAR(200), BeginFundDt DATE,	EndFundDt DATE, LoanCt INT,	FundedLoanAmt MONEY, VolumeRank INT, GovtPct FLOAT, WAFICO FLOAT, WALTV FLOAT,
								WADTI FLOAT, WAC FLOAT, CPRLoanct Int, ModelCPR FLOAT, ActualCPR FLOAT,CPRScore FLOAT, RecapLoanct Int, CohortRecap FLOAT, ActualRecap FLOAT, RecapScore FLOAT, TotalScore FLOAT, 
								[Account Manager_East] VARCHAR(1000), [Account Manager_West] VARCHAR(1000), [Account Manager_Central] VARCHAR(1000), [Account Manager_NULL] VARCHAR(1000),
								[ClientID_Delegated - Best Efforts] VARCHAR(500), [ClientID_Delegated - Mandatory] VARCHAR(500), [ClientID_Direct] VARCHAR(500), [ClientID_FHLMC CRX] VARCHAR(500), 
								[ClientID_FNMA SMP] VARCHAR(500), [ClientID_GNMA PIIT] VARCHAR(500), [ClientID_Non-delegated] VARCHAR(500)
								)


			set @Metric = 'Total'
	
			WHILE @Metric <> 'Exit'

				BEGIN

						DROP TABLE IF EXISTS #MetricSummary

--						SELECT	TOP 500
						SELECT	
								ProductCategory = @Product,
								MetricName = @Metric,
								MetricValue = CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB WHEN 'Channel' THEN a.Channel END,
								a.BeginFundDt,
								a.EndFundDt,
								LoanCt = SUM(a.LoanCt), 
								FundedLoanAmt = SUM(a.FundedLoanAmt),
								VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),

								GovtPct = CASE WHEN SUM(a.FundedLoanAmt) = 0 THEN 0 ELSE SUM(a.GovtPct * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) end,

								WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.WAFICO ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							

								WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.WALTV ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

								WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.WADTI ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,


								WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.WAC ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,
			
								CPRLoanct = CONVERT(Int, NULL),
								ModelCPR = CONVERT(FLOAT, NULL),
								ActualCPR = CONVERT(FLOAT, NULL),
								CPRScore = CONVERT(FLOAT, 1),

								RecapLoanct = CONVERT(Int, NULL),
								CohortRecap = CONVERT(FLOAT, NULL),
								ActualRecap = CONVERT(FLOAT, NULL),
								RecapScore = CONVERT(FLOAT, 1),

								TotalScore = CONVERT(FLOAT, 1)
						INTO #MetricSummary	
						FROM #VolumeSummary a
						GROUP BY CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB WHEN 'Channel' THEN a.Channel END,
						a.BeginFundDt, a.EndFundDt

					;WITH Prepay AS
									(SELECT	MetricValue = CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB WHEN 'Channel' THEN a.Channel END, 
											LoanCt = SUM(a.LoanCt), 
											UniqueLoanCt = SUM(a.UniqueLoanCt), 

											AdjROE = CASE WHEN SUM(a.EquityAmt) = 0 then NULL ELSE SUM(a.AdjPTIAnnual) / SUM(a.EquityAmt) end, 
											GAAPROE = CASE WHEN SUM(a.EquityAmt) = 0 then NULL ELSE SUM(a.GAAPPTIAnnual) / SUM(a.EquityAmt) END,

											ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal), 
											ActualCPR = CASE WHEN SUM(a.BeginPrinBal) = 0 THEN NULL ELSE 1 - (POWER(SUM(a.EndPrinBal) / SUM(a.BeginPrinBal), 12)) END
									FROM #PerformSummary a
									where a.LoanCt > 0
									GROUP BY CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB WHEN 'Channel' THEN a.Channel END),

						Recap AS
									(SELECT	MetricValue = CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB WHEN 'Channel' THEN a.Channel END, 
											LoanCt = SUM(a.LoanCt), 
											UniqueLoanCt = SUM(a.UniqueLoanCt), 

											CohortRecap = SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal), 
											ActualRecap = SUM(a.RecapBal) / SUM(a.PIFBal)
									FROM #PerformSummary a
									where a.PIFBal > 0
									AND a.PIFReason IN ('Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance') 
									GROUP BY CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB WHEN 'Channel' THEN a.Channel END)

--						select * from Prepay;
--						select * from Recap;


					UPDATE a
					SET 
						a.CPRLoanct = ISNULL(b.UniqueLoanCt,0),
						a.ModelCPR = b.ModelCPR,
						a.ActualCPR = b.ActualCPR,
						a.CPRScore =  CASE when ISNULL(b.ModelCPR, 0) = 0 OR b.LoanCt < 10 THEN 1 ELSE b.ActualCPR / b.ModelCPR end,
		
						a.RecapLoanct = ISNULL(c.UniqueLoanCt,0),
						a.CohortRecap = c.CohortRecap,
						a.ActualRecap = c.ActualRecap,
						a.RecapScore = case WHEN ISNULL(c.CohortRecap, 0) = 0 or ISNULL(c.LoanCt, 0) < 5 THEN 1 else c.ActualRecap / c.CohortRecap end
					FROM #MetricSummary a

						LEFT JOIN Prepay b ON a.MetricValue = b.MetricValue

						LEFT JOIN Recap c ON a.MetricValue = c.MetricValue

					UPDATE a
					SET a.TotalScore = CASE WHEN a.CPRScore = 0 THEN a.RecapScore * .5  WHEN a.RecapScore = 0 THEN (1 / a.CPRScore) * .5 ELSE (1 / a.CPRScore) * a.RecapScore end
					FROM #MetricSummary a

--					select * from #MetricSummary;

					IF @Metric = 'Client'  

--						INSERT INTO #AllMetrics
--						SELECT TOP 500 MetricOrder = @LoopCt, SortOrder = ROW_NUMBER() OVER (ORDER BY  a.VolumeRank), a.*
--						FROM #MetricSummary a
--						ORDER BY 2

						INSERT INTO #AllMetrics
						SELECT MetricOrder = @LoopCt, SortOrder = ROW_NUMBER() OVER (ORDER BY  a.VolumeRank), a.*, 
								seller.[Account Manager_East], 
								seller.[Account Manager_West], 
								seller.[Account Manager_Central], 
								seller.[Account Manager_NULL], 
								seller.[ClientID_Delegated - Best Efforts], 
								seller.[ClientID_Delegated - Mandatory], 
								seller.[ClientID_Direct], 
								seller.[ClientID_FHLMC CRX], 
								seller.[ClientID_FNMA SMP], 
								seller.[ClientID_GNMA PIIT], 
								seller.[ClientID_Non-delegated]

						FROM #MetricSummary a
						OUTER APPLY (
							select 
								* 
							from #ClientSellerIDAE as x
							WHERE x.Client = a.MetricValue   -- MetricValue contains Client name
						) seller
						ORDER BY 2

					ELSE IF @Metric = 'Product'  

						INSERT INTO #AllMetrics
						SELECT	MetricOrder = @LoopCt, 
								SortOrder = ROW_NUMBER() OVER (ORDER BY CASE a.MetricValue WHEN 'Conventional' then 1 when 'FHA' THEN 2 WHEN 'VA' THEN 3 WHEN 'USDA' THEN 4 else 5 END), 
								ProductCategory = @Product,
								a.MetricName,
								a.MetricValue,
								a.BeginFundDt,
								a.EndFundDt,
								LoanCt = ISNULL(b.LoanCt,0),
								FundedLoanAmt = ISNULL(b.FundedLoanAmt, 0),
								b.VolumeRank,
								b.GovtPct,
								b.WAFICO,
								b.WALTV,
								b.WADTI,
								b.WAC,
								b.CPRLoanct,
								b.ModelCPR,
								b.ActualCPR,
								b.CPRScore,
								b.RecapLoanct,
								b.CohortRecap,
								b.ActualRecap,
								b.RecapScore,
								b.TotalScore, 
								NULL AS [Account Manager_East], 
								NULL AS [Account Manager_West], 
								NULL AS [Account Manager_Central], 
								NULL AS [Account Manager_NULL], 
								NULL AS [ClientID_Delegated - Best Efforts], 
								NULL AS [ClientID_Delegated - Mandatory], 
								NULL AS [ClientID_Direct], 
								NULL AS [ClientID_FHLMC CRX], 
								NULL AS [ClientID_FNMA SMP], 
								NULL AS [ClientID_GNMA PIIT], 
								NULL AS [ClientID_Non-delegated]

						FROM #OutputTemplate  a

								LEFT join  #MetricSummary b on a.MetricName = b.MetricName and a.MetricValue = b.MetricValue

						WHERE a.MetricName = @Metric
						ORDER BY 2
                
					ELSE IF @Metric = 'Avg Loan Bal'  

						INSERT INTO #AllMetrics
						SELECT MetricOrder = @LoopCt, SortOrder = ROW_NUMBER() OVER (ORDER BY  CONVERT(MONEY, a.MetricValue)),
								ProductCategory = @Product,
								a.MetricName,
								a.MetricValue,
								a.BeginFundDt,
								a.EndFundDt,
								LoanCt = ISNULL(b.LoanCt,0),
								FundedLoanAmt = ISNULL(b.FundedLoanAmt, 0),
								b.VolumeRank,
								b.GovtPct,
								b.WAFICO,
								b.WALTV,
								b.WADTI,
								b.WAC,
								b.CPRLoanct,
								b.ModelCPR,
								b.ActualCPR,
								b.CPRScore,
								b.RecapLoanct,
								b.CohortRecap,
								b.ActualRecap,
								b.RecapScore,
								b.TotalScore, 
								NULL AS [Account Manager_East], 
								NULL AS [Account Manager_West], 
								NULL AS [Account Manager_Central], 
								NULL AS [Account Manager_NULL], 
								NULL AS [ClientID_Delegated - Best Efforts], 
								NULL AS [ClientID_Delegated - Mandatory], 
								NULL AS [ClientID_Direct], 
								NULL AS [ClientID_FHLMC CRX], 
								NULL AS [ClientID_FNMA SMP], 
								NULL AS [ClientID_GNMA PIIT], 
								NULL AS [ClientID_Non-delegated]

						FROM #OutputTemplate  a

								LEFT join  #MetricSummary b on a.MetricName = b.MetricName and a.MetricValue = b.MetricValue

						WHERE a.MetricName = @Metric
						ORDER BY 2

					ELSE

						INSERT INTO #AllMetrics
						SELECT MetricOrder = @LoopCt, SortOrder = ROW_NUMBER() OVER (ORDER BY  a.MetricValue),
								ProductCategory = @Product,
								a.MetricName,
								a.MetricValue,
								a.BeginFundDt,
								a.EndFundDt,
								LoanCt = ISNULL(b.LoanCt,0),
								FundedLoanAmt = ISNULL(b.FundedLoanAmt, 0),
								b.VolumeRank,
								b.GovtPct,
								b.WAFICO,
								b.WALTV,
								b.WADTI,
								b.WAC,
								b.CPRLoanct,
								b.ModelCPR,
								b.ActualCPR,
								b.CPRScore,
								b.RecapLoanct,
								b.CohortRecap,
								b.ActualRecap,
								b.RecapScore,
								b.TotalScore, 
								NULL AS [Account Manager_East], 
								NULL AS [Account Manager_West], 
								NULL AS [Account Manager_Central], 
								NULL AS [Account Manager_NULL], 
								NULL AS [ClientID_Delegated - Best Efforts], 
								NULL AS [ClientID_Delegated - Mandatory], 
								NULL AS [ClientID_Direct], 
								NULL AS [ClientID_FHLMC CRX], 
								NULL AS [ClientID_FNMA SMP], 
								NULL AS [ClientID_GNMA PIIT], 
								NULL AS [ClientID_Non-delegated]

						FROM #OutputTemplate  a

								LEFT join  #MetricSummary b on a.MetricName = b.MetricName and a.MetricValue = b.MetricValue

						WHERE a.MetricName = @Metric
						ORDER BY 2

					SET @Metric = CASE @Metric WHEN 'Total' THEN 'State' WHEN 'State' THEN 'FICO' WHEN 'FICO' THEN 'LTV' WHEN 'LTV' THEN 'DTI' WHEN 'DTI' THEN 'Interest Rate' WHEN 'Interest Rate' THEN 'Product' WHEN 'Product' THEN 'Avg Loan Bal' WHEN 'Avg Loan Bal' THEN 'Channel' when 'Channel' then 'Client' when 'Client' then 'Exit' END
            
					SET @LoopCt = @LoopCt + 1

				END

				SELECT a.*
				FROM #AllMetrics a
				ORDER BY 1, 2


-- SAVE CLIENT METRICS FOR PRODUCT = 'ALL' - This will be used for Decision Trees analysis
IF @Product = 'All'
BEGIN
	DROP TABLE IF EXISTS Analytics_Rpt.RPT.B2BPricing_ClientMetrics_Snapshot;
    SELECT
        @EndScoreDt AS AsofDt,
        *
    INTO Analytics_Rpt.RPT.B2BPricing_ClientMetrics_Snapshot
    FROM #AllMetrics
    WHERE MetricName = 'Client';
END

				SET @Product = CASE @Product when 'All' THEN 'Conventional' when 'Conventional' THEN 'FHA' when 'FHA' THEN 'VA' when 'VA' THEN 'USDA' when 'USDA' THEN 'Exit' END

				IF ISNULL(@pi_Product, 'All') NOT IN ('All') set @Product = 'Exit'
                
				SET @LoopCt = 1

		END

------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------
---------------------------------------------------------------------------------
-- Create a table Client x State from #VolumeSummaryAll, using same architecture as AllMetrics
---------------------------------------------------------------------------------
DROP TABLE IF EXISTS #ClientState_VolumeSummaryAll;
---------------------------------------------------------------------------------
-- (1) INITIALIZE #ClientState: GROUP BY Client, State, BeginFundDt, EndFundDt
---------------------------------------------------------------------------------
SELECT  
      a.Client,
      a.PropState,
      a.BeginFundDt,
      a.EndFundDt,

      -- Base aggregated origination + performance metrics
								LoanCt = SUM(a.LoanCt), 
								FundedLoanAmt = SUM(a.FundedLoanAmt),
								VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),

								GovtPct = CASE WHEN SUM(a.FundedLoanAmt) = 0 THEN 0 ELSE SUM(a.GovtPct * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) end,

								WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.WAFICO ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							

								WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.WALTV ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

								WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.WADTI ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,


								WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.WAC ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,
			
								CPRLoanct = CONVERT(Int, NULL),
								ModelCPR = CONVERT(FLOAT, NULL),
								ActualCPR = CONVERT(FLOAT, NULL),
								CPRScore = CONVERT(FLOAT, 1),

								RecapLoanct = CONVERT(Int, NULL),
								CohortRecap = CONVERT(FLOAT, NULL),
								ActualRecap = CONVERT(FLOAT, NULL),
								RecapScore = CONVERT(FLOAT, 1),

								TotalScore = CONVERT(FLOAT, 1)

INTO #ClientState_VolumeSummaryAll
FROM #VolumeSummaryAll a
GROUP BY 
      a.Client,
      a.PropState,
      a.BeginFundDt,
      a.EndFundDt;

---------------------------------------------------------------------------------
-- (2) PREPAY CALCULATIONS (CPR)
---------------------------------------------------------------------------------
;WITH Prepay AS (
    SELECT
         a.Client,
         a.PropState,
											LoanCt = SUM(a.LoanCt), 
											UniqueLoanCt = SUM(a.UniqueLoanCt), 

											AdjROE = CASE WHEN SUM(a.EquityAmt) = 0 then NULL ELSE SUM(a.AdjPTIAnnual) / SUM(a.EquityAmt) end, 
											GAAPROE = CASE WHEN SUM(a.EquityAmt) = 0 then NULL ELSE SUM(a.GAAPPTIAnnual) / SUM(a.EquityAmt) END,

											ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal), 
											ActualCPR = CASE WHEN SUM(a.BeginPrinBal) = 0 THEN NULL ELSE 1 - (POWER(SUM(a.EndPrinBal) / SUM(a.BeginPrinBal), 12)) END

    FROM #PerformSummaryAll a
	where a.LoanCt > 0
    GROUP BY a.Client, a.PropState
),

---------------------------------------------------------------------------------
-- (3) RECAP CALCULATIONS
---------------------------------------------------------------------------------
Recap AS (
    SELECT
        a.Client,
        a.PropState,
											LoanCt = SUM(a.LoanCt), 
											UniqueLoanCt = SUM(a.UniqueLoanCt), 

											CohortRecap = SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal), 
											ActualRecap = SUM(a.RecapBal) / SUM(a.PIFBal)
    FROM #PerformSummaryAll a
    WHERE a.PIFBal > 0
	AND a.PIFReason IN ('Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance') 
    GROUP BY a.Client, a.PropState
)

---------------------------------------------------------------------------------
-- (4) UPDATE #ClientState WITH CPR, RECAP
---------------------------------------------------------------------------------
UPDATE a
SET 
						a.CPRLoanct = ISNULL(b.UniqueLoanCt,0),
						a.ModelCPR = b.ModelCPR,
						a.ActualCPR = b.ActualCPR,
						a.CPRScore =  CASE when ISNULL(b.ModelCPR, 0) = 0 OR b.LoanCt < 10 THEN 1 ELSE b.ActualCPR / b.ModelCPR end,
		
						a.RecapLoanct = ISNULL(c.UniqueLoanCt,0),
						a.CohortRecap = c.CohortRecap,
						a.ActualRecap = c.ActualRecap,
						a.RecapScore = case WHEN ISNULL(c.CohortRecap, 0) = 0 or ISNULL(c.LoanCt, 0) < 5 THEN 1 else c.ActualRecap / c.CohortRecap end

FROM #ClientState_VolumeSummaryAll a
LEFT JOIN Prepay b
       ON a.Client = b.Client
      AND a.PropState  = b.PropState
LEFT JOIN Recap c
       ON a.Client = c.Client
      AND a.PropState  = c.PropState;

---------------------------------------------------------------------------------
-- (5) UPDATE #ClientState WITH TOTALSCORE
---------------------------------------------------------------------------------
					UPDATE a
					SET a.TotalScore = CASE WHEN a.CPRScore = 0 THEN a.RecapScore * .5  WHEN a.RecapScore = 0 THEN (1 / a.CPRScore) * .5 ELSE (1 / a.CPRScore) * a.RecapScore end
					FROM #ClientState_VolumeSummaryAll a;

---------------------------------------------------------------------------------
-- (6) ATTACH SELLER FIELDS FROM #ClientSellerIDAE (Client metric)
---------------------------------------------------------------------------------

SELECT 
    cs.*,
	seller.[Account Manager_East], 
	seller.[Account Manager_West], 
	seller.[Account Manager_Central], 
	seller.[Account Manager_NULL], 
	seller.[ClientID_Delegated - Best Efforts], 
	seller.[ClientID_Delegated - Mandatory], 
	seller.[ClientID_Direct], 
	seller.[ClientID_FHLMC CRX], 
	seller.[ClientID_FNMA SMP], 
	seller.[ClientID_GNMA PIIT], 
	seller.[ClientID_Non-delegated]

FROM #ClientState_VolumeSummaryAll cs
LEFT JOIN #ClientSellerIDAE seller
       ON UPPER(LTRIM(RTRIM(seller.Client))) 
		   = UPPER(LTRIM(RTRIM(cs.Client)));

---------------------------------------------------------------------------------
-- END OF #ClientState
---------------------------------------------------------------------------------

--	select * from #ClientState_VolumeSummaryAll;

---------------------------------------------------------------------------------
---------------------------------------------------------------------------------

---------------------------------------------------------------------------------
---------------------------------------------------------------------------------
-- Create a table Client x State x Product from #VolumeSummaryAll, using same architecture as AllMetrics
---------------------------------------------------------------------------------
DROP TABLE IF EXISTS #ClientStateProduct_VolumeSummaryAll;
---------------------------------------------------------------------------------
-- (1) INITIALIZE #ClientState: GROUP BY Client, State, Product, BeginFundDt, EndFundDt
---------------------------------------------------------------------------------
SELECT  
      a.Client,
      a.PropState,
	  a.ProductCategory,
      a.BeginFundDt,
      a.EndFundDt,

      -- Base aggregated origination + performance metrics
								LoanCt = SUM(a.LoanCt), 
								FundedLoanAmt = SUM(a.FundedLoanAmt),
								VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),

								GovtPct = CASE WHEN SUM(a.FundedLoanAmt) = 0 THEN 0 ELSE SUM(a.GovtPct * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) end,

								WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.WAFICO ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							

								WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.WALTV ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

								WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.WADTI ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,


								WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.WAC ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,
			
								CPRLoanct = CONVERT(Int, NULL),
								ModelCPR = CONVERT(FLOAT, NULL),
								ActualCPR = CONVERT(FLOAT, NULL),
								CPRScore = CONVERT(FLOAT, 1),

								RecapLoanct = CONVERT(Int, NULL),
								CohortRecap = CONVERT(FLOAT, NULL),
								ActualRecap = CONVERT(FLOAT, NULL),
								RecapScore = CONVERT(FLOAT, 1),

								TotalScore = CONVERT(FLOAT, 1)

INTO #ClientStateProduct_VolumeSummaryAll
FROM #VolumeSummaryAll a
GROUP BY 
      a.Client,
      a.PropState,
	  a.ProductCategory,
      a.BeginFundDt,
      a.EndFundDt;

---------------------------------------------------------------------------------
-- (2) PREPAY CALCULATIONS (CPR)
---------------------------------------------------------------------------------
DROP TABLE IF EXISTS #Prepay;
    SELECT
         a.Client,
         a.PropState,
		 a.ProductCategory, 
											LoanCt = SUM(a.LoanCt), 
											UniqueLoanCt = SUM(a.UniqueLoanCt), 

											AdjROE = CASE WHEN SUM(a.EquityAmt) = 0 then NULL ELSE SUM(a.AdjPTIAnnual) / SUM(a.EquityAmt) end, 
											GAAPROE = CASE WHEN SUM(a.EquityAmt) = 0 then NULL ELSE SUM(a.GAAPPTIAnnual) / SUM(a.EquityAmt) END,

											ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal), 
											ActualCPR = CASE WHEN SUM(a.BeginPrinBal) = 0 THEN NULL ELSE 1 - (POWER(SUM(a.EndPrinBal) / SUM(a.BeginPrinBal), 12)) END

	INTO #Prepay
    FROM #PerformSummaryAll a
	where a.LoanCt > 0
    GROUP BY a.Client, a.PropState, a.ProductCategory;

---------------------------------------------------------------------------------
-- (3) RECAP CALCULATIONS
---------------------------------------------------------------------------------
DROP TABLE IF EXISTS #Recap;
    SELECT
        a.Client,
        a.PropState,
		a.ProductCategory, 
											LoanCt = SUM(a.LoanCt), 
											UniqueLoanCt = SUM(a.UniqueLoanCt), 

											CohortRecap = SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal), 
											ActualRecap = SUM(a.RecapBal) / SUM(a.PIFBal)
	INTO #Recap
    FROM #PerformSummaryAll a
    WHERE a.PIFBal > 0
	AND a.PIFReason IN ('Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance') 
    GROUP BY a.Client, a.PropState, a.ProductCategory;

---------------------------------------------------------------------------------
-- (4) UPDATE #ClientState WITH CPR, RECAP
---------------------------------------------------------------------------------
UPDATE a
SET 
						a.CPRLoanct = ISNULL(b.UniqueLoanCt,0),
						a.ModelCPR = b.ModelCPR,
						a.ActualCPR = b.ActualCPR,
						a.CPRScore =  CASE when ISNULL(b.ModelCPR, 0) = 0 OR b.LoanCt < 10 THEN 1 ELSE b.ActualCPR / b.ModelCPR end,
		
						a.RecapLoanct = ISNULL(c.UniqueLoanCt,0),
						a.CohortRecap = c.CohortRecap,
						a.ActualRecap = c.ActualRecap,
						a.RecapScore = case WHEN ISNULL(c.CohortRecap, 0) = 0 or ISNULL(c.LoanCt, 0) < 5 THEN 1 else c.ActualRecap / c.CohortRecap end

FROM #ClientStateProduct_VolumeSummaryAll a
LEFT JOIN #Prepay b
       ON a.Client = b.Client
      AND a.PropState  = b.PropState
	  AND a.ProductCategory = b.ProductCategory
LEFT JOIN #Recap c
       ON a.Client = c.Client
      AND a.PropState  = c.PropState
	  AND a.ProductCategory = c.ProductCategory
	  ;

---------------------------------------------------------------------------------
-- (5) UPDATE #ClientState WITH TOTALSCORE
---------------------------------------------------------------------------------
					UPDATE a
					SET a.TotalScore = CASE WHEN a.CPRScore = 0 THEN a.RecapScore * .5  WHEN a.RecapScore = 0 THEN (1 / a.CPRScore) * .5 ELSE (1 / a.CPRScore) * a.RecapScore end
					FROM #ClientStateProduct_VolumeSummaryAll a;

---------------------------------------------------------------------------------
-- (6) ATTACH SELLER FIELDS FROM #ClientSellerIDAE (Client metric)
---------------------------------------------------------------------------------

SELECT 
    cs.*,
	seller.[Account Manager_East], 
	seller.[Account Manager_West], 
	seller.[Account Manager_Central], 
	seller.[Account Manager_NULL], 
	seller.[ClientID_Delegated - Best Efforts], 
	seller.[ClientID_Delegated - Mandatory], 
	seller.[ClientID_Direct], 
	seller.[ClientID_FHLMC CRX], 
	seller.[ClientID_FNMA SMP], 
	seller.[ClientID_GNMA PIIT], 
	seller.[ClientID_Non-delegated]

FROM #ClientStateProduct_VolumeSummaryAll cs
LEFT JOIN #ClientSellerIDAE seller
       ON UPPER(LTRIM(RTRIM(seller.Client))) 
		   = UPPER(LTRIM(RTRIM(cs.Client)));

---------------------------------------------------------------------------------
-- END OF #ClientState
---------------------------------------------------------------------------------

--	select * from #ClientState_VolumeSummaryAll;

---------------------------------------------------------------------------------
---------------------------------------------------------------------------------





















































------------------------------------------------------------------------------------------------------

		--Weighted Client 

		DROP TABLE IF EXISTS #ClientVolume

			SELECT	BeginFundDt = @BeginRefFundDt,
					EndFundDt = @EndRefFundDt,
					a.Client, 
					LoanCt = COUNT(*),
					FundedLoanAmt = SUM(a.FundedLoanAmt),
					VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),
					GovtPct = SUM(case when a.ProductCategory IN ('FHA', 'VA', 'USDA') then 1 else 0 END) / SUM(1.0),

					WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.FICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
									SUM(CASE WHEN ISNULL(a.FICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.FICO ELSE 0 END) 
										/ SUM(CASE WHEN ISNULL(a.FICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,

					WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.LTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
									SUM(CASE WHEN ISNULL(a.LTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.LTV ELSE 0 END) 
										/ SUM(CASE WHEN ISNULL(a.LTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

					WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.DTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
									SUM(CASE WHEN ISNULL(a.DTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.DTI ELSE 0 END) 
										/ SUM(CASE WHEN ISNULL(a.DTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,

					WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.Rate, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
									SUM(CASE WHEN ISNULL(a.Rate, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.Rate ELSE 0 END) 
										/ SUM(CASE WHEN ISNULL(a.Rate, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END

			INTO #ClientVolume
			FROM adm.ODS.vwOrigLoan a
			WHERE a.Channel IN ('Correspondent', 'Co-Issue')
			and a.Client is not null
			AND a.FundDt BETWEEN @BeginRefFundDt and @EndRefFundDt
			and a.FundedLoanAmt > 0
			AND a.MortgageType <> 'Reverse'
			  -- Drop any LoanNum that EVER XFERed (transferred) 
			  AND NOT EXISTS (
					SELECT 1
					FROM Analytics_Rpt.md.B2BPricingHist as h
					WHERE h.LoanNum = a.LoanNum
					  AND (h.ResolType = 'XFER' OR h.Channel = 'Bulk')
			  )

			GROUP BY a.Client


--			select * from #ClientVolume;

			DROP TABLE IF EXISTS #AllMetricsWeighted

			CREATE TABLE #AllMetricsWeighted (ProductCategory VARCHAR(50), MetricName varchar(50), MetricValue VARCHAR(200), LoanCt INT,	FundedLoanAmt MONEY, VolumeRank INT, GovtPct FLOAT, WAFICO FLOAT, WALTV FLOAT,
								WADTI FLOAT, WAC FLOAT, CPRLoanct Int, ModelCPR FLOAT, ActualCPR FLOAT,CPRScore FLOAT, RecapLoanct Int, CohortRecap FLOAT, ActualRecap FLOAT, RecapScore FLOAT, TotalScore FLOAT)


			set @Metric = 'Total'
	
			WHILE @Metric <> 'Exit'

				BEGIN

						DROP TABLE IF EXISTS #MetricSummaryWeighted

						SELECT	
								a.ProductCategory,
								MetricName = @Metric,
								MetricValue = CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB END,
								LoanCt = SUM(a.LoanCt), 
								FundedLoanAmt = SUM(a.FundedLoanAmt),
								VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),

								GovtPct = CASE WHEN SUM(a.FundedLoanAmt) = 0 THEN 0 ELSE SUM(a.GovtPct * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) end,

								WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.WAFICO ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							

								WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.WALTV ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

								WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.WADTI ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,


								WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.WAC ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,
			
								CPRLoanct = SUM(a.UniqueLoanCt), 
								ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal), 
								ActualCPR = CASE WHEN SUM(a.BeginPrinBal) = 0 THEN NULL ELSE 1 - (POWER(SUM(a.EndPrinBal) / SUM(a.BeginPrinBal), 12)) END,
								CPRScore =  CONVERT(FLOAT, NULL),


								RecapLoanct = SUM(CASE WHEN a.PIFBal > 0 THEN a.UniqueLoanCt ELSE 0 end), 
								CohortRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal) end, 
								ActualRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.RecapBal) / SUM(a.PIFBal) END,
								RecapScore = CONVERT(FLOAT, 1),

								TotalScore = CONVERT(FLOAT, NULL)

						INTO #MetricSummaryWeighted
						FROM #PerformSummaryAll a
						GROUP BY a.ProductCategory, CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB END

					UPDATE a
					SET 
						a.CPRScore =  CASE when ISNULL(a.ModelCPR, 0) = 0 OR a.CPRLoanct < 10 THEN 1 ELSE a.ActualCPR / a.ModelCPR end,
						a.RecapScore = case WHEN ISNULL(a.CohortRecap, 0) = 0 or ISNULL(a.RecapLoanct, 0) < 5 THEN 1 else a.ActualRecap / a.CohortRecap end
					FROM #MetricSummaryWeighted a


					UPDATE a
					SET a.TotalScore = CASE WHEN a.CPRScore = 0 THEN a.RecapScore * .5  WHEN a.RecapScore = 0 THEN (1 / a.CPRScore) * .5 ELSE (1 / a.CPRScore) * a.RecapScore end
					FROM #MetricSummaryWeighted a

--					select * from #MetricSummaryWeighted;

					INSERT INTO #AllMetricsWeighted
					SELECT a.*
					FROM #MetricSummaryWeighted a
					ORDER BY 2

					SET @Metric = CASE @Metric WHEN 'Total' THEN 'State' WHEN 'State' THEN 'FICO' WHEN 'FICO' THEN 'LTV' WHEN 'LTV' THEN 'DTI' WHEN 'DTI' THEN 'Interest Rate' WHEN 'Interest Rate' THEN 'Product' WHEN 'Product' THEN 'Avg Loan Bal' WHEN 'Avg Loan Bal' THEN 'Exit' end

				END

--				select * from #AllMetricsWeighted;


			DROP TABLE IF EXISTS #ClientMetrics

			CREATE TABLE #ClientMetrics (Client varchar(200), ProductCategory VARCHAR(50), MetricName varchar(50), MetricValue VARCHAR(200), LoanCt INT,	FundedLoanAmt MONEY, VolumeRank INT, GovtPct FLOAT, WAFICO FLOAT, WALTV FLOAT,
								WADTI FLOAT, WAC FLOAT, CPRLoanct Int, ModelCPR FLOAT, ActualCPR FLOAT,CPRScore FLOAT, RecapLoanct Int, CohortRecap FLOAT, ActualRecap FLOAT, RecapScore FLOAT, TotalScore FLOAT)


			set @Metric = 'Total'
	
			WHILE @Metric <> 'Exit'

				BEGIN

						DROP TABLE IF EXISTS #ClientSummary

						SELECT	
								a.Client, 
								a.ProductCategory,
								MetricName = @Metric,
								MetricValue = CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB END,
								LoanCt = SUM(a.LoanCt), 
								FundedLoanAmt = SUM(a.FundedLoanAmt),
								VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),

								GovtPct = CASE WHEN SUM(a.FundedLoanAmt) = 0 THEN 0 ELSE SUM(a.GovtPct * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) end,

								WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.WAFICO ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							

								WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.WALTV ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

								WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.WADTI ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,


								WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.WAC ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,
			
								CPRLoanct = SUM(a.UniqueLoanCt), 
								ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal), 
								ActualCPR = CASE WHEN SUM(a.BeginPrinBal) = 0 THEN NULL ELSE 1 - (POWER(SUM(a.EndPrinBal) / SUM(a.BeginPrinBal), 12)) END,
								CPRScore =  CONVERT(FLOAT, null),


								RecapLoanct = SUM(CASE WHEN a.PIFBal > 0 THEN a.UniqueLoanCt ELSE 0 end), 
								CohortRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal) end, 
								ActualRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.RecapBal) / SUM(a.PIFBal) END,
								RecapScore = CONVERT(FLOAT, null),

								TotalScore = CONVERT(FLOAT, null)

						INTO #ClientSummary	
						FROM #PerformSummaryAll a
						GROUP BY a.Client, a.ProductCategory, CASE @Metric WHEN 'Total' THEN 'Total' WHEN 'Client' THEN a.Client WHEN 'State' THEN a.PropState WHEN 'FICO' THEN a.FICOBand when 'LTV' THEN a.LTVBand WHEN 'DTI' THEN a.DTIBand WHEN 'Interest Rate' THEN a.IntRateBand WHEN 'Product' THEN a.ProductCategory WHEN 'Avg Loan Bal' THEN a.ALB END


					UPDATE a
					SET 
						a.CPRScore =  CASE when ISNULL(a.ModelCPR, 0) = 0 OR a.CPRLoanct < 10 THEN 1 ELSE a.ActualCPR / a.ModelCPR end,
						a.RecapScore = case WHEN ISNULL(a.CohortRecap, 0) = 0 or ISNULL(a.RecapLoanct, 0) < 5 THEN 1 else a.ActualRecap / a.CohortRecap end
					FROM #ClientSummary a


					UPDATE a
					SET a.TotalScore = CASE WHEN a.CPRScore = 0 THEN a.RecapScore * .5  WHEN a.RecapScore = 0 THEN (1 / a.CPRScore) * .5 ELSE (1 / a.CPRScore) * a.RecapScore end
					FROM #ClientSummary a

--					select * from #ClientSummary;


					INSERT INTO #ClientMetrics
					SELECT a.*
					FROM #ClientSummary a
					ORDER BY 2

--					select top 1000 *  from #ClientMetrics;

					SET @Metric = CASE @Metric WHEN 'Total' THEN 'State' WHEN 'State' THEN 'FICO' WHEN 'FICO' THEN 'LTV' WHEN 'LTV' THEN 'DTI' WHEN 'DTI' THEN 'Interest Rate' WHEN 'Interest Rate' THEN 'Product' WHEN 'Product' THEN 'Avg Loan Bal' WHEN 'Avg Loan Bal' THEN 'Exit' end

				END

--		select top 1000 *  from #ClientMetrics;

---------------------------------------------------------------------------------
---------------------------------------------------------------------------------
-- Create a table Client x State from #PerformSummaryAll
DROP TABLE IF EXISTS #ClientState_PerformSummaryAll;

						SELECT	
								a.Client, 
								State = a.PropState,
								LoanCt = SUM(a.LoanCt), 
								FundedLoanAmt = SUM(a.FundedLoanAmt),
								VolumeRank = ROW_NUMBER() OVER (ORDER BY  SUM(a.FundedLoanAmt) DESC),

								GovtPct = CASE WHEN SUM(a.FundedLoanAmt) = 0 THEN 0 ELSE SUM(a.GovtPct * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) end,

								WAFICO = CASE WHEN SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt * a.WAFICO ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAFICO, 0) BETWEEN 580 AND 850 THEN a.FundedLoanAmt ELSE 0 END) END,
							

								WALTV = CASE WHEN SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt * a.WALTV ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WALTV, 0) BETWEEN .10 AND .99 THEN a.FundedLoanAmt ELSE 0 END) END,

								WADTI = CASE WHEN SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt * a.WADTI ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WADTI, 0) BETWEEN .05 AND .65 THEN a.FundedLoanAmt ELSE 0 END) END,


								WAC = CASE WHEN SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) = 0 THEN NULL ELSE 
												SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt * a.WAC ELSE 0 END) 
													/ SUM(CASE WHEN ISNULL(a.WAC, 0) BETWEEN .02 AND .10 THEN a.FundedLoanAmt ELSE 0 END) END,
			
								CPRLoanct = SUM(a.UniqueLoanCt), 
								ModelCPR = SUM(a.ModelCPR * a.BeginPrinBal) / SUM(a.BeginPrinBal), 
								ActualCPR = CASE WHEN SUM(a.BeginPrinBal) = 0 THEN NULL ELSE 1 - (POWER(SUM(a.EndPrinBal) / SUM(a.BeginPrinBal), 12)) END,
								CPRScore =  CONVERT(FLOAT, null),


								RecapLoanct = SUM(CASE WHEN a.PIFBal > 0 THEN a.UniqueLoanCt ELSE 0 end), 
								CohortRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.CohortRecap * a.PIFBal) / SUM(a.PIFBal) end, 
								ActualRecap = CASE WHEN SUM(a.PIFBal) = 0 THEN NULL ELSE SUM(a.RecapBal) / SUM(a.PIFBal) END,
								RecapScore = CONVERT(FLOAT, null),

								TotalScore = CONVERT(FLOAT, null)

						INTO #ClientState_PerformSummaryAll	
						FROM #PerformSummaryAll a
						GROUP BY a.Client, a.PropState

					UPDATE a
					SET 
						a.CPRScore =  CASE when ISNULL(a.ModelCPR, 0) = 0 OR a.CPRLoanct < 10 THEN 1 ELSE a.ActualCPR / a.ModelCPR end,
						a.RecapScore = case WHEN ISNULL(a.CohortRecap, 0) = 0 or ISNULL(a.RecapLoanct, 0) < 5 THEN 1 else a.ActualRecap / a.CohortRecap end
					FROM #ClientState_PerformSummaryAll a


					UPDATE a
					SET a.TotalScore = CASE WHEN a.CPRScore = 0 THEN a.RecapScore * .5  WHEN a.RecapScore = 0 THEN (1 / a.CPRScore) * .5 ELSE (1 / a.CPRScore) * a.RecapScore end
					FROM #ClientState_PerformSummaryAll a

--					select * from #ClientState_PerformSummaryAll;

---------------------------------------------------------------------------------
---------------------------------------------------------------------------------

		DROP TABLE IF EXISTS #ClientAnalysis


			;WITH ClientScore as (Select a.Client, 
										OverallScore = SUM(a.FundedLoanAmt * (a.TotalScore / b.TotalScore)) / SUM(a.FundedLoanAmt)
								  FROM #ClientMetrics a

										inner JOIN #AllMetricsWeighted b ON a.ProductCategory = b.ProductCategory and a.MetricName = b.MetricName AND a.MetricValue = b.MetricValue

								  WHERE a.TotalScore IS NOT NULL
								  AND isnull(b.TotalScore, 0) <> 0
								  GROUP BY a.Client)

						SELECT a.*, b.OverallScore, AvgB2BScore = CONVERT(FLOAT, NULL), ClientGroup = CONVERT(VARCHAR(50), null)
						INTO #ClientAnalysis
						from #ClientVolume a

							inner JOIN ClientScore b on a.Client = b.Client

						ORDER BY a.VolumeRank

--						select * from #ClientAnalysis;

			;WITH ClientAvg AS (SELECT AvgB2BScore = SUM(a.OverallScore * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) FROM #ClientAnalysis a )

			update a
			SET a.AvgB2BScore = b.AvgB2BScore, 
				a.ClientGroup = CASE WHEN a.OverallScore >= 1.25 * b.AvgB2BScore THEN 1 WHEN a.OverallScore <= 0.75 * b.AvgB2BScore THEN 3 else 2 end
			from #ClientAnalysis a

				CROSS JOIN ClientAvg b


--			SELECT a.*
--			FROM #ClientAnalysis a
--			WHERE a.VolumeRank <= 500
--			ORDER BY a.VolumeRank


	DROP TABLE IF EXISTS #ClientDetails

			SELECT	a.Client, 
					a.ProductCategory,
					a.MetricName,
					a.MetricValue,
					OverallScore = SUM(a.FundedLoanAmt * (a.TotalScore / b.TotalScore)) / SUM(a.FundedLoanAmt),
					AvgB2BScore = CONVERT(FLOAT, NULL),
					ClientGroup = CONVERT(VARCHAR(50), null),
					VolumeRank = CONVERT(INT, NULL)
			INTO #ClientDetails
			FROM #ClientMetrics a

				inner JOIN #AllMetricsWeighted b ON a.ProductCategory = b.ProductCategory and a.MetricName = b.MetricName AND a.MetricValue = b.MetricValue

			WHERE a.TotalScore IS NOT NULL
			AND isnull(b.TotalScore, 0) <> 0
			AND a.Client IS not null
			GROUP BY a.Client, a.ProductCategory, a.MetricName, a.MetricValue


			UPDATE a
			SET a.VolumeRank = b.VolumeRank
			from #ClientDetails a

				INNER JOIN #ClientVolume b ON a.Client = b.Client


			;WITH ClientAvg AS (SELECT AvgB2BScore = SUM(a.OverallScore * a.FundedLoanAmt) / SUM(a.FundedLoanAmt) FROM #ClientAnalysis a )

			update a
			SET a.AvgB2BScore = b.AvgB2BScore, 
				a.ClientGroup = CASE WHEN a.OverallScore >= 1.25 * b.AvgB2BScore THEN 1 WHEN a.OverallScore <= 0.75 * b.AvgB2BScore THEN 3 else 2 end
			from #ClientDetails a

				CROSS JOIN ClientAvg b

			
--			SELECT a.*
--			FROM #ClientDetails a
--			WHERE a.VolumeRank IS NOT NULL
--			ORDER BY a.VolumeRank, 2, 3, 4

