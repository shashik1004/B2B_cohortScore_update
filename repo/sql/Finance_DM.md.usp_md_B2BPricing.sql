USE [Analytics_Rpt]
GO
/****** Object:  StoredProcedure [md].[usp_md_B2BPricing]    Script Date: 4/30/2026 9:24:11 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROC  [md].[usp_md_B2BPricing] (@pi_AsofDt DATE, @pi_Loan VARCHAR(10) = null)

AS

/******************************************************************************************************************************************************************

	Origination B2B Pricing
	Created by: Jack Cavanagh 
	Created: 01/11/2026
	Sample: 

		EXEC md.usp_md_B2BPricing  @pi_AsofDt =  '2026-06-30'

		UpdateDt		UpdatedBy			UpdateDesc
		--------		------------		-------------------------------------------------------------------------
		03/05/2026		Cavanagh			Disconnected Loan Level Servicing
		04/29/2026		Cavanagh			Expanded to include all MSR
		05/03/2026		Adamakis			Remove duplicates when joining with fnc_Stg.stg.RL_Dashboard_Details and change the AsOfDt to reflect the month that this code is executed
******************************************************************************************************************************************************************/

/***************************************************************************************************
DATA TIMING & REPORTING CONVENTION (IMPORTANT)

This procedure intentionally separates the *source servicing month* from the *reporting AsOf month*
to avoid dependency on mid-month servicing P&L closes.

--------------------------------------------------------------------------------
KEY DATES
--------------------------------------------------------------------------------
@pi_AsofDt
    = Reporting / History month end
    = The AsOfDt stored in Analytics_Rpt.md.B2BPricingHist
    = Example (May run): '2026-04-30'

@SourceAsofDt = EOMONTH(@pi_AsofDt, -1)
    = Prior month-end servicing snapshot
    = Example (May run): '2026-03-31'

--------------------------------------------------------------------------------
RATIONALE
--------------------------------------------------------------------------------
• MSR Values (md.MSRValuesLoanLevel) are finalized one month in arrears.
• Servicing Loan-Level P&L (md.ServicingLoanLevelPnL) is not final until books close mid-month.
• To avoid delaying reporting, this procedure pulls PRIOR-month source data while labeling results
  as the CURRENT reporting month.
• Payoff reasons and recapture attribution may change until mid-month; therefore results should
  be treated as PRELIMINARY.

--------------------------------------------------------------------------------
TABLE DATE USAGE (Example: Run in May with @pi_AsofDt = '2026-04-30')
--------------------------------------------------------------------------------

1. ods.vwServLoanHist a  (PRIMARY DRIVER – Servicing Snapshot)
   • a.AsofDt = @SourceAsofDt
   • Example: 2026-03-31
   • Provides balances, delinquency, servicing status, and active loan population

2. md.MSRValuesLoanLevel b
   • b.AsofDt = @SourceAsofDt
   • Example: 2026-03-31
   • MSR valuation (known one month in arrears)

3. md.MSRValuesLoanLevel (prior-period comparison)
   • b.AsofDt = EOMONTH(@SourceAsofDt, -1)
   • Example: 2026-02-28
   • Used to calculate month-over-month MSR fair value change

4. md.ServicingLoanLevelPnL e
   • e.AsofDt = @SourceAsofDt
   • Example: 2026-03-31
   • Servicing revenue / expense used on a PRELIMINARY basis

5. ods.vwServLoanHist d  (NEXT-MONTH EVENTS)
   • d.AsofDt = EOMONTH(@SourceAsofDt, 1)
   • Example: 2026-04-30
   • Used to detect payoff, resolution type (PIF, XFER), and end balances

6. Finance_DM.Ods.SasModelScores g
   • g.EndDt = '12/31/9999'
   • Uses current active record
   • Provides payoff behavior, refinance intent, and recapture attributes

7. fnc_Stg.stg.RL_Dashboard_Details (via RecapAgg CTE)
   • Uses latest available funded refinance record (fundssentdate IS NOT NULL)
   • Not AsOfDt-driven
   • Provides borrower requested loan amount for recapture calculations

8. adm.ODS.vwOrigLoan c
   • Not AsOfDt-driven
   • Provides origination attributes (Channel, SubChannel, Client, ProductCategory)

--------------------------------------------------------------------------------
OUTPUT
--------------------------------------------------------------------------------
• Analytics_Rpt.md.B2BPricingHist.AsofDt = @pi_AsofDt
  (Example: 2026-04-30)

• Data reflects:
    - Prior-month servicing snapshot
    - Current-month payoff / resolution activity
    - Preliminary servicing P&L

--------------------------------------------------------------------------------
IMPORTANT IMPLEMENTATION NOTE
--------------------------------------------------------------------------------
Once #LoanLevelHist is created:
• AsofDt represents the REPORTING month (not the source month)
• Do NOT use a.AsofDt for source-relative month calculations downstream
• Use @SourceAsofDt explicitly when performing month-over-month source logic

***************************************************************************************************/

DECLARE @AsofDt DATE, @SourceAsofDt DATE, @Loan VARCHAR(10), @GSEAdvRate FLOAT, @GNMAAdvRate FLOAT, @CDAdjPTIBps FLOAT, @GSEMSRBps FLOAT, @GNMAMSRBps float


-- @AsofDt = reporting / history month (the month that this SP is executed)
		SET @AsofDt = @pi_AsofDt

-- @SourceAsofDt = servicing source month (previous month)
		SET @SourceAsofDt = EOMONTH(@pi_AsofDt, -1)

		SET @Loan = @pi_Loan

		SET @GSEAdvRate = 0.70

		SET @GNMAAdvRate = 0.65

		set @CDAdjPTIBps = 185


		--Default if no MSR value yet
		set @GSEMSRBps = 150

		set @GNMAMSRBps = 200


	
	DROP TABLE IF EXISTS #LoanLevelHist
	
	;WITH RecapAgg AS (
		SELECT
			priorloanrecordationoriginalloannumber,
			SUM(borrowerrequestedloanamount) AS borrowerrequestedloanamount
		FROM fnc_Stg.stg.RL_Dashboard_Details WITH (NOLOCK)
		WHERE fundssentdate IS NOT NULL
		GROUP BY priorloanrecordationoriginalloannumber
	)
	    
		SELECT	
				@AsofDt AS AsofDt,
				a.LoanNum,
				a.ActiveFlag,
				InvestorGroup = CASE WHEN a.InvestorGroup = 'PLS' AND a.LoanType IN ('CONV-PMI', 'CONV') THEN 'GSE' WHEN a.InvestorGroup = 'PLS' THEN 'GNMA' ELSE a.InvestorGroup END,
				c.ProductCategory,
				a.LoanType,
				Channel = ISNULL(c.Channel, 'Bulk'),
				SubChannel = ISNULL(c.SubChannel, 'Bulk'),
				a.PropState,
				c.Client,
				a.ServicingType,
				a.ServicingSubType,
				c.FundDt,
				c.FundedLoanAmt,
				FirstMSRVal = CONVERT(MONEY, 0),
				a.DelqStatusMBA, 
				a.DelqDaysMBA, 
				a.PrinBal, 
				MSRValAmt = ISNULL(b.ValueTotal, 0),
				b.ModelCPR,
				d.ResolType,
				d.ResolDt,
				PIFReason = CASE WHEN EOMONTH(g.PAYMENT_IN_FULL_DATE) = d.AsofDt THEN 
									CASE When g.purpose IN ('Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance') then g.purpose else 'No New Loan' END
							ELSE NULL end,
				SolicitableFlag = CASE WHEN g.solicitability = 'Solicit' then 1 else 0 end,
				PeerRecapFlag = CASE WHEN EOMONTH(g.PAYMENT_IN_FULL_DATE) = d.AsofDt AND g.category = 'Refinance' AND g.solicitability = 'Solicit' and g.mrc_recap_drop_reason IS null then 1 else 0 end,
				BeginPrinBal = a.PrinBal,
				EndPrinBal = CASE WHEN d.ActiveFlag = 0 AND d.ResolType IN ('PIF') THEN 0 WHEN d.ActiveFlag = 0 THEN a.PrinBal WHEN a.Prinbal - d.PrinBal >= 10000 THEN d.PrinBal ELSE a.PrinBal END,
				
				CohortRecap = CONVERT(FLOAT, NULL),
				RecapBal =  CASE WHEN EOMONTH(g.PAYMENT_IN_FULL_DATE) = d.AsofDt then ISNULL(h.borrowerrequestedloanamount,0) ELSE NULL end,
				PIFBal = CASE WHEN EOMONTH(g.PAYMENT_IN_FULL_DATE) = d.AsofDt THEN g.FIRST_PRINCIPAL_BALANCE else NULL end,

				ServRevAmt = e.ServSubServFeesAmt + e.GFeeAmt + e.MasterServAmt + e.GoSAmt + e.FloatAmt + e.OthAncAmt, + e.OthRevAmt,
				ServMSRRunAmt = e.MSRRunoffAmt + e.ESSRunoffAmt,
				ServOpexAmt = e.CompBenAmt + e.ServOrigAmt + e.TechCommAmt + e.ProfFeesAmt + e.OccupAmt + e.OthOpAmt,
				ServOthIncExpAmt = e.IntIncAmt + e.IntExpAmt + e.ESSFeesAmt + e.OthIncExpAmt,
				ServCorpAllocAmt = e.CorpAllocAmt,
				MSRFVChgNet = CONVERT(MONEY, 0),
				ServAdjPTI = CONVERT(MONEY, 0),
				ServGAAPPTI = CONVERT(MONEY, 0),
				ServCashPTI = CONVERT(MONEY, 0),
				CDAdjPTI = case when EOMONTH(g.PAYMENT_IN_FULL_DATE) = d.AsofDt THEN ISNULL(h.borrowerrequestedloanamount,0) * (@CDAdjPTIBps / 10000.0) ELSE 0 END,
				CDCashPTI = CONVERT(MONEY, 0),

				EquityAmt = ISNULL(e.EquityAmt, 0)

		INTO #LoanLevelHist
		FROM adm.ods.vwServLoanHist a

			INNER JOIN  Finance_DM.md.MSRValuesLoanLevel b ON a.LoanNum = b.LoanNum AND a.AsofDt = b.Asofdt

			LEFT JOIN adm.ODS.vwOrigLoan c ON a.LoanNum = c.LoanNum

			Left JOIN  adm.ods.vwServLoanHist d ON a.LoanNum = d.LoanNum AND d.AsofDt = EOMONTH(a.Asofdt, 1)

			LEFT JOIN  Finance_DM.md.ServicingLoanLevelPnL e ON a.LoanNum = e.LoanNum AND a.AsofDt = e.Asofdt

			LEFT JOIN Finance_DM.Ods.SasModelScores g ON a.LoanNum = g.LOAN_NUMBER and g.EndDt = '12/31/9999'

			LEFT JOIN RecapAgg h
				ON a.LoanNum = h.priorloanrecordationoriginalloannumber

		WHERE a.AsofDt = @SourceAsofDt
--		AND a.ServicingType = 'MSR'



		UPDATE a
		SET a.FirstMSRVal = b.ValueTotal
        FROM #LoanLevelHist a

			INNER JOIN Finance_DM.md.MSRValuesLoanLevel b ON a.LoanNum = b.LoanNum AND b.Asofdt = (SELECT MIN(x.AsofDt) FROM Finance_DM.md.MSRValuesLoanLevel x WHERE b.LoanNum = x.LoanNum)


		UPDATE a
		SET a.FirstMSRVal = CASE WHEN a.InvestorGroup = 'GSE' then a.FundedLoanAmt * (@GSEMSRBps / 10000.0) else a.FundedLoanAmt * (@GNMAMSRBps / 10000.0) END
        FROM #LoanLevelHist a
		where a.FirstMSRVal = 0



		UPDATE a
		SET a.MSRFVChgNet = a.MSRValAmt - b.ValueTotal
        FROM #LoanLevelHist a

			INNER JOIN Finance_DM.md.MSRValuesLoanLevel b ON a.LoanNum = b.LoanNum AND b.Asofdt = EOMONTH(@SourceAsofDt, -1)

		where a.ServicingType = 'MSR'



		UPDATE a
		SET a.ServAdjPTI = a.ServRevAmt + a.ServMSRRunAmt + a.ServOpexAmt + a.ServOthIncExpAmt + ServCorpAllocAmt,
			a.ServGAAPPTI = a.ServRevAmt + a.ServMSRRunAmt + a.ServOpexAmt + a.ServOthIncExpAmt + ServCorpAllocAmt + a.MSRFVChgNet,
			a.ServCashPTI = a.ServRevAmt + a.ServMSRRunAmt * CASE WHEN a.InvestorGroup = 'GSE' THEN @GSEAdvRate ELSE @GNMAAdvRate end + a.ServOpexAmt + a.ServOthIncExpAmt + ServCorpAllocAmt
		FROM #LoanLevelHist a


		UPDATE a
		SET a.CDAdjPTI = a.CDAdjPTI - Case when a.ServicingType = 'MSR' THEN 0 ELSE a.FirstMSRVal END,
			a.CDCashPTI = a.CDAdjPTI -	a.FirstMSRVal
			
        FROM #LoanLevelHist a
		where a.RecapBal > 0


		;With CohortRecap AS
				(Select 
				PIFMonth = EOMONTH(a.PAYMENT_IN_FULL_DATE),
				b.ProductCategory,
				a.purpose,
				PeerRecapFlag = CASE WHEN a.category = 'Refinance' AND a.solicitability = 'Solicit' and a.mrc_recap_drop_reason IS null then 1 else 0 end,

				RecapRate = SUM(ISNULL(c.borrowerrequestedloanamount, 0)) / SUM(CONVERT(FLOAT, a.first_principal_balance))

				from Finance_DM.Ods.SasModelScores   a with (nolock)

					INNER JOIN adm.ODS.vwOrigLoan b ON a.LOAN_NUMBER = b.LoanNum					
	
					LEFT join fnc_Stg.stg.RL_Dashboard_Details c with (nolock)
						ON a.loan_number = c.priorloanrecordationoriginalloannumber
						AND c.fundssentdate is not null
		
				WHERE a.purpose IN ('Purchase', 'Rate/Term Refinance', 'Cash-Out Refinance') 
				and a.EndDt = '12/31/9999'
				and a.investor_type IN ('Freddie Mac (FHLMC)', 'Fannie Mae (FNMA)', 'Ginnie Mae (GNMA)')
				GROUP BY EOMONTH(a.PAYMENT_IN_FULL_DATE), b.ProductCategory, a.purpose, CASE WHEN a.category = 'Refinance' AND a.solicitability = 'Solicit' and a.mrc_recap_drop_reason IS null then 1 else 0 END
				having SUM(CONVERT(FLOAT, a.first_principal_balance)) > 0)


		update a
		set a.CohortRecap = b.RecapRate
		FROM #LoanLevelHist a

			inner join CohortRecap b ON EOMONTH(@SourceAsofDt, 1) = b.PIFMonth AND a.ProductCategory = b.ProductCategory AND a.PIFReason = b.purpose AND a.PeerRecapFlag = b.PeerRecapFlag


		--DROP TABLE if EXISTS Analytics_Rpt.md.B2BPricingHist
		DELETE a FROM Analytics_Rpt.md.B2BPricingHist a WHERE a.AsofDt = @AsofDt

		INSERT INTO Analytics_Rpt.md.B2BPricingHist
		SELECT a.*
		--INTO Analytics_Rpt.md.B2BPricingHist
		FROM #LoanLevelHist a


