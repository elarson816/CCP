** Bivariate Analysis (Females Only)
cd $directory

local multivariate $multivariate

** Significant variables from bivriate analysis
/*  Knows LARCS
	Knows Short Acting
	Knows more than the median number of methods
	Used FP services in the last 12 months
	PII
	Exposed to FP messaging
	Has at least 1 cellphone in the house
	Age
	Religion
	Marital Status
*/

********************************************************
*** CLEAN FAMILY PLANNING DATA ***
********************************************************
* Set Put Excel
putexcel set "$putexcel_set", modify sheet("`multivariate'")

use "1. Data/baseline_femaledata.dta", clear

* Correlation of similar variables
corr heard_combined_shortacting heard_combined_larc heard_combined_gt_median // heard_combined_gt_median has narrowest CI
corr fp_messaging media_cellphone_often
corr gem_pv_index gem_rh_index gem_sr_index
corr visited_provider_12mo MII used_fp_quality
corr couple_communication couple_communication_index

* Run the regression
logistic mcp ur ib24.county_id ///
			 i.age_cat inunion education i.religion given_birth ///
			 heard_combined_larc heard_combined_gt_median ///
			 i.fp_self_efficacy i.fp_perceived_norms i.couple_communication ///
			 visited_provider_12mo MII ///
			 fp_messaging media_cellphone_often ///
			 i.gem_pv_index i.gem_rh_index i.gem_sr_index ///
			 decision_numchildren_yn decision_contraception_yn decision_selfill_yn i.couple_communication_index ///
			 media_radio_often media_tv_often
			
	** PutExcel
	putexcel (A1)=etable
			


