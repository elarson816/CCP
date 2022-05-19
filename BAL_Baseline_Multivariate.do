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

* Run the regression
logistic cp ur ib24.county_id ///
			age_cat inunion education i.religion ///
			heard_combined_gt_median ///
			visited_provider_12mo PII ///
			fp_messaging media_cellphone_often ///
			i.gem_pv_index i.gem_rh_index i.gem_sr_index ///
			decision_numchildren_yn decision_contraception_yn decision_selfill_yn i.couple_communication_index ///
			media_radio_often media_tv_often
			
	** PutExcel
	putexcel (A1)=etable
			
* Run the regression
logistic mcp ur ib24.county_id ///
			age_cat inunion education i.religion ///
			heard_combined_gt_median ///
			visited_provider_12mo PII ///
			fp_messaging media_cellphone_often ///
			i.gem_pv_index i.gem_rh_index i.gem_sr_index ///
			decision_numchildren_yn decision_contraception_yn decision_selfill_yn i.couple_communication_index ///
			media_radio_often media_tv_often
			
	** PutExcel
	putexcel (J1)=etable
			
	
/* Run the regression
. logistic cp ur ib24.county_id ///
>                         age_cat inunion education i.religion ///
>                         heard_combined_gt_median ///
>                         visited_provider_12mo PII ///
>                         fp_messaging media_cellphone_often ///
>                         i.gem_pv_index i.gem_rh_index i.gem_sr_index ///
>                         decision_numchildren_yn decision_contraception_yn decision_selfill_yn i.couple_communication_index ///
>                         media_radio_often media_tv_often

Logistic regression                             Number of obs     =        480
                                                LR chi2(26)       =     138.87
                                                Prob > chi2       =     0.0000
Log likelihood = -260.00219                     Pseudo R2         =     0.2108

------------------------------------------------------------------------------------------------
                            cp | Odds Ratio   Std. Err.      z    P>|z|     [95% Conf. Interval]
-------------------------------+----------------------------------------------------------------
                            ur |   .7568808   .2095512    -1.01   0.314     .4399087    1.302244
                               |
                     county_id |
                         Bomi  |    .756642   .2942369    -0.72   0.473     .3530873    1.621433
                         Bong  |   .6052362   .1837205    -1.65   0.098     .3338421    1.097258
                               |
                       age_cat |   1.348289   .4253116     0.95   0.343     .7265665    2.502019
                       inunion |   1.076973   .3581258     0.22   0.824     .5612483    2.066592
                     education |   .9855719    .253781    -0.06   0.955     .5949866    1.632561
                               |
                      religion |
                       Muslim  |   .6326179   .2103642    -1.38   0.169     .3296793    1.213923
            Other/Traditional  |   .7152307   .4710434    -0.51   0.611     .1967224     2.60039
                               |
      heard_combined_gt_median |   1.768123   .4375042     2.30   0.021     1.088657    2.871664
         visited_provider_12mo |   4.038904   1.235986     4.56   0.000     2.217067    7.357805
                           PII |    1.68951    .651705     1.36   0.174     .7932712    3.598322
                  fp_messaging |    2.02334    .508931     2.80   0.005     1.235851    3.312621
         media_cellphone_often |   1.836108   .6401372     1.74   0.081     .9271219    3.636298
                               |
                  gem_pv_index |
     Moderate Gender Inequity  |   .7857354    .231107    -0.82   0.412     .4414829    1.398424
         High Gender Inequity  |   .6618913   .2189325    -1.25   0.212     .3461277    1.265718
                               |
                  gem_rh_index |
     Moderate Gender Inequity  |   1.713564    .502608     1.84   0.066      .964347    3.044862
         High Gender Inequity  |   1.540887   .4860191     1.37   0.170     .8304034    2.859254
                               |
                  gem_sr_index |
     Moderate Gender Inequity  |   .8802606   .2608778    -0.43   0.667     .4924311    1.573537
         High Gender Inequity  |   .7962497   .2683999    -0.68   0.499     .4112693    1.541602
                               |
       decision_numchildren_yn |   .8714156   .2974235    -0.40   0.687     .4463763    1.701177
     decision_contraception_yn |    1.01966   .3666432     0.05   0.957     .5039473    2.063124
           decision_selfill_yn |   1.204339   .4893154     0.46   0.647     .5431387    2.670463
                               |
    couple_communication_index |
Moderate Couple Communication  |   .9541291   .2767729    -0.16   0.871     .5403702    1.684701
Frequent Couple Communication  |   1.046956   .3123989     0.15   0.878     .5833651    1.878956
                               |
             media_radio_often |    .760081   .1751061    -1.19   0.234     .4839067    1.193873
                media_tv_often |   .7444105   .2011931    -1.09   0.275     .4382851    1.264353
                         _cons |   .2487948   .2114233    -1.64   0.102     .0470434    1.315782
------------------------------------------------------------------------------------------------
Note: _cons estimates baseline odds.




