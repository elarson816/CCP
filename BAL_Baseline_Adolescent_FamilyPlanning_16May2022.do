clear
capture log close

/*** Breathrough Action Liberia - Adolescent Family Planning ***

* Input: BAL_baseline_femaleadol_cleaned_dataset_v3.dta 
		 BAL_baseline_male adolescent_cleaned_dataset_v2.dta
* Output: BAL Baseline Dummy Tables Adolescent 02.xls		 
*/

********************************************************
*** SET UP DO FILE ***
********************************************************

* Directory
cd "/Users/Beth/Documents/CCP/BA_Adolsecents_Liberia"

* Datasets
global male_data "1. Data/BAL_baseline_maleadolescent_cleaned_dataset_v4.dta"
global male_total_data "1. Data/BAL_baseline_maleadolescent_total_cleaned_dataset_v4.dta"
global female_ums_data "1. Data/BAL_baseline_femaleadol_umsexactive_cleaned_dataset_v4.dta"
global female_ums_total_data "1. Data/BAL_baseline_femaleadol_umsexactive_total_cleaned_dataset_v4.dta"
global female_iu_data "1. Data/BAL_baseline_femaleadol_inunion_cleaned_dataset_v4.dta.dta"
global female_iu_total_data "1. Data/BAL_baseline_femaleadol_inunion_total_cleaned_dataset_v4.dta.dta"

* Date 
local today=c(current_date)
local c_today= "`today'"
global date=subinstr("`c_today'", " ", "",.)

di %td_CY-N-D  date("$S_DATE", "DMY")
local newdate: di %td_CY-N-D date("$S_DATE","DMY")
di "`newdate'"
di trim("`newdate'")
global todaysdate: di trim("`newdate'")

* Set Put Excel
global putexcel_set "2. Analysis/BAL Baseline Dummy Tables Adolescent 02.xlsx"

* Set TopFive Program
local topfive "/Users/Beth/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baselin_TopFive.do"

* Set macros for cells
local 1cellnum 5
local female_ums_bongcol1 "B"
local female_ums_bombicol1 "C"
local female_ums_totalcol1 "D"
local female_ums_controlol1 "E"
local female_iu_bongcol1 "F"
local female_iu_bombicol1 "G"
local female_iu_totalcol1 "H"
local female_iu_controlol1 "I"
local male_bongcol1 "J"
local male_bombicol1 "K"
local male_totalcol1 "L"
local male_controlol1 "M"

* Set macros for sheets
local part_1 "Part 1"
local part_2 "Part 2"
local part_3 "Part 3"
local part_4 "Part 4"

* Create log
log using "2. Analysis/BAL_Baseline_Adolescent_FamilyPlanning_$date.log", replace

********************************************************
*** READ IN AND PREPARE DATA ***
********************************************************
/*
* Generate separate female and total datasets

* Females
use "1. Data/BAL_baseline_femaleadol_cleaned_dataset_v4.dta", clear

	** Unmarried Sexually Active
	preserve
	
	keep if marital_status==1
	gen id_random=_n
	
	save "1. Data/BAL_baseline_femaleadol_umsexactive_cleaned_dataset_v4.dta", replace
	
	recode county_id (22 23=22)
	drop if county_id==24
	
	save "1. Data/BAL_baseline_femaleadol_umsexactive_total_cleaned_dataset_v4.dta", replace
	
	restore
	
	** Married
	preserve
	
	keep if inlist(marital_status, 2, 3)
	gen id_random=_n
	
	save "1. Data/BAL_baseline_femaleadol_inunion_cleaned_dataset_v4.dta.dta", replace
	
	recode county_id (22 23=22)
	drop if county_id==24
	
	save "1. Data/BAL_baseline_femaleadol_inunion_total_cleaned_dataset_v4.dta.dta", replace
	
	restore
	
* Males
use "1. Data/BAL_baseline_male adolescent_cleaned_dataset_v4.dta"

gen id_random=_n

save "1. Data/BAL_baseline_maleadolescent_cleaned_dataset_v4.dta", replace

preserve

recode county_id (22 23=22)
drop if county_id==24 

save "1. Data/BAL_baseline_maleadolescent_total_cleaned_dataset_v4.dta", replace
	
restore	

foreach sex in male female_ums female_iu male_total male_total female_ums_total female_iu_total {
	
* Read in data
use "$`sex'_data", clear
*use "1. Data/BAL_baseline_femaleadol_cleaned_dataset_v3.dta", clear

* Rename and clean variables

	** Only keep sexually active
	drop if fp_408_1==1

	** Number of people per region
	gen one=1
	egen N_region=count(one), by(county_id)
	
	** Generate total region 
	gen county_total=1 if inlist(county_id, 22, 23)
	egen N_region_total=count(one) if county_total==1
	
	*** Heard of Contraceptive Methods (self cited)
	capture rename fp_401_1 heard_self_pills
	capture rename fp_401_2 heard_self_injectables
	capture rename fp_401_3 heard_self_ec
	capture rename fp_401_4 heard_self_mc
	capture rename fp_401_5 heard_self_fc
	capture rename fp_401_6 heard_self_othermodern
	capture rename fp_401_7 heard_self_iud
	capture rename fp_401_8 heard_self_implants
	capture rename fp_401_9 heard_self_fs
	capture rename fp_401_10 heard_self_ms
	capture rename fp_401_11 heard_self_breasfeed
	capture rename fp_401_13 heard_self_calendar
	capture rename fp_401_14 heard_self_cyclebeads
	capture rename fp_401__98 heard_self_dk
	
		** Generate methods if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk {
			capture confirm var heard_self_`method'
			if _rc!=0 {
				gen heard_self_`method'=0
				}
			}
			
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk
		global var1 heard_self
		global var2 heard_self
		global first breastfeed
		global last withdrawal
		run `topfive'
		
	*** Heard of Contraceptive Methods (prompted)
	capture rename fp_402_1 heard_prompted_pills
	capture rename fp_402_2 heard_prompted_injectables
	capture rename fp_402_3 heard_prompted_ec
	capture rename fp_402_4 heard_prompted_mc
	capture rename fp_402_5 heard_prompted_fc
	capture rename fp_402_6 heard_prompted_othermodern
	capture rename fp_402_7 heard_prompted_iud
	capture rename fp_402_8 heard_prompted_implants
	capture rename fp_402_9 heard_prompted_fs
	capture rename fp_402_10 heard_prompted_ms
	capture rename fp_402_11 heard_prompted_breastfeed
	capture rename fp_402_12 heard_prompted_withdrawal
	capture rename fp_402_13 heard_prompted_calendar
	capture rename fp_402_14 heard_prompted_cyclebeads
	capture rename fp_402__98 heard_prompted_dk
	
		** Generate methods if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk {
			capture confirm var heard_prompted_`method'
			if _rc!=0 {
				gen heard_prompted_`method'=0
				}
			}
		
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk
		global var1 heard_prompted
		global var2 hrd_prmpt
		global first breastfeed
		global last withdrawal
		run `topfive'
			
	*** Heard of Contraceptive Methods (Traditional, Short Acting, LARC)
	foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk {
		gen heard_combined_`method'=0
			replace heard_combined_`method'=1 if heard_self_`method'==1 | heard_prompted_`method'==1
			
		label val heard_combined_`method' noyesdk
		label var heard_combined_`method' "Heard of Method: `method'"
		}
		
	gen heard_combined_larc=0
		replace heard_combined_larc=1 if heard_combined_iud==1 /// IUD
									   | heard_combined_implants==1 /// Implants
									   | heard_combined_fs==1 /// Female Sterilizaiton
									   | heard_combined_ms==1 // Male Sterilization
	gen heard_combined_shortacting=0
		replace heard_combined_shortacting=1 if heard_combined_pills==1 /// Pills
											  | heard_combined_injectables==1 /// Injectables
											  | heard_combined_ec==1 /// Emergency Contraception
											  | heard_combined_mc==1 /// Male Condoms
											  | heard_combined_fc==1 /// Female Condoms
											  | heard_combined_othermodern==1 /// Other Modern
											  | heard_combined_cyclebeads==1 // Cycle Beads
	gen heard_combined_traditional=0
		replace heard_combined_traditiona=1 if heard_combined_withdrawal==1 /// Withdrawal
											 | heard_combined_calendar==1 /// Calendar
											 | heard_combined_breastfeed==1 // Breastfeeding
											 
	
											 
	*** Seen a family planning provider in the last 12 months
	rename fp_403 visited_provider_12mo
	
	*** Informed of family planning methods they already knew about
	rename fp_404 provider_informed_knowmethods
		recode provider_informed_knowmethods -98=.
	
	*** Provider informed of problems/delayed pregnancy
	rename fp_405 provider_informed_problems
		recode provider_informed_problems -98=.
	
	*** Provider informed in case of problems
	rename fp_406 provider_informed_whattodo
		recode provider_informed_whattodo -98=.
	
	*** Current family planning method
	capture rename fp_407_1 current_nomethod
	capture rename fp_407_2 current_pills
	capture rename fp_407_3 current_injectables
	capture rename fp_407_4 current_ec
	capture rename fp_407_5 current_mc
	capture rename fp_407_6 current_fc
	capture rename fp_407_7 current_othermodern
	capture rename fp_408_8 current_iud
	capture rename fp_407_9 current_implants
	capture rename fp_407_10 current_fs
	capture rename fp_407_11 current_ms
	capture rename fp_407_12 current_breastfeed
	capture rename fp_407_13 current_withdrawal
	capture rename fp_407_14 current_calendar
	capture rename fp_407_15 current_cyclebeads
	capture rename fp_407__98 current_dk
	
		** Generate methods if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk {
			capture confirm var current_`method'
			if _rc!=0 {
				gen current_`method'=0
				}
			}
			
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk nomethod 
		global var1 current
		global var2 current
		global first breastfeed
		global last withdrawal
		run `topfive'
		
	capture rename fp_407_trad current_traditional
	capture rename fp_407_mod current_shortacting
	capture rename fp_407_LARC current_larc

		** Generate methods if they don't exist
		foreach method in traditional shortacting larc {
			capture confirm var current_`method'
			if _rc!=0 {
				gen current_`method'=0
				}
			}
			
	*** Reasons for not using a family planning method
	capture rename fp_408_1 reason_nosex
	capture rename fp_408_2 reason_dontknowmethods
	capture rename fp_408_3 reason_infecund
	capture rename fp_408_4 reason_wanttogetpregnant
	capture rename fp_408_5 reason_partner
	capture rename fp_408_6 reason_religion
	capture rename fp_408_7 reason_family
	capture rename fp_408_8 reason_health
	capture rename fp_408_9 reaason_methodnotavailable
	capture rename fp_408_10 reason_price
	capture rename fp_408_11 reason_sideffects
	capture rename fp_408_12 reason_noteffective
	capture rename fp_408_13 reason_distance
	capture rename fp_408_14 reason_noreason
	capture rename fp_408_97 reason_other
	capture rename fp_408__98 reason_dontknow
	
		** Generate reasons if they don't exist
		foreach reason in nosex dontknowmethods infecund wanttogetpregnant partner religion family health methodnotavailable price sideeffects noteffective distance noreason other dontknow {
			capture confirm var reason_`reason'
			if _rc!=0 {
				gen reason_`reason'=0
				}
			}
			
		** Top 5
		global list nosex dontknowmethods infecund wanttogetpregnant partner religion family health methodnotavailable price sideeffects noteffective distance noreason other dontknow
		global var1 reason
		global var2 r
		global first distance
		global last wanttogetpregnant
		run `topfive'
	
	*** Previous method
	capture rename fp_409_1 previous_nomethod
	capture rename fp_409_2 previous_pills
	capture rename fp_409_3 previous_injectables
	capture rename fp_409_4 previous_ec
	capture rename fp_409_5 previous_mc
	capture rename fp_409_6 previoius_fc
	capture rename fp_409_7 previous_othermodern
	capture rename fp_409_8 previous_iud
	capture rename fp_409_9 previous_implants
	capture rename fp_409_10 previous_fs
	capture rename fp_409_11 previous_ms
	capture rename fp_409_12 previous_breastfeed
	capture rename fp_409_13 previous_withdrawal
	capture rename fp_409_14 previous_calendar
	capture rename fp_409_15 previous_cyclebeads
	capture rename fp_409_97 previous_other
	capture rename fp_409__98 previous_dk
		
		
		** Generate reasons if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk nomethod {
			capture confirm var previous_`method'
			if _rc!=0 {
				gen previous_`method'=0
				}
			}
		
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk nomethod
		global var1 previous
		global var2 prev
		global first breastfeed
		global last withdrawal
		run `topfive'	
		
	capture rename fp_409_trad previous_traditional
	capture rename fp_409_mod previous_shortacting
	capture rename fp_409_LARC previous_larc
	
		** Generate reasons if they don't exist
		foreach method in traditional shortacting larc {
			capture confirm var previous_`method'
			if _rc!=0 {
				gen previous_`method'=0
				}
			}
	
	*** Average length of use
	gen lengthofuse_6mo=0
		replace lengthofuse_6mo=1 if fp_410<=6
	gen lengthofuse_6mo_12mo=0
		replace lengthofuse_6mo_12mo=1 if (fp_410>6 & fp_410<=12)
	gen lengthofuse_13mo_24mo=0
		replace lengthofuse_13mo_24mo=1 if (fp_410>12 & fp_410<=24)
	gen lengthofuse_25mo_48mo=0
		replace lengthofuse_25mo_48mo=1 if (fp_410>15 & fp_410<=48)
	gen lengthofuse_49mo=0
		replace lengthofuse_49mo=1 if fp_410>48
	
	gen lengthofuse=0
		replace lengthofuse=1 if lengthofuse_6mo==1
		replace lengthofuse=2 if lengthofuse_6mo_12mo==1
		replace lengthofuse=3 if lengthofuse_13mo_24mo==1
		replace lengthofuse=4 if lengthofuse_25mo_48mo==1
		replace lengthofuse=5 if lengthofuse_49mo==1
		
	label define lengthofuse 1 "<=6 months" 2 "6 - 12 months" 3 "13 - 24 months" 4 "25-48 months" 5 "49+ months"
		label val lengthofuse lengthofuse
		
	*** Number of beans
		
		** Unwanted belly
		rename fp_411 bean_unwantedbelly
	
		** FP help prevent
		rename fp_412 bean_fphelpprevent
	
		** Modern FP improves lives
		rename fp_413 bean_fpimproveslives
	
		** Cause problems to the womb
		rename fp_414 bean_wombproblems
		
		** Short birth intervals are a problem
		rename fp_415 bean_shortintervals
		
		** Post-partum family planning
		rename fp_416 bean_ppfp
		
		** FP gives more time
		rename fp_417 bean_fpprovidestime
		
		** FP reduces labido
		rename fp_418 bean_fpreduceslabido
		
		** Woman's duty to avoid getting pregnant
		rename fp_419 bean_womansduty
		
		** Women can suggest condom use like men
		rename fp_420 bean_suggestcondoms
		
		** Talk about FP with partner
		rename fp_421 bean_discussfp
		
		** Agree about FP with partner
		rename fp_422 bean_agreeonfp
		
		** Men can be mad if wife asks to use a condom
		rename fp_424 bean_husbandangrycondom
		
		** Real women give birth
		rename fp_425 bean_womanbirth
		
		** Real men have child
		rename fp_426 bean_manchild
		
		** Talked about important needs with provider
		rename fp_430 bean_providerneeds
		
		** Provider answered all questions
		rename fp_431 bean_providerquestions
		
		** Provider gave good information
		rename fp_432 bean_providerinfo
		
		** Provider helped to make own choices
		rename fp_433 bean_providerautonomy
		
		** Collapse beans questions
		foreach bean in unwantedbelly fphelpprevent fpimproveslives wombproblems ///
						shortintervals ppfp fpprovidestime ///
						fpreduceslabido womansduty ///
						suggestcondoms discussfp agreeonfp husbandangrycondom ///
						womanbirth manchild ///
						providerneeds providerquestions providerinfo providerautonomy {
			gen bean_`bean'_0_33=0 if bean_`bean'!=.
				replace bean_`bean'_0_33=1 if bean_`bean'<=33
			gen bean_`bean'_34_66=0 if bean_`bean'!=.
				replace bean_`bean'_34_66=1 if bean_`bean'>33 & bean_`bean'<=66
			gen bean_`bean'_67_100=0 if bean_`bean'!=.
				replace bean_`bean'_67_100=1 if bean_`bean'>66
			}
		
	*** Number of women in the community 
	
		** Using FP
		rename fp_427 number_usingfp
		
		** Receiving FP advice from provider
		rename fp_428 number_fpadvice
		
	*** Received family planning advice
	rename fp_429 provider_received_advice
	
* Only keep necessary variables
keep id_random county_id N_region ///
	 heard_self* heard_prompted* heard_combined* ///
	 visited_provider_12mo provider_* ///
	 current_* reason_* previous_* ///
	 lengthof* bean* bean* number_* ///
	 top5_*

save "1. Data/baseline_`sex'data.dta", replace
}	

* Merge Total into datasets
foreach sex in female_iu_total female_ums_total male_total {
	use "1. Data/baseline_`sex'data.dta", replace
	
	rename * *_t
	rename id_random_t id_random
	
	save "1. Data/baseline_`sex'data_v2.dta", replace
	}

use "1. Data/baseline_female_iudata.dta", clear
	merge 1:1 id_random  using "1. Data/baseline_female_iu_totaldata_v2.dta"
	save "1. Data/baseline_female_iudata_merged.dta", replace
	
use "1. Data/baseline_female_umsdata.dta", clear
	merge 1:1 id_random using "1. Data/baseline_female_ums_totaldata_v2.dta"
	save "1. Data/baseline_female_umsdata_merged.dta", replace
	
use "1. Data/baseline_maledata.dta", clear
	merge 1:1 id_random using "1. Data/baseline_male_totaldata_v2.dta"
	save "1. Data/baseline_maledata_merged.dta", replace


*/
********************************************************
*** PART 1: CONTRACEPTIVE USE ***
********************************************************

foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_1'")
local N `1cellnum'

* Create globals to use in Top 5
	
	** Counties
	foreach county in 22 23 24 {
		sum N_region if county_id==`county'
		matrix n_`county'=r(N)
		local n_`county'=r(N)
		}
		
		local region_22=1
		local region_23=`n_22'+`region_22'
		local region_24=`n_23'+`region_23'
	
	** Totals
	sum N_region_t
	matrix n_total=r(N)
	local n_total=r(N)
	
* Enter Ns
local celltext N=`n_23'

di "`celltext'"
putexcel ``group'_bongcol1'`N'="`celltext'"

local celltext N=`n_22'
putexcel ``group'_bombicol1'`N'="`celltext'"

local celltext N=`n_total'
putexcel ``group'_totalcol1'`N'="`celltext'"

local celltext N=`n_24'
putexcel ``group'_controlol1'`N'="`celltext'"

local N=`N'+2

* 401 Which family planning methods have you heard of (heard_self)
	
	** Counties
	foreach county in 22 23 24 {
		foreach top5 in 1 2 3 4 5 {
			
			** Item
			local top5_`top5'_heard_self_item_`county'=top5_`top5'_heard_self_item[`region_`county'']
			
			** Percent
			sum    top5_`top5'_heard_self if county_id==`county'
			matrix top5_`top5'_heard_self_`county'=r(mean)
			matrix top5_`top5'_heard_self_`county'_total=r(N)
			}
		}
	
	** Total
	foreach top5 in 1 2 3 4 5 {
			
		** Item
		local top5_`top5'_heard_self_item_t=top5_`top5'_heard_self_item_t[1]
		
		** Percent
		sum    top5_`top5'_heard_self_t 
		matrix top5_`top5'_heard_self_t=r(mean)
		matrix top5_`top5'_heard_self_t_total=r(N)
		}
	
		** Put Excel
		
			*** Totals
			putexcel ``group'_bongcol1'`N'=  matrix(top5_1_heard_self_23_total)
			putexcel ``group'_bombicol1'`N'= matrix(top5_1_heard_self_22_total)
			putexcel ``group'_totalcol1'`N'= matrix(top5_1_heard_self_t_total)
			putexcel ``group'_controlol1'`N'=matrix(top5_1_heard_self_24_total)
				local N=`N'+1
			
			*** Items and Percents
			forval n = 1/5 {
				putexcel ``group'_bongcol1'`N'=  "`top5_`n'_heard_self_item_23'"
				putexcel ``group'_bombicol1'`N'= "`top5_`n'_heard_self_item_22'"
				putexcel ``group'_totalcol1'`N'= "`top5_`n'_heard_self_item_t'"
				putexcel ``group'_controlol1'`N'="`top5_`n'_heard_self_item_24'"
					local N=`N'+1
					
				putexcel ``group'_bongcol1'`N'=  matrix(top5_`n'_heard_self_23)
				putexcel ``group'_bombicol1'`N'= matrix(top5_`n'_heard_self_22)
				putexcel ``group'_totalcol1'`N'= matrix(top5_`n'_heard_self_t)
				putexcel ``group'_controlol1'`N'=matrix(top5_`n'_heard_self_24)
					local N=`N'+1	
				}
			
* 402 Have you heard about these family planning (heard_prompted)

	** Counties
	foreach county in 22 23 24 {
		foreach top5 in 1 2 3 4 5 {
			
			** Item
			local top5_`top5'_heard_prompted_item_`county'=top5_`top5'_heard_prompted_item[`region_`county'']
			
			** Percent
			sum    top5_`top5'_heard_prompted if county_id==`county'
			matrix top5_`top5'_heard_prompted_`county'=r(mean)
			matrix top5_`top5'_heard_prompted_`county'_total=r(N)
			}
		}	
		
	** Total
	foreach top5 in 1 2 3 4 5 {
			
		** Item
		local top5_`top5'_heard_prompted_item_t=top5_`top5'_heard_prompted_item_t[1]
		
		** Percent
		sum    top5_`top5'_heard_prompted_t 
		matrix top5_`top5'_heard_prompted_t=r(mean)
		matrix top5_`top5'_heard_prompted_t_total=r(N)
		}
	
	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(top5_1_heard_prompted_23_total)
		putexcel ``group'_bombicol1'`N'= matrix(top5_1_heard_prompted_22_total)
		putexcel ``group'_totalcol1'`N'= matrix(top5_1_heard_prompted_t_total)
		putexcel ``group'_controlol1'`N'=matrix(top5_1_heard_prompted_24_total)
			local N=`N'+1
		
		*** Items and Percents
		forval n = 1/5 {
			putexcel ``group'_bongcol1'`N'=  "`top5_`n'_heard_prompted_item_23'"
			putexcel ``group'_bombicol1'`N'= "`top5_`n'_heard_prompted_item_22'"
			putexcel ``group'_totalcol1'`N'= "`top5_`n'_heard_prompted_item_t'"
			putexcel ``group'_controlol1'`N'="`top5_`n'_heard_prompted_item_24'"
				local N=`N'+1
				
			putexcel ``group'_bongcol1'`N'=  matrix(top5_`n'_heard_prompted_23)
			putexcel ``group'_bombicol1'`N'= matrix(top5_`n'_heard_prompted_22)
			putexcel ``group'_totalcol1'`N'= matrix(top5_`n'_heard_prompted_t)
			putexcel ``group'_controlol1'`N'=matrix(top5_`n'_heard_prompted_24)
				local N=`N'+1	
			}

* 401 and 402 LARCS, Short Acting and Traditional

	** Counties
	foreach county in 22 23 24 {
		
		** Traditional
		sum heard_combined_traditional if county_id==`county'
			matrix heard_combined_trad_n_`county'=r(N)
			matrix heard_combined_trad_p_`county'=r(mean)
		
		** Short-Acting
		sum heard_combined_shortacting if county_id==`county'
			matrix heard_combined_short_n_`county'=r(N)
			matrix heard_combined_short_p_`county'=r(mean)
			
		** LARC
		sum heard_combined_larc if county_id==`county'
			matrix heard_combined_larc_n_`county'=r(N)
			matrix heard_combined_larc_p_`county'=r(mean)
			
		}
	
	** Total 
	
		** Traditional
		sum heard_combined_traditional_t
			matrix heard_combined_trad_n_t=r(N)
			matrix heard_combined_trad_p_t=r(mean)
		
		** Short-Acting
		sum heard_combined_shortacting_t 
			matrix heard_combined_short_n_t=r(N)
			matrix heard_combined_short_p_t=r(mean)
			
		** LARC
		sum heard_combined_larc_t 
			matrix heard_combined_larc_n_t=r(N)
			matrix heard_combined_larc_p_t=r(mean)

	** Put Excel 
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(heard_combined_trad_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(heard_combined_trad_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(heard_combined_trad_n_t)
		putexcel ``group'_controlol1'`N'=matrix(heard_combined_trad_n_24)
			local N=`N'+1
			
		*** Percents
		foreach type in trad short larc {
			putexcel ``group'_bongcol1'`N'=  matrix(heard_combined_`type'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(heard_combined_`type'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(heard_combined_`type'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(heard_combined_`type'_p_24)
				local N=`N'+1		
			}

* 407 What family planning methods are you currently using (current)			

	** Counties
	foreach county in 22 23 24 {
		foreach top5 in 1 2 3 4 5 {
			
			** Item
			local top5_`top5'_current_item_`county'=top5_`top5'_current_item[`region_`county'']
			
			** Percent
			sum    top5_`top5'_current if county_id==`county'
			matrix top5_`top5'_current_`county'=r(mean)
			matrix top5_`top5'_current_`county'_total=r(N)
			}
		}
	
	** Total
	foreach top5 in 1 2 3 4 5 {
			
		** Item
		local top5_`top5'_current_item_t=top5_`top5'_current_item[1]
		
		** Percent
		sum    top5_`top5'_current_t
		matrix top5_`top5'_current_t=r(mean)
		matrix top5_`top5'_current_t_total=r(N)
		}
		
	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(top5_1_current_23_total)
		putexcel ``group'_bombicol1'`N'= matrix(top5_1_current_22_total)
		putexcel ``group'_totalcol1'`N'= matrix(top5_1_current_t_total)
		putexcel ``group'_controlol1'`N'=matrix(top5_1_current_24_total)
			local N=`N'+1
		
		*** Items and Percents
		forval n = 1/5 {
			putexcel ``group'_bongcol1'`N'=  "`top5_`n'_current_item_23'"
			putexcel ``group'_bombicol1'`N'= "`top5_`n'_current_item_22'"
			putexcel ``group'_totalcol1'`N'= "`top5_`n'_current_item_t'"
			putexcel ``group'_controlol1'`N'="`top5_`n'_current_item_24'"
				local N=`N'+1
				
			putexcel ``group'_bongcol1'`N'=  matrix(top5_`n'_current_23)
			putexcel ``group'_bombicol1'`N'= matrix(top5_`n'_current_22)
			putexcel ``group'_totalcol1'`N'= matrix(top5_`n'_current_t)
			putexcel ``group'_controlol1'`N'=matrix(top5_`n'_current_24)
				local N=`N'+1	
			}
			
* 407 No Method, Traditional, Short Acting and LARC

	** Counties
	foreach county in 22 23 24 {
		
		** No Method
		sum current_nomethod if county_id==`county'
			matrix current_none_n_`county'=r(N)
			matrix current_none_p_`county'=r(mean)
		
		** Traditional
		sum current_traditional if county_id==`county'
			matrix current_trad_n_`county'=r(N)
			matrix current_trad_p_`county'=r(mean)
		
		** Short-Acting
		sum current_shortacting if county_id==`county'
			matrix current_short_n_`county'=r(N)
			matrix current_short_p_`county'=r(mean)
			
		** LARC
		sum current_larc if county_id==`county'
			matrix current_larc_n_`county'=r(N)
			matrix current_larc_p_`county'=r(mean)
			
		}
		
	** Total
	
		** No Method
		sum current_nomethod_t
			matrix current_none_n_t=r(N)
			matrix current_none_p_t=r(mean)
		
		** Traditional
		sum current_traditional_t
			matrix current_trad_n_t=r(N)
			matrix current_trad_p_t=r(mean)
		
		** Short-Acting
		sum current_shortacting_t
			matrix current_short_n_t=r(N)
			matrix current_short_p_t=r(mean)
			
		** LARC
		sum current_larc_t
			matrix current_larc_n_t=r(N)
			matrix current_larc_p_t=r(mean)

	** Put Excel 
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(current_none_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(current_none_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(current_none_n_t)
		putexcel ``group'_controlol1'`N'=matrix(current_none_n_24)
			local N=`N'+1
			
		*** Percents
		foreach type in none trad short larc {
			putexcel ``group'_bongcol1'`N'=  matrix(current_`type'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(current_`type'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(current_`type'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(current_`type'_p_24)
				local N=`N'+1		
			}			
			
* 408 Reasons for not using fp (reason)

	** Counties
	foreach county in 22 23 24 {
		foreach top5 in 1 2 3 4 5 {
			
			** Item
			local top5_`top5'_reason_item_`county'=top5_`top5'_reason_item[`region_`county'']
			
			** Percent
			sum    top5_`top5'_reason if county_id==`county'
			matrix top5_`top5'_reason_`county'=r(mean)
			matrix top5_`top5'_reason_`county'_total=r(N)
			}
		}
		
	** Total
	foreach top5 in 1 2 3 4 5 {
			
		** Item
		local top5_`top5'_reason_item_t=top5_`top5'_reason_item_t[1]
		
		** Percent
		sum    top5_`top5'_reason_t
		matrix top5_`top5'_reason_t=r(mean)
		matrix top5_`top5'_reason_t_total=r(N)
		}

	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(top5_1_reason_23_total)
		putexcel ``group'_bombicol1'`N'= matrix(top5_1_reason_22_total)
		putexcel ``group'_totalcol1'`N'= matrix(top5_1_reason_t_total)
		putexcel ``group'_controlol1'`N'=matrix(top5_1_reason_24_total)
			local N=`N'+1
		
		*** Items and Percents
		forval n = 1/5 {
			putexcel ``group'_bongcol1'`N'=  "`top5_`n'_reason_item_23'"
			putexcel ``group'_bombicol1'`N'= "`top5_`n'_reason_item_22'"
			putexcel ``group'_totalcol1'`N'= "`top5_`n'_reason_item_t'"
			putexcel ``group'_controlol1'`N'="`top5_`n'_reason_item_24'"
				local N=`N'+1
				
			putexcel ``group'_bongcol1'`N'=  matrix(top5_`n'_reason_23)
			putexcel ``group'_bombicol1'`N'= matrix(top5_`n'_reason_22)
			putexcel ``group'_totalcol1'`N'= matrix(top5_`n'_reason_t)
			putexcel ``group'_controlol1'`N'=matrix(top5_`n'_reason_24)
				local N=`N'+1	
			}

* 409 Previous family planning method (previous)			

	** Counties
	foreach county in 22 23 24 {
		foreach top5 in 1 2 3 4 5 {
			
			** Item
			local top5_`top5'_previous_item_`county'=top5_`top5'_previous_item[`region_`county'']
			
			** Percent
			sum    top5_`top5'_previous if county_id==`county'
			matrix top5_`top5'_previous_`county'=r(mean)
			matrix top5_`top5'_previous_`county'_total=r(N)
			}
		}
		
	** Total
	foreach top5 in 1 2 3 4 5 {
		
		** Item
		local top5_`top5'_previous_item_t=top5_`top5'_previous_item_t[1]
		
		** Percent
		sum    top5_`top5'_previous_t
		matrix top5_`top5'_previous_t=r(mean)
		matrix top5_`top5'_previous_t_total=r(N)
		}

	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(top5_1_previous_23_total)
		putexcel ``group'_bombicol1'`N'= matrix(top5_1_previous_22_total)
		putexcel ``group'_totalcol1'`N'= matrix(top5_1_previous_t_total)
		putexcel ``group'_controlol1'`N'=matrix(top5_1_previous_24_total)
			local N=`N'+1
		
		*** Items and Percents
		forval n = 1/5 {
			putexcel ``group'_bongcol1'`N'=  "`top5_`n'_previous_item_23'"
			putexcel ``group'_bombicol1'`N'= "`top5_`n'_previous_item_22'"
			putexcel ``group'_totalcol1'`N'= "`top5_`n'_previous_item_t'"
			putexcel ``group'_controlol1'`N'="`top5_`n'_previous_item_24'"
				local N=`N'+1
				
			putexcel ``group'_bongcol1'`N'=  matrix(top5_`n'_previous_23)
			putexcel ``group'_bombicol1'`N'= matrix(top5_`n'_previous_22)
			putexcel ``group'_totalcol1'`N'= matrix(top5_`n'_previous_t)
			putexcel ``group'_controlol1'`N'=matrix(top5_`n'_previous_24)
				local N=`N'+1	
			}			
			
* 409 No Method, Traditional, Short Acting and LARC	

	** Counties
	foreach county in 22 23 24 {
		
		** No Method
		sum previous_nomethod if county_id==`county'
			matrix previous_none_n_`county'=r(N)
			matrix previous_none_p_`county'=r(mean)
		
		** Traditional
		sum previous_traditional if county_id==`county'
			matrix previous_trad_n_`county'=r(N)
			matrix previous_trad_p_`county'=r(mean)
		
		** Short-Acting
		sum previous_shortacting if county_id==`county'
			matrix previous_short_n_`county'=r(N)
			matrix previous_short_p_`county'=r(mean)
			
		** LARC
		sum previous_larc if county_id==`county'
			matrix previous_larc_n_`county'=r(N)
			matrix previous_larc_p_`county'=r(mean)
			
		}
		
	** Total
		
		** No Method
		sum previous_nomethod_t
			matrix previous_none_n_t=r(N)
			matrix previous_none_p_t=r(mean)
		
		** Traditional
		sum previous_traditional_t
			matrix previous_trad_n_t=r(N)
			matrix previous_trad_p_t=r(mean)
		
		** Short-Acting
		sum previous_shortacting_t
			matrix previous_short_n_t=r(N)
			matrix previous_short_p_t=r(mean)
			
		** LARC
		sum previous_larc_t
			matrix previous_larc_n_t=r(N)
			matrix previous_larc_p_t=r(mean)

	** Put Excel 
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(previous_none_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(previous_none_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(previous_none_n_t)
		putexcel ``group'_controlol1'`N'=matrix(previous_none_n_24)
			local N=`N'+1
			
		*** Percents
		foreach type in none trad short larc {
			putexcel ``group'_bongcol1'`N'=  matrix(previous_`type'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(previous_`type'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(previous_`type'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(previous_`type'_p_24)
				local N=`N'+1		
			}	
			
* 410 Start using method (lengthofuse)

	** Counties
	foreach county in 22 23 24 {
		
		** <=6 months
		sum lengthofuse_6mo if county_id==`county'
			matrix lengthofuse_6mo_n_`county'=r(N)
			matrix lengthofuse_6mo_p_`county'=r(mean)
		
		** 7-12 months
		sum lengthofuse_6mo_12mo if county_id==`county'
			matrix lengthofuse_6mo_12mo_n_`county'=r(N)
			matrix lengthofuse_6mo_12mo_p_`county'=r(mean)
		
		** 13-24 months
		sum lengthofuse_13mo_24mo if county_id==`county'
			matrix lengthofuse_13mo_24mo_n_`county'=r(N)
			matrix lengthofuse_13mo_24mo_p_`county'=r(mean)
			
		** 15-48 months
		sum lengthofuse_25mo_48mo if county_id==`county'
			matrix lengthofuse_25mo_48mo_n_`county'=r(N)
			matrix lengthofuse_25mo_48mo_p_`county'=r(mean)
			
		** 49 + months
		sum lengthofuse_49mo if county_id==`county'
			matrix lengthofuse_49mo_n_`county'=r(N)
			matrix lengthofuse_49mo_p_`county'=r(mean)
			
		}
		
	** Total
	
		** <=6 months
		sum lengthofuse_6mo_t
			matrix lengthofuse_6mo_n_t=r(N)
			matrix lengthofuse_6mo_p_t=r(mean)
		
		** 7-12 months
		sum lengthofuse_6mo_12mo_t
			matrix lengthofuse_6mo_12mo_n_t=r(N)
			matrix lengthofuse_6mo_12mo_p_t=r(mean)
		
		** 13-24 months
		sum lengthofuse_13mo_24mo_t
			matrix lengthofuse_13mo_24mo_n_t=r(N)
			matrix lengthofuse_13mo_24mo_p_t=r(mean)
			
		** 15-48 months
		sum lengthofuse_25mo_48mo_t
			matrix lengthofuse_25mo_48mo_n_t=r(N)
			matrix lengthofuse_25mo_48mo_p_t=r(mean)
			
		** 49 + months
		sum lengthofuse_49mo_t
			matrix lengthofuse_49mo_n_t=r(N)
			matrix lengthofuse_49mo_p_t=r(mean)

	** Put Excel 
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(lengthofuse_6mo_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(lengthofuse_6mo_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(lengthofuse_6mo_n_t)
		putexcel ``group'_controlol1'`N'=matrix(lengthofuse_6mo_n_24)
			local N=`N'+1
			
		*** Percents
		foreach length in 6mo 6mo_12mo 13mo_24mo 25mo_48mo 49mo  {
			putexcel ``group'_bongcol1'`N'=  matrix(lengthofuse_`length'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(lengthofuse_`length'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(lengthofuse_`length'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(lengthofuse_`length'_p_24)
				local N=`N'+1		
			}	

	}

********************************************************
*** PART 2: Beans ***
********************************************************

foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_2'")
local N `1cellnum'

* Create globals to use in Top 5
	
	** Counties
	foreach county in 22 23 24 {
		sum N_region if county_id==`county'
		matrix n_`county'=r(N)
		local n_`county'=r(N)
		}
		
		local region_22=1
		local region_23=`n_22'+`region_22'
		local region_24=`n_23'+`region_23'
	
	** Totals
	sum N_region_t
	matrix n_total=r(N)
	local n_total=r(N)
	
* Enter Ns
local celltext N=`n_23'

di "`celltext'"
putexcel ``group'_bongcol1'`N'="`celltext'"

local celltext N=`n_22'
putexcel ``group'_bombicol1'`N'="`celltext'"

local celltext N=`n_total'
putexcel ``group'_totalcol1'`N'="`celltext'"

local celltext N=`n_24'
putexcel ``group'_controlol1'`N'="`celltext'"

local N=`N'+2

* All the beans (bean_)
foreach bean in unwantedbelly fphelpprevent fpimproveslives wombproblems ///
				shortintervals ppfp fpprovidestime ///
				fpreduceslabido womansduty ///
				suggestcondoms discussfp agreeonfp husbandangrycondom ///
				womanbirth manchild {
					
	** Counties
	foreach county in 22 23 24 {
		
		** 0-33
		sum bean_`bean'_0_33 if county_id==`county'
			matrix b_`bean'_0_33_n_`county'=r(N)
			matrix b_`bean'_0_33_p_`county'=r(mean)
			
		** 34-66
		sum bean_`bean'_34_66 if county_id==`county'
			matrix b_`bean'_34_66_n_`county'=r(N)
			matrix b_`bean'_34_66_p_`county'=r(mean)

		** 67-100
		sum bean_`bean'_67_100 if county_id==`county'
			matrix b_`bean'_67_100_n_`county'=r(N)
			matrix b_`bean'_67_100_p_`county'=r(mean)	
		}
		
	** Total
		
		** 0-33
		sum bean_`bean'_0_33_t
			matrix b_`bean'_0_33_n_t=r(N)
			matrix b_`bean'_0_33_p_t=r(mean)
			
		** 34-66
		sum bean_`bean'_34_66_t
			matrix b_`bean'_34_66_n_t=r(N)
			matrix b_`bean'_34_66_p_t=r(mean)

		** 67-100
		sum bean_`bean'_67_100_t
			matrix b_`bean'_67_100_n_t=r(N)
			matrix b_`bean'_67_100_p_t=r(mean)

			
	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(b_`bean'_0_33_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(b_`bean'_0_33_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(b_`bean'_0_33_n_t)
		putexcel ``group'_controlol1'`N'=matrix(b_`bean'_0_33_n_24)
			local N=`N'+1
			
		*** Percents
		foreach amount in 0_33 34_66 67_100  {
			putexcel ``group'_bongcol1'`N'=  matrix(b_`bean'_`amount'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(b_`bean'_`amount'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(b_`bean'_`amount'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(b_`bean'_`amount'_p_24)
				local N=`N'+1		
			}	
		}
	}

********************************************************
*** PART 3: Neighborhood ***
********************************************************	
	
foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_3'")
local N `1cellnum'

* Create globals to use in Top 5
	
	** Counties
	foreach county in 22 23 24 {
		sum N_region if county_id==`county'
		matrix n_`county'=r(N)
		local n_`county'=r(N)
		}
		
		local region_22=1
		local region_23=`n_22'+`region_22'
		local region_24=`n_23'+`region_23'
	
	** Totals
	sum N_region_t
	matrix n_total=r(N)
	local n_total=r(N)
	
* Enter Ns
local celltext N=`n_23'

di "`celltext'"
putexcel ``group'_bongcol1'`N'="`celltext'"

local celltext N=`n_22'
putexcel ``group'_bombicol1'`N'="`celltext'"

local celltext N=`n_total'
putexcel ``group'_totalcol1'`N'="`celltext'"

local celltext N=`n_24'
putexcel ``group'_controlol1'`N'="`celltext'"

local N=`N'+2
	
* Number questions (number_)
foreach number in usingfp fpadvice {	
	
	** Counties
	foreach county in 22 23 24 {
	
		tab number_`number' if county_id==`county', matcell(table)
			matrix low=table[1,1]
			matrix medium=table[2,1]
			matrix high=table[3,1]
			matrix total=low[1,1]+medium[1,1]+high[1,1]
		
		matrix number_`number'_n_`county'=total[1,1]
		
		foreach level in low medium high {
			matrix number_`level'_`number'_p_`county'=`level'[1,1]/total[1,1]
			}
		}
		
	** Total
	tab number_`number'_t, matcell(table)
			matrix low=table[1,1]
			matrix medium=table[2,1]
			matrix high=table[3,1]
			matrix total=low[1,1]+medium[1,1]+high[1,1]
		
		matrix number_`number'_n_t=total[1,1]
		
		foreach level in low medium high {
			matrix number_`level'_`number'_p_t=`level'[1,1]/total[1,1]
			}
		
	** Put Excel 
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(number_`number'_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(number_`number'_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(number_`number'_n_t)
		putexcel ``group'_controlol1'`N'=matrix(number_`number'_n_24)
			local N=`N'+1
			
		*** Percents
		foreach level in low medium high  {
			putexcel ``group'_bongcol1'`N'=  matrix(number_`level'_`number'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(number_`level'_`number'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(number_`level'_`number'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(number_`level'_`number'_p_24)
				local N=`N'+1		
			}	
		}
	}
*/
	
********************************************************
*** PART 4: Provider ***
********************************************************	
	
foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_4'")
local N `1cellnum'

* Create globals to use in Top 5
	
	** Counties
	foreach county in 22 23 24 {
		sum N_region if county_id==`county'
		matrix n_`county'=r(N)
		local n_`county'=r(N)
		}
		
		local region_22=1
		local region_23=`n_22'+`region_22'
		local region_24=`n_23'+`region_23'
	
	** Totals
	sum N_region_t
	matrix n_total=r(N)
	local n_total=r(N)
	
* Enter Ns
local celltext N=`n_23'

di "`celltext'"
putexcel ``group'_bongcol1'`N'="`celltext'"

local celltext N=`n_22'
putexcel ``group'_bombicol1'`N'="`celltext'"

local celltext N=`n_total'
putexcel ``group'_totalcol1'`N'="`celltext'"

local celltext N=`n_24'
putexcel ``group'_controlol1'`N'="`celltext'"

local N=`N'+2	
	
* 403 Have you seen a fp provider (visited_provider_12mo)
	** Counties
	foreach county in 22 23 24 {
		
		tab visited_provider_12mo if county_id==`county', matcell(table)
			matrix no=table[1,1]
			matrix yes=table[2,1]
			matrix total=no[1,1]+yes[1,1]
			
		matrix visited_provider_12mo_n_`county'=total[1,1]
		
		foreach response in no yes {
			matrix visited_provider_12mo_`response'_p_`county'=`response'[1,1]/total[1,1]
			}
		}
		
	** Total
	tab visited_provider_12mo_t, matcell(table)
			matrix no=table[1,1]
			matrix yes=table[2,1]
			matrix total=no[1,1]+yes[1,1]
			
		matrix visited_provider_12mo_n_t=total[1,1]
		
		foreach response in no yes {
			matrix visited_provider_12mo_`response'_p_t=`response'[1,1]/total[1,1]
			}
	
	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(visited_provider_12mo_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(visited_provider_12mo_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(visited_provider_12mo_n_t)
		putexcel ``group'_controlol1'`N'=matrix(visited_provider_12mo_n_24)
			local N=`N'+1
			
		*** Percents
		foreach response in no yes  {
			putexcel ``group'_bongcol1'`N'=  matrix(visited_provider_12mo_`response'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(visited_provider_12mo_`response'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(visited_provider_12mo_`response'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(visited_provider_12mo_`response'_p_24)
				local N=`N'+1		
			}

* 404-406 Family Planning Information (provider_informed)
foreach info in knowmethods problems whattodo {	
	
	** Counties
	foreach county in 22 23 24 {
	
		tab provider_informed_`info' if county_id==`county', matcell(table)
			matrix no=table[1,1]
			matrix yes=table[2,1]
			matrix no_p=table[3,1]
			matrix total=no[1,1]+yes[1,1]+no_p[1,1]
		
		matrix info_`info'_n_`county'=total[1,1]
		
		foreach response in no yes no_p {
			matrix info_`response'_`info'_p_`county'=`response'[1,1]/total[1,1]
			}
		}
	
	** Total
	tab provider_informed_`info'_t, matcell(table)
			matrix no=table[1,1]
			matrix yes=table[2,1]
			matrix no_p=table[3,1]
			matrix total=no[1,1]+yes[1,1]+no_p[1,1]
		
		matrix info_`info'_n_t=total[1,1]
		
		foreach response in no yes no_p {
			matrix info_`response'_`info'_p_t=`response'[1,1]/total[1,1]
			}
		
	** Put Excel 
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(info_`info'_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(info_`info'_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(info_`info'_n_t)
		putexcel ``group'_controlol1'`N'=matrix(info_`info'_n_24)
			local N=`N'+1
			
		*** Percents
		foreach response in no yes no_p  {
			putexcel ``group'_bongcol1'`N'=  matrix(info_`response'_`info'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(info_`response'_`info'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(info_`response'_`info'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(info_`response'_`info'_p_24)
				local N=`N'+1		
			}	
		}
	
* 429 Family Planning Advice (provider_received_advice)	
	
	** Counties
	foreach county in 22 23 24 {
		
		tab provider_received_advice if county_id==`county', matcell(table)
			matrix no=table[1,1]
			matrix yes=table[2,1]
			matrix total=no[1,1]+yes[1,1]
			
		matrix received_advice_n_`county'=total[1,1]
		
		foreach response in no yes {
			matrix received_advice_`response'_p_`county'=`response'[1,1]/total[1,1]
			}
		}
		
	** Total
	tab provider_received_advice_t, matcell(table)
			matrix no=table[1,1]
			matrix yes=table[2,1]
			matrix total=no[1,1]+yes[1,1]
			
		matrix received_advice_n_t=total[1,1]
		
		foreach response in no yes {
			matrix received_advice_`response'_p_t=`response'[1,1]/total[1,1]
			}
	
	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(received_advice_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(received_advice_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(received_advice_n_t)
		putexcel ``group'_controlol1'`N'=matrix(received_advice_n_24)
			local N=`N'+1
			
		*** Percents
		foreach response in no yes  {
			putexcel ``group'_bongcol1'`N'=  matrix(received_advice_`response'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(received_advice_`response'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(received_advice_`response'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(received_advice_`response'_p_24)
				local N=`N'+1		
			}
			
* All the beans (bean_)
foreach bean in providerneeds providerquestions providerinfo providerautonomy {
					
	** Counties
	foreach county in 22 23 24 {
		
		** 0-33
		sum bean_`bean'_0_33 if county_id==`county'
			matrix b_`bean'_0_33_n_`county'=r(N)
			matrix b_`bean'_0_33_p_`county'=r(mean)
			
		** 34-66
		sum bean_`bean'_34_66 if county_id==`county'
			matrix b_`bean'_34_66_n_`county'=r(N)
			matrix b_`bean'_34_66_p_`county'=r(mean)

		** 67-100
		sum bean_`bean'_67_100 if county_id==`county'
			matrix b_`bean'_67_100_n_`county'=r(N)
			matrix b_`bean'_67_100_p_`county'=r(mean)	
		}
		
	** Total
		** 0-33
		sum bean_`bean'_0_33_t
			matrix b_`bean'_0_33_n_t=r(N)
			matrix b_`bean'_0_33_p_t=r(mean)
			
		** 34-66
		sum bean_`bean'_34_66_t
			matrix b_`bean'_34_66_n_t=r(N)
			matrix b_`bean'_34_66_p_t=r(mean)

		** 67-100
		sum bean_`bean'_67_100_t
			matrix b_`bean'_67_100_n_t=r(N)
			matrix b_`bean'_67_100_p_t=r(mean)
		
	** Put Excel
	
		*** Totals
		putexcel ``group'_bongcol1'`N'=  matrix(b_`bean'_0_33_n_23)
		putexcel ``group'_bombicol1'`N'= matrix(b_`bean'_0_33_n_22)
		putexcel ``group'_totalcol1'`N'= matrix(b_`bean'_0_33_n_t)
		putexcel ``group'_controlol1'`N'=matrix(b_`bean'_0_33_n_24)
			local N=`N'+1
			
		*** Percents
		foreach amount in 0_33 34_66 67_100  {
			putexcel ``group'_bongcol1'`N'=  matrix(b_`bean'_`amount'_p_23)
			putexcel ``group'_bombicol1'`N'= matrix(b_`bean'_`amount'_p_22)
			putexcel ``group'_totalcol1'`N'= matrix(b_`bean'_`amount'_p_t)
			putexcel ``group'_controlol1'`N'=matrix(b_`bean'_`amount'_p_24)
				local N=`N'+1		
			}
	}
}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
