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
cd "/Users/ealarson/Documents/CCP/BA_Adolsecents_Liberia"

* Datasets
global male_data "1. Data/BAL_baseline_male adolescent_cleaned_dataset_v2.dta"
global female_data "1. Data/BAL_baseline_femaleadol_cleaned_dataset_v3.dta"

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
local topfive "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baselin_TopFive.do"

* Set macros for cells
local 1cellnum 5
local unmar_bongcol1 "B"
local unmar_bombicol1 "C"
local unmar_totalcol1 "D"
local unmar_controlol1 "E"
local mar_bongcol1 "F"
local mar_bombicol1 "G"
local mar_totalcol1 "H"
local mar_controlol1 "I"
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

*foreach sex in female {
	
* Read in data
*use "$`sex'_data", clear
use "1. Data/BAL_baseline_femaleadol_cleaned_dataset_v3.dta", clear

* Rename and clean variables

	** Number of people per region
	gen one=1
	egen N_region=count(one), by(county_id)
	
	** Only keep sexually active
	drop if fp_408_1==1
	
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
	capture rename fp_401_11 heard_self_lam
	capture rename fp_401_13 heard_self_calendar
	capture rename fp_401_14 heard_self_cyclebeads
	capture rename fp_401__98 heard_self_dk
	
		** Generate methods if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk {
			capture confirm var heard_self_`method'
			if _rc!=0 {
				gen heard_self_`method'=0
				}
			}
			
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk
		global var1 heard_self
		global var2 heard_self
		global first calendar
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
	capture rename fp_402_11 heard_prompted_lam
	capture rename fp_402_12 heard_prompted_withdrawal
	capture rename fp_402_13 heard_prompted_calendar
	capture rename fp_402_14 heard_prompted_cyclebeads
	capture rename fp_402__98 heard_prompted_dk
	
		** Generate methods if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk {
			capture confirm var heard_prompted_`method'
			if _rc!=0 {
				gen heard_prompted_`method'=0
				}
			}
		
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk
		global var1 heard_prompted
		global var2 hrd_prmpt
		global first calendar
		global last withdrawal
		run `topfive'
			
	*** Heard of Contraceptive Methods (Traditional, Short Acting, LARC)
	foreach method in pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk {
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
											  | heard_combined_lam==1 /// Breastfeeding
											  | heard_combined_cyclebeads==1 // Cycle Beads
	gen heard_combined_traditional=0
		replace heard_combined_traditiona=1 if heard_combined_withdrawal==1 /// Withdrawal
											 | heard_combined_calendar==1 // Calendar
											 
	
											 
	*** Seen a family planning provider in the last 12 months
	rename fp_403 visited_provider_12mo
	
	*** Informed of family planning methods they already knew about
	rename fp_404 provider_informed_knowmethods
	
	*** Provider informed of problems/delayed pregnancy
	rename fp_405 provider_informed_problems
	
	*** Provider informed in case of problems
	rename fp_406 provider_informed_whattodo
	
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
	capture rename fp_407_12 current_lam
	capture rename fp_407_13 current_withdrawal
	capture rename fp_407_14 current_calendar
	capture rename fp_407_15 current_cyclebeads
	capture rename fp_407__98 current_dk
	capture rename fp_407_trad current_traditional
	capture rename fp_407_mod current_shortacting
	capture rename fp_407_LARC current_larc
	
		** Generate methods if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk {
			capture confirm var current_`method'
			if _rc!=0 {
				gen current_`method'=0
				}
			}
			
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk nomethod 
		global var1 current
		global var2 current
		global first calendar
		global last withdrawal
		run `topfive'
	
		** Composite
		gen cp=0 if current_nomethod==1
			replace cp=14 if current_withdrawal==1
			replace cp=13 if current_calendar==1
			replace cp=12 if current_lam==1
			replace cp=11 if current_cyclebeads==1
			replace cp=10 if current_othermodern==1
			replace cp=9 if current_fc==1
			replace cp=8 if current_mc==1
			replace cp=7 if current_ec==1
			replace cp=6 if current_pills==1
			replace cp=5 if current_injectables==1
			replace cp=4 if current_iud==1
			replace cp=3 if current_implant==1
			replace cp=2 if current_ms==1
			replace cp=1 if current_fs==1
			
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
	capture rename fp_409_12 previous_lam
	capture rename fp_409_13 previous_withdrrawl
	capture rename fp_409_14 previous_calendar
	capture rename fp_409_15 previous_cyclebeads
	capture rename fp_409_97 previous_other
	capture rename fp_409__98 previous_dk
	capture rename fp_409_trad previous_traditional
	capture rename fp_409_mod previous_shortacting
	capture rename fp_409_LARC previous_larc
		
		
		** Generate reasons if they don't exist
		foreach method in pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk nomethod {
			capture confirm var previous_`method'
			if _rc!=0 {
				gen previous_`method'=0
				}
			}
		
		** Top 5
		global list pills injectables ec mc fc othermodern iud implants fs ms lam withdrawal calendar cyclebeads dk nomethod
		global var1 previous
		global var2 prev
		global first calendar
		global last withdrawal
		run `topfive'	
	
	
	*** Average length of use
	gen lengthofuse_6mo=0
		replace lengthofuse_6mo=1 if fp_410<=6
	gen lengthofuse_6mo_12mo=0
		replace lengthofuse_6mo_12mo=1 if 6<fp_410<=12
	gen lengthofuse_13mo_24mo=0
		replace lengthofuse_13mo_24mo=1 if 12<fp_410<=24
	gen lengthofuse_25mo_48mo=0
		replace lengthofuse_25mo_48mo=1 if 15<fp_410<=48
	gen lengthofuse_49mo=0
		replace lengthofuse_49mo=1 if 48<fp_410
	
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
		rename fp_411 numberofbeans_unwantedbelly
	
		** FP help prevent
		rename fp_412 numberofbeans_fphelpprevent
	
		** Modern FP improves lives
		rename fp_413 numberofbeans_fpimproveslives
	
		** Cause problems to the womb
		rename fp_414 numberofbeans_wombproblems
		
		** Short birth intervals are a problem
		rename fp_415 numberofbeans_shortintervals
		
		** Post-partum family planning
		rename fp_416 numberofbeans_ppfp
		
		** FP gives more time
		rename fp_417 numberofbeans_fpprovidestime
		
		** FP reduces labido
		rename fp_418 numberofbeans_fpreduceslabido
		
		** Woman's duty to avoid getting pregnant
		rename fp_419 numberofbeans_womansduty
		
		** Women can suggest condom use like men
		rename fp_420 numberofbeans_suggestcondoms
		
		** Talk about FP with partner
		rename fp_421 numberofbeans_discussfp
		
		** Agree about FP with partner
		rename fp_422 numberofbeans_agreeonfp
		
		** Men can be mad if wife asks to use a condom
		rename fp_424 numberofbeans_husbandangrycondom
		
		** Real women give birth
		rename fp_425 numberofbeans_womanbirth
		
		** Real men have child
		rename fp_426 numberofbeans_manhcild
		
		** Talked about important needs with provider
		rename fp_430 numberofbeans_providerneeds
		
		** Provider answered all questions
		rename fp_431 numberofbeans_providerquestions
		
		** Provider gave good information
		rename fp_432 numberofbeans_providerinfo
		
		** Provider helped to make own choices
		rename fp_433 numberofbeans_providerautonomy
		
	*** Number of women in the community 
	
		** Using FP
		rename fp_427 number_usingfp
		
		** Receiving FP advice from provider
		rename fp_428 number_fpadvice
		
	*** Received family planning advice
	rename fp_429 provider_received_advice
	
	*** 
	
	assert 0
* Only keep necessary variables
keep county_id ///
	 heard_self* heard_prompted* heard_combined* ///
	 visited_provider_12mo provider_* ///
	 current_* cp reason_* previous_* ///
	 lengthof* numberofbeans* number_* ///
	 top5_*


save "1. Data/baseline_`sex'data.dta", replace
}

********************************************************
*** PART 1 ***
********************************************************

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_1'")
local N `1cellnum'

* Female Dataset
use "1. Data/baseline_femaledata.dta", clear

* Generate Subgroup Variable


*** Unmarried Subgroup
foreach subgrp in unmar mar {
	
}







