** Cleaning
cd $directory

* Set TopFive Program
local topfive $topfive

********************************************************
*** GENERATE SEPARATE FEMALE AND TOTAL DATASETS ***
********************************************************

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

********************************************************
*** CLEAN FEMALE DATA ***
********************************************************

* Read in data
use "1. Data/BAL_baseline_femaleadol_cleaned_dataset_v4.dta", clear

* Only keep sexually active
drop if fp_408_1==1

* Number of people total
gen one=1
egen N_total=count(one)

* Main Decision Maker

	** Number of children
	rename _505_5p decision_numchildren
		gen decision_numchildren_yn=1 if decision_numchildren==1
			replace decision_numchildren_yn=0 if inlist(decision_numchildren, 2, 4)
	
	** Whether to use contraception
	rename _505_6p decision_contraception
		gen decision_contraception_yn=1 if decision_contraception==1
			replace decision_contraception_yn=0 if inlist(decision_contraception, 2, 4)
			
	** Going to health center if ill
	rename _505_7p decision_selfill
		gen decision_selfill_yn=1 if decision_selfill==1
			replace decision_selfill_yn=0 if inlist(decision_selfill, 2, 4)
			
* Number of three decisions
gen count_decision=0
	foreach decision in numchildren contraception selfill {
		replace count_decision=count_decision+1 if decision_`decision'_yn==1
	}
			
* Home Environment
	
	** Stressful
	rename ls_516 home_pregnancy_stressful
		
		*** Impute Missing (Missing >5%)
		egen home_pregnancy_stressful_mode=mode(home_pregnancy_stressful)
		replace home_pregnancy_stressful=home_pregnancy_stressful_mode if home_pregnancy_stressful==-98
	
	** Supportive
	rename ls_515 home_pregnancy_supportive

* FP messaging
rename me_917 fp_messaging
	recode fp_messaging -98=1
	
* Currently Using a Contraceptive Method
gen cp=0
	replace cp=1 if fp_407_trad==1 | fp_407_mod==1 | fp_407_LARC
gen mcp=0
	replace mcp=1 if fp_407_mod==1 | fp_407_LARC
	
* Heard of Contraceptive Methods (self cited)
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
capture rename fp_401_11 heard_self_breastfeed
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
	
* Heard of Contraceptive Methods (prompted)
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
		
* Heard of Contraceptive Methods (Traditional, Short Acting, LARC)
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
											
* Median number of methods known
gen count_heard_combined=0
	foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk {
		replace count_heard_combined=count_heard_combined+1 if heard_combined_`method'==1
		}

egen median_heard_combined=median(count_heard_combined)

gen heard_combined_gt_median=1 if count_heard_combined>=median_heard_combined
	replace heard_combined_gt_median=0 if count_heard_combined<median_heard_combined
	
* Perceived risk of pregnancy
rename fp_411 pregnancy_risk
	recode pregnancy_risk -98=.
	recode pregnancy_risk (0 2 10 20 30=1) (40 50 60=2) (70 80 90 100=3)
	

* Self-efficacy to use FP
rename fp_412 fp_self_efficacy
	recode fp_self_efficacy (0 2 10 20 30=1) (40 50 60=2) (70 80 90 100=3)
	label val fp_self_efficacy beans

* Perceived norms
rename fp_427 fp_perceived_norms

* Favorable FP attitudes median
gen fp_attitude=0
forvalues num = 414/422 {
	replace fp_attitude=fp_attitude+fp_`num'
	}

egen median_fp_attitude_overall=median(fp_attitude)

gen median_fp_attitude=0 if fp_attitude<median_fp_attitude_overall
	replace median_fp_attitude=1 if fp_attitude>=median_fp_attitude_overall
	
* Couple Communication
rename ls_501_1_cat couple_communication

* Seen a family planning provider in the last 12 months
rename fp_403 visited_provider_12mo

* Informed of family planning methods they already knew about
rename fp_404 provider_informed_knowmethods
	recode provider_informed_knowmethods -98=1

* Provider informed of problems/delayed pregnancy
rename fp_405 provider_informed_problems
	recode provider_informed_problems -98=1

* Provider informed in case of problems
rename fp_406 provider_informed_whattodo
	recode provider_informed_whattodo -98=1
	
* Gen Method Information Index
gen MII=0
	replace MII=1 if provider_informed_knowmethods==1 & provider_informed_problems==1 & provider_informed_whattodo==1
	
* Used FP services and quality received
gen used_fp_yesquality=0
	replace used_fp_yesquality=1 if MII==1 & visited_provider_12mo==1
gen used_fp_notquality=0
	replace used_fp_notquality=1 if visited_provider_12mo==1 & MII==0
gen used_fp_quality=1 if visited_provider_12mo==0
	replace used_fp_quality=2 if used_fp_notquality==1
	replace used_fp_quality=3 if used_fp_yesquality==1

	label define fp_quality 1 "Did not use" 2 "Used, poor quality" 3 "Used, good quality"
	label val used_fp_quality fp_quality
	
* Listens to the radio 
rename me_901 media_radio
gen media_radio_often=0 if inlist(media_radio, 5, 4)
	replace media_radio_often=1 if inlist(media_radio, 3, 2, 1)
	
* Watches TV
rename me_902 media_tv
gen media_tv_often=0 if inlist(media_tv, 5, 4)
	replace media_tv_often=1 if inlist(media_tv, 3, 2, 1)
	
* Cellphone possession
rename me_903 media_cellphone
gen media_cellphone_often=0 if media_cellphone==0
	replace media_cellphone_often=1 if media_cellphone>=1
	
* Age
gen age_cat=.
	replace age_cat=1 if inlist(demo_101, 14, 15, 16)
	replace age_cat=2 if inlist(demo_101, 17, 18, 19)
	
* Education
gen education=1 if demo_103==1
	replace education=2 if inlist(demo_103, 2, 3)

* Area of Residence
rename comm_classification ur
	
* Marital Status
gen inunion=0 if marital_status==1
	replace inunion=1 if inlist(marital_status, 2, 3)
	
* Given Birth
gen given_birth=0 if mch_204==.
	replace given_birth=1 if mch_204!=.
	
* Run Couple Communication .do file
do $gem

save "1. Data/baseline_femaledata.dta", replace
		

********************************************************
*** CLEAN FAMILY PLANNING DATA ***
********************************************************

foreach sex in male female_ums female_iu male_total female_ums_total female_iu_total {
	
* Read in data
use "$`sex'_data", clear

* Only keep sexually active
drop if fp_408_1==1

* Generate n 
gen n=1
egen total_n=count(n)

* Heard of Contraceptive Methods (self cited)
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
capture rename fp_401_11 heard_self_breastfeed
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
	
* Heard of Contraceptive Methods (prompted)
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
		
* Heard of Contraceptive Methods (Traditional, Short Acting, LARC)
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
											

										 
* Seen a family planning provider in the last 12 months
rename fp_403 visited_provider_12mo

* Informed of family planning methods they already knew about
rename fp_404 provider_informed_knowmethods
	recode provider_informed_knowmethods -98=1

* Provider informed of problems/delayed pregnancy
rename fp_405 provider_informed_problems
	recode provider_informed_problems -98=1

* Provider informed in case of problems
rename fp_406 provider_informed_whattodo
	recode provider_informed_whattodo -98=1

* Current family planning method
capture rename fp_407_1 current_nomethod
capture rename fp_407_2 current_pills
capture rename fp_407_3 current_injectables
capture rename fp_407_4 current_ec
capture rename fp_407_5 current_mc
capture rename fp_407_6 current_fc
capture rename fp_407_7 current_othermodern
capture rename fp_407_8 current_iud
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
		
* Reasons for not using a family planning method
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
capture rename fp_408_11 reason_sideeffects
capture rename fp_408_12 reason_noteffective
capture rename fp_408_13 reason_distance
capture rename fp_408_14 reason_nointerest
capture rename fp_408_97 reason_other
capture rename fp_408__98 reason_dontknow

	** Generate reasons if they don't exist
	foreach reason in nosex dontknowmethods infecund wanttogetpregnant partner religion family health methodnotavailable price sideeffects noteffective distance nointerest other dontknow {
		capture confirm var reason_`reason'
		if _rc!=0 {
			gen reason_`reason'=0
			}
		}
		
	** Top 5
	global list nosex dontknowmethods infecund wanttogetpregnant partner religion family health methodnotavailable price sideeffects noteffective distance nointerest other dontknow
	global var1 reason
	global var2 r
	global first distance
	global last wanttogetpregnant
	run `topfive'

* Previous method
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
	foreach method in pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads other dk nomethod {
		capture confirm var previous_`method'
		if _rc!=0 {
			gen previous_`method'=0
			}
		}
	
	** Top 5
	global list pills injectables ec mc fc othermodern iud implants fs ms breastfeed withdrawal calendar cyclebeads dk other nomethod
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

* Average length of use
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
	
* Number of beans
	
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
	
* Number of women in the community 

	** Using FP
	rename fp_427 number_usingfp
	
	** Receiving FP advice from provider
	rename fp_428 number_fpadvice
	
* Received family planning advice
rename fp_429 provider_received_advice

* Only keep necessary variables
keep id_random county_id total_n ///
	 heard_self* heard_prompted* heard_combined* ///
	 visited_provider_12mo provider_* ///
	 current_* reason_* previous_* ///
	 lengthof* bean* bean* number_* ///
	 top5_*

save "1. Data/baseline_`sex'data.dta", replace
}	

********************************************************
*** MERGE TOTAL AND ALL COUNTY DATASETS ***
********************************************************

* Setup Total Dataset for Merge
foreach sex in female_iu_total female_ums_total male_total {
	use "1. Data/baseline_`sex'data.dta", replace
	
	rename * *_t
	rename id_random_t id_random
	
	save "1. Data/baseline_`sex'data_v2.dta", replace
	}

* Merge and Save Updated Datasets for Tables
use "1. Data/baseline_female_iudata.dta", clear
	merge 1:1 id_random  using "1. Data/baseline_female_iu_totaldata_v2.dta"
	save "1. Data/baseline_female_iudata_merged.dta", replace
	
use "1. Data/baseline_female_umsdata.dta", clear
	merge 1:1 id_random using "1. Data/baseline_female_ums_totaldata_v2.dta"
	save "1. Data/baseline_female_umsdata_merged.dta", replace
	
use "1. Data/baseline_maledata.dta", clear
	merge 1:1 id_random using "1. Data/baseline_male_totaldata_v2.dta"
	save "1. Data/baseline_maledata_merged.dta", replace

	
	
	
