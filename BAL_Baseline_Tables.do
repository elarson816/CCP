*** Tables
cd $directory

* Set macros for cells (General Tables)
local cellnum1 $cellnum1 
local female_umscol $female_umscol
local female_iucol $female_iucol 
local malecol $malecol

* Set macros for cells (Demographic Table)
local cellnum2 $cellnum2
local female_totalcol $female_totalcol


* Set macros for sheets
local part_1 $part_1
local part_2 $part_2
local part_3 $part_3
local part_4 $part_4
local norms $norms

********************************************************
*** PART 1: CONTRACEPTIVE USE ***
********************************************************
/*
foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_1'")
local N `cellnum1'

* Create counts
gen n=1
sum n
matrix n=r(N)
local n=r(N) 
	
* Enter N
local celltext N=`n'
putexcel ``group'col'`N'="`celltext'"

local N=`N'+2

* 401 Which family planning methods have you heard of (heard_self)
foreach n in 1 2 3 4 5 {
			
	** Item
	local top5_`n'_heard_self_item=top5_`n'_heard_self_item[1]
	
	** Percent
	sum    top5_`n'_heard_self
	matrix top5_`n'_heard_self=r(mean)
	matrix top5_`n'_heard_self_total=r(N)
	}
	
	
	** Put Excel
		
		*** Total
		putexcel ``group'col'`N'= matrix(top5_1_heard_self_total)
			local N=`N'+1
		
		*** Items and Percents
		foreach n in 1 2 3 4 5 {
			putexcel ``group'col'`N'=  "`top5_`n'_heard_self_item'"
				local N=`N'+1
			putexcel ``group'col'`N'=matrix(top5_`n'_heard_self)
				local N=`N'+1
		}



* 402 Have you heard about these family planning (heard_prompted)
foreach n in 1 2 3 4 5 {
			
	** Item
	local top5_`n'_heard_prompted_item=top5_`n'_heard_prompted_item[1]
	
	** Percent
	sum    top5_`n'_heard_prompted
	matrix top5_`n'_heard_prompted=r(mean)
	matrix top5_`n'_heard_prompted_total=r(N)
	}
	
	
	** Put Excel
		
		*** Total
		putexcel ``group'col'`N'= matrix(top5_1_heard_prompted_total)
			local N=`N'+1
		
		*** Items and Percents
		foreach n in 1 2 3 4 5 {
			putexcel ``group'col'`N'=  "`top5_`n'_heard_prompted_item'"
				local N=`N'+1
			putexcel ``group'col'`N'=matrix(top5_`n'_heard_prompted)
				local N=`N'+1
		}

* 401 and 402 LARCS, Short Acting and Traditional
foreach method in traditional shortacting larc {

	** Methods
	sum heard_combined_`method'
		matrix n_hcombined_`method'=r(N)
		matrix p_hcombined_`method'=r(mean)		
	}
	
	** Put Excel 
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_hcombined_shortacting)
			local N=`N'+1
			
		*** Percents
		foreach method in traditional shortacting larc {
			putexcel ``group'col'`N'=  matrix(p_hcombined_`method')
				local N=`N'+1		
			}

* 407 What family planning methods are you currently using (current)			
foreach n in 1 2 3 4 5 {
			
	** Item
	local top5_`n'_current_item=top5_`n'_current_item[1]
	
	** Percent
	sum    top5_`n'_current
	matrix top5_`n'_current=r(mean)
	matrix top5_`n'_current_total=r(N)
	}
	
	
	** Put Excel
		
		*** Total
		putexcel ``group'col'`N'= matrix(top5_1_current_total)
			local N=`N'+1
		
		*** Items and Percents
		foreach n in 1 2 3 4 5 {
			putexcel ``group'col'`N'=  "`top5_`n'_current_item'"
				local N=`N'+1
			putexcel ``group'col'`N'=matrix(top5_`n'_current)
				local N=`N'+1
			}
			
* 407 No Method, Traditional, Short Acting and LARC
foreach method in nomethod traditional shortacting larc {

	** Methods
	sum current_`method'
		matrix n_current_`method'=r(N)
		matrix p_current_`method'=r(mean)		
	}
	
	** Put Excel 
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_current_shortacting)
			local N=`N'+1
			
		*** Percents
		foreach method in nomethod traditional shortacting larc {
			putexcel ``group'col'`N'=  matrix(p_current_`method')
				local N=`N'+1		
			}
			
* 408 Reasons for not using fp (reason)
foreach n in 1 2 3 4 5 {
			
	** Item
	local top5_`n'_reason_item=top5_`n'_reason_item[1]
	
	** Percent
	sum    top5_`n'_reason
	matrix top5_`n'_reason=r(mean)
	matrix top5_`n'_reason_total=r(N)
	}
	
	
	** Put Excel
		
		*** Total
		putexcel ``group'col'`N'= matrix(top5_1_reason_total)
			local N=`N'+1
		
		*** Items and Percents
		foreach n in 1 2 3 4 5 {
			putexcel ``group'col'`N'=  "`top5_`n'_reason_item'"
				local N=`N'+1
			putexcel ``group'col'`N'=matrix(top5_`n'_reason)
				local N=`N'+1
			}
			

* 409 Previous family planning method (previous)			
foreach n in 1 2 3 4 5 {
			
	** Item
	local top5_`n'_previous_item=top5_`n'_previous_item[1]
	
	** Percent
	sum    top5_`n'_previous
	matrix top5_`n'_previous=r(mean)
	matrix top5_`n'_previous_total=r(N)
	}
	
	
	** Put Excel
		
		*** Total
		putexcel ``group'col'`N'= matrix(top5_1_previous_total)
			local N=`N'+1
		
		*** Items and Percents
		foreach n in 1 2 3 4 5 {
			putexcel ``group'col'`N'=  "`top5_`n'_previous_item'"
				local N=`N'+1
			putexcel ``group'col'`N'=matrix(top5_`n'_previous)
				local N=`N'+1
			}
			
* 409 No Method, Traditional, Short Acting and LARC	
foreach method in nomethod traditional shortacting larc {

	** Methods
	sum previous_`method'
		matrix n_previous_`method'=r(N)
		matrix p_previous_`method'=r(mean)		
	}
	
	** Put Excel 
	
		*** Totals
		putexcel ``group'col'`N'=  matrix(n_previous_shortacting)
			local N=`N'+1
			
		*** Percents
		foreach method in nomethod traditional shortacting larc {
			putexcel ``group'col'`N'=  matrix(p_previous_`method')
				local N=`N'+1		
			}
			
* 410 Start using method (lengthofuse)
foreach length in 6mo 6mo_12mo 13mo_24mo 25mo_48mo 49mo {
	
	** Length
	sum lengthofuse_`length'
		matrix n_lengthofuse_`length'=r(N)
		matrix p_lengthofuse_`length'=r(mean)
	}
	
	** Put Excel
	
		*** Total
		putexcel ``group'col'`N'=	matrix(n_lengthofuse_6mo)
			local N=`N'+1
		
		*** Percents
		foreach length in 6mo 6mo_12mo 13mo_24mo 25mo_48mo 49mo {
			putexcel ``group'col'`N'=	matrix(p_lengthofuse_`length')
				local N=`N'+1
			}
	}
*/
********************************************************
*** PART 2: Beans ***
********************************************************
/*
foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_2'")
local N `cellnum1'

* Create counts
gen n=1
sum n
matrix n=r(N)
local n=r(N) 
	
* Enter N
local celltext N=`n'
putexcel ``group'col'`N'="`celltext'"

local N=`N'+2

* All the beans (bean_)
foreach bean in unwantedbelly fphelpprevent fpimproveslives wombproblems ///
				shortintervals ppfp fpprovidestime ///
				fpreduceslabido womansduty ///
				suggestcondoms discussfp agreeonfp husbandangrycondom ///
				womanbirth manchild {
	
	** Amount
	foreach amount in 0_33 34_66 67_100  {
		sum bean_`bean'_`amount'
			matrix n_b_`bean'_`amount'=r(N)
			matrix p_b_`bean'_`amount'=r(mean)
		}
				
	** Put Excel
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_b_`bean'_0_33)
			local N=`N'+1
			
		*** Percents
		foreach amount in 0_33 34_66 67_100  {
			putexcel ``group'col'`N'=  matrix(p_b_`bean'_`amount')
				local N=`N'+1		
			}	
		}
	}
*/
********************************************************
*** PART 3: Neighborhood ***
********************************************************	
/*	
foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_3'")
local N `cellnum1'

* Create counts
gen n=1
sum n
matrix n=r(N)
local n=r(N) 
	
* Enter N
local celltext N=`n'
putexcel ``group'col'`N'="`celltext'"

local N=`N'+2
	
* Number questions (number_)
foreach number in usingfp fpadvice {	
	
	tab number_`number', matcell(table)
		matrix low=table[1,1]
		matrix medium=table[2,1]
		matrix high=table[3,1]
		matrix total=low[1,1]+medium[1,1]+high[1,1]
	
	matrix n_number_`number'=total[1,1]
	
	foreach level in low medium high {
		matrix p_number_`level'_`number'=`level'[1,1]/total[1,1]
		}
				
	** Put Excel 
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_number_`number')
			local N=`N'+1
			
		*** Percents
		foreach level in low medium high  {
			putexcel ``group'col'`N'=  matrix(p_number_`level'_`number')
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
local N `cellnum1'

* Create counts
gen n=1
sum n
matrix n=r(N)
local n=r(N) 
	
* Enter N
local celltext N=`n'
putexcel ``group'col'`N'="`celltext'"

local N=`N'+2
	
* 403 Have you seen a fp provider (visited_provider_12mo)
tab visited_provider_12mo, matcell(table)
	matrix no=table[1,1]
	matrix yes=table[2,1]
	matrix total=no[1,1]+yes[1,1]
			
	matrix n_visited_provider_12mo=total[1,1]
		
	foreach response in no yes {
		matrix p_visited_provider_12mo_`response'=`response'[1,1]/total[1,1]
		}
		
	** Put Excel
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_visited_provider_12mo)
			local N=`N'+1
			
		*** Percents
		foreach response in no yes  {
			putexcel ``group'col'`N'=  matrix(p_visited_provider_12mo_`response')
				local N=`N'+1		
			}
* 404-406 Family Planning Information (provider_informed)
foreach info in knowmethods problems whattodo {	
	tab provider_informed_`info', matcell(table)
		matrix no=table[1,1]
		matrix yes=table[2,1]
		matrix no_p=table[3,1]
		matrix total=no[1,1]+yes[1,1]+no_p[1,1]
			
		matrix n_info_`info'=total[1,1]
		
		foreach response in no yes no_p {
			matrix p_info_`response'_`info'=`response'[1,1]/total[1,1]
			}

	** Put Excel 
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_info_`info')
			local N=`N'+1
			
		*** Percents
		foreach response in no yes no_p  {
			putexcel ``group'col'`N'=  matrix(p_info_`response'_`info')
				local N=`N'+1		
			}	
		}

* 429 Family Planning Advice (provider_received_advice)	
tab provider_received_advice, matcell(table)
	matrix no=table[1,1]
	matrix yes=table[2,1]
	matrix total=no[1,1]+yes[1,1]
		
	matrix n_received_advice=total[1,1]
		
	foreach response in no yes {
		matrix p_received_advice_`response'=`response'[1,1]/total[1,1]
		}
	
	** Put Excel
	
		*** Total
		putexcel ``group'col'`N'=  matrix(n_received_advice)
			local N=`N'+1
			
		*** Percents
		foreach response in no yes  {
			putexcel ``group'col'`N'= matrix(p_received_advice_`response')
				local N=`N'+1		
			}
			
* All the beans (bean_)
foreach bean in providerneeds providerquestions providerinfo providerautonomy {
					
	foreach amount in 0_33 34_66 67_100  {
		sum bean_`bean'_`amount'
			matrix n_b_`bean'_`amount'=r(N)
			matrix p_b_`bean'_`amount'=r(mean)
		}
		
	** Put Excel
	
		*** Totals
		putexcel ``group'col'`N'=  matrix(n_b_`bean'_0_33)
			local N=`N'+1
			
		*** Percents
		foreach amount in 0_33 34_66 67_100  {
			putexcel ``group'col'`N'=  matrix(p_b_`bean'_`amount')
				local N=`N'+1		
			}
	}
}
*/
********************************************************
*** Norms ***
********************************************************	
		
use "1. Data/baseline_femaledata.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`norms'")
local N `cellnum2'

* Total 
sum N_total
local N_total=r(N)
local celltext N=`N_total'
putexcel `female_totalcol'`N'="`celltext'"

local N=`N'+2

* Partner Violence (gem_pv_index), Reproductive health (gem_rh_index), Sexual relatinoships (gem_sr_index), and Domestic chores and daily life (gem_dcdl_index)

foreach index in pv rh sr dcdl {
	
	** Counties
	foreach county in 22 23 24 {
		tab gem_`index'_index if county==`county', matcell(table)
		
			*** Generate Counts
			matrix n_low_`index'_`county'= table[1,1]
			matrix n_mid_`index'_`county'= table[2,1]
			matrix n_high_`index'_`county'=table[3,1]
			matrix n_total_`index'_`county'=n_low_`index'_`county'[1,1]+n_mid_`index'_`county'[1,1]+n_high_`index'_`county'[1,1]
			
			*** Generate Percents
			foreach level in low mid high {
				matrix p_`level'_`index'_`county'=n_`level'_`index'_`county'[1,1]/n_total_`index'_`county'[1,1]
				}
		}
		
	** Total
	tab gem_`index'_index, matcell(table)
		
		*** Generate Counts
		matrix n_low_`index'_total= table[1,1]
		matrix n_mid_`index'_total= table[2,1]
		matrix n_high_`index'_total=table[3,1]
		matrix n_total_`index'_total=n_low_`index'_total[1,1]+n_mid_`index'_total[1,1]+n_high_`index'_total[1,1]
		
		** Generate Percents
		foreach level in low mid high {
			matrix p_`level'_`index'_total=n_`level'_`index'_total[1,1]/n_total_`index'_total[1,1]
			}
	
	** PutExcel 
	foreach level in low mid high {
		putexcel `female_bongcol'`N'=   matrix(p_`level'_`index'_23)
		putexcel `female_bombicol'`N'=  matrix(p_`level'_`index'_22)
		putexcel `female_controlcol'`N'=matrix(p_`level'_`index'_24)
		putexcel `female_totalcol'`N'=  matrix(p_`level'_`index'_total)
			local N=`N'+1
		}
	
	local N=`N'+1	
	}
	
* Self is Main Decision Maker (decision_)
foreach decision in numchildren contraception selfill {
	
	** Counties
	foreach county in 22 23 24 {
		tab decision_`decision' if county==`county', matcell(table)
		
			*** Generate Counts
			matrix n_self_`decision'_`county'= table[1,1]
			matrix n_partner_`decision'_`county'= table[2,1]
			matrix n_other_`decision'_`county'=table[3,1]
			matrix n_total_`decision'_`county'=n_self_`decision'_`county'[1,1]+n_partner_`decision'_`county'[1,1]+n_other_`decision'_`county'[1,1]
			
			*** Generate Percents
			matrix p_self_`decision'_`county'=n_self_`decision'_`county'[1,1]/n_total_`decision'_`county'[1,1]
				
		}
		
	** Total
	tab decision_`decision', matcell(table)
		
		*** Generate Counts
		matrix n_self_`decision'_total= table[1,1]
		matrix n_partner_`decision'_total= table[2,1]
		matrix n_other_`decision'_total=table[3,1]
		matrix n_total_`decision'_total=n_self_`decision'_total[1,1]+n_partner_`decision'_total[1,1]+n_other_`decision'_total[1,1]
		
		** Generate Percents
		matrix p_self_`decision'_total=n_self_`decision'_total[1,1]/n_total_`decision'_total[1,1]
				
	** PutExcel 
	putexcel `female_bongcol'`N'=   matrix(p_self_`decision'_23)
	putexcel `female_bombicol'`N'=  matrix(p_self_`decision'_22)
	putexcel `female_controlcol'`N'=matrix(p_self_`decision'_24)
	putexcel `female_totalcol'`N'=  matrix(p_self_`decision'_total)
		local N=`N'+1
	}
	
local N=`N'+1
		
* Home Pregnancy Environment (home_pregnancy)
foreach support in supportive stressful {
	
	** Counties
	foreach county in 22 23 24 {
		tab home_pregnancy_`support' if county==`county', matcell(table)
		
			*** Generate Counts
			matrix n_low_`support'_`county'= table[1,1]
			matrix n_mid_`support'_`county'= table[2,1]
			matrix n_high_`support'_`county'=table[3,1]
			matrix n_total_`support'_`county'=n_low_`support'_`county'[1,1]+n_mid_`support'_`county'[1,1]+n_high_`support'_`county'[1,1]
			
			*** Generate Percents
			foreach level in low mid high {
				matrix p_`level'_`support'_`county'=n_`level'_`support'_`county'[1,1]/n_total_`support'_`county'[1,1]
				}
		}
		
	** Total
	tab home_pregnancy_`support', matcell(table)
		
		*** Generate Counts
		matrix n_low_`support'_total= table[1,1]
		matrix n_mid_`support'_total= table[2,1]
		matrix n_high_`support'_total=table[3,1]
		matrix n_total_`support'_total=n_low_`support'_total[1,1]+n_mid_`support'_total[1,1]+n_high_`support'_total[1,1]
		
		** Generate Percents
		foreach level in low mid high {
			matrix p_`level'_`support'_total=n_`level'_`support'_total[1,1]/n_total_`support'_total[1,1]
			}
	
	** PutExcel 
	foreach level in low mid high {
		putexcel `female_bongcol'`N'=   matrix(p_`level'_`support'_23)
		putexcel `female_bombicol'`N'=  matrix(p_`level'_`support'_22)
		putexcel `female_controlcol'`N'=matrix(p_`level'_`support'_24)
		putexcel `female_totalcol'`N'=  matrix(p_`level'_`support'_total)
			local N=`N'+1
		}
	
	local N=`N'+1	
	}
		
* Exposed to FP messages in past six months (fp_messaging)
		
	** Counties
	foreach county in 22 23 24 {
		tab fp_messaging if county==`county', matcell(table)
		
			*** Generate Counts
			matrix n_no_`county'= table[1,1]
			matrix n_yes_`county'= table[2,1]
			matrix n_total_`county'=n_yes_`county'[1,1]+n_no_`county'[1,1]
			
			*** Generate Percents
			matrix p_yes_`county'=n_yes_`county'[1,1]/n_total_`county'[1,1]
				
		}
		
	** Total
	tab fp_messaging, matcell(table)
		
			*** Generate Counts
			matrix n_no_total= table[1,1]
			matrix n_yes_total= table[2,1]
			matrix n_total_total=n_yes_total[1,1]+n_no_total[1,1]
			
			*** Generate Percents
			matrix p_yes_total=n_yes_total[1,1]/n_total_total[1,1]
				
	** PutExcel 
	putexcel `female_bongcol'`N'=   matrix(p_yes_23)
	putexcel `female_bombicol'`N'=  matrix(p_yes_22)
	putexcel `female_controlcol'`N'=matrix(p_yes_24)
	putexcel `female_totalcol'`N'=  matrix(p_yes_total)
	
	
	
	
