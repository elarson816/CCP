*** Tables
cd $directory

* Set macros for cells (General Tables)
local cellnum1 $cellnum1
local female_ums_bongcol1 $female_ums_bongcol1
local female_ums_bombicol1 $female_ums_bombicol1
local female_ums_totalcol1 $female_ums_totalcol1
local female_ums_controlol1 $female_ums_controlol1
local female_iu_bongcol1 $female_iu_bongcol1
local female_iu_bombicol1 $female_iu_bombicol1
local female_iu_totalcol1 $female_iu_totalcol1
local female_iu_controlol1 $female_iu_controlol1
local male_bongcol1 $male_bongcol1
local male_bombicol1 $male_bombicol1
local male_totalcol1 $male_totalcol1
local male_controlol1 $male_controlol1

* Set macros for cells (Demographic Table)
local cellnum2 $cellnum2
local female_bongcol $female_bongcol
local female_bombicol $female_bombicol 
local female_controlcol $female_controlcol 
local female_totalcol $female_totalcol

* Set macros for sheets
local part_1 $part_1
local part_2 $part_2
local part_3 $part_3
local part_4 $part_4
local demographic $demographic

********************************************************
*** PART 1: CONTRACEPTIVE USE ***
********************************************************

foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_1'")
local N `cellnum1'

* Create counts
	
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
local N `cellnum1'

* Create counts

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
local N `cellnum1'

* Create counts
	
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

	
********************************************************
*** PART 4: Provider ***
********************************************************	
	
foreach group in female_ums female_iu male {
	
use "1. Data/baseline_`group'data_merged.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`part_4'")
local N `cellnum1'

* Create counts
	
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
*/
********************************************************
*** DEMOGRAPHIC ***
********************************************************	
		
use "1. Data/baseline_femaledata.dta", clear

* Set putexcel
putexcel set "$putexcel_set", modify sheet("`demographic'")
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
	
	
	
	
