** Bivariate Analysis (Females Only)
cd $directory

* Set macros for cells (Bivariate Analysis)
local cellnum3 $cellnum3 
local fp_col   $fp_col 
local chi2_col $chi2_col 

* Set macros for sheets
local bi_part_1 $bi_part_1
local bi_part_2 $bi_part_2
local bi_part_3 $bi_part_3

********************************************************
*** Part 1 ***
********************************************************
use "1. Data/baseline_femaledata.dta", clear

* Set Put Excel
putexcel set "$putexcel_set", modify sheet("`bi_part_1'")
local N `cellnum3'

* Create counts
sum cp if cp==1
matrix n_cp=r(N)
local n_cp=r(N)

* Enter Ns
local celltext (N=`n_cp')
putexcel `chi2_col'`N'="`celltext'"

local N=`N'+4

* Knows methods
foreach method in trad shortacting larc gt_median {
	tab heard_combined_`method' cp, co chi matcell(table)
	
		** Counts
		matrix n_`method'_no=table[1,2]
		matrix n_`method'_yes=table[2,2]
		matrix n_`method'_total=n_`method'_no[1,1]+n_`method'_yes[1,1]
		
		** Percents
		matrix p_`method'_no=n_`method'_no[1,1]/n_`method'_total[1,1]
		matrix p_`method'_yes=n_`method'_yes[1,1]/n_`method'_total[1,1]
		
		** Chi2
		matrix chi2_`method'=r(p)
		
		** PutExcel
		putexcel `fp_col'`N'=matrix(p_`method'_no)
			local N=`N'+1
		putexcel `fp_col'`N'=matrix(p_`method'_yes)
		putexcel `chi2_col'`N'=matrix(chi2_`method')
			local N=`N'+2

	}

* Used FP services in last 12 months
tab visited_provider_12mo cp, co chi matcell(table)

		** Counts
		matrix n_visited_no=table[1,2]
		matrix n_visited_yes=table[2,2]
		matrix n_visited_total=n_visited_no[1,1]+n_visited_yes[1,1]
		
		** Percents
		matrix p_visited_no=n_visited_no[1,1]/n_visited_total[1,1]
		matrix p_visited_yes=n_visited_yes[1,1]/n_visited_total[1,1]
		
		** Chi2
		matrix chi2_visited=r(p)
		
		** PutExcel
		putexcel `fp_col'`N'=matrix(p_visited_no)
			local N=`N'+1
		putexcel `fp_col'`N'=matrix(p_visited_yes)
		putexcel `chi2_col'`N'=matrix(chi2_visited)
			local N=`N'+2

* Patient Information Index
tab PII cp, co chi matcell(table)

		** Counts
		matrix n_PII_no=table[1,2]
		matrix n_PII_yes=table[2,2]
		matrix n_PII_total=n_PII_no[1,1]+n_PII_yes[1,1]
		
		** Percents
		matrix p_PII_no=n_PII_no[1,1]/n_PII_total[1,1]
		matrix p_PII_yes=n_PII_yes[1,1]/n_PII_total[1,1]
		
		** Chi2
		matrix chi2_PII=r(p)
		
		** PutExcel
		putexcel `fp_col'`N'=matrix(p_PII_no)
			local N=`N'+1
		putexcel `fp_col'`N'=matrix(p_PII_yes)
		putexcel `chi2_col'`N'=matrix(chi2_PII)
			local N=`N'+2
			
* Used FP services and quality received

********************************************************
*** Part 2 ***
********************************************************
use "1. Data/baseline_femaledata.dta", clear

* Set Put Excel
putexcel set "$putexcel_set", modify sheet("`bi_part_2'")
local N `cellnum3'

* Create counts
sum cp if cp==1
matrix n_cp=r(N)
local n_cp=r(N)

* Enter Ns
local celltext (N=`n_cp')
putexcel `chi2_col'`N'="`celltext'"

local N=`N'+4

* GEM Indexes
foreach index in pv rh sr dcdl {
	tab gem_`index'_index cp, co chi matcell(table)
		
		*** Counts
		matrix n_low_`index'= table[1,2]
		matrix n_mid_`index'= table[2,2]
		matrix n_high_`index'=table[3,2]
		matrix n_total_`index'=n_low_`index'[1,1]+n_mid_`index'[1,1]+n_high_`index'[1,1]
		
		** Percents
		foreach level in low mid high {
			matrix p_`level'_`index'=n_`level'_`index'[1,1]/n_total_`index'[1,1]
			}
			
		** Chi2
		matrix chi2_`index'=r(p)
	
	** PutExcel 
	foreach level in low mid high {
		putexcel `fp_col'`N'= matrix(p_`level'_`index')
			local N=`N'+1
		}
			local N=`N'-1
		putexcel `chi2_col'`N'=matrix(chi2_`index')
	
	local N=`N'+2	
	}
	
* Decision Making
foreach decision in numchildren contraception selfill {
	tab decision_`decision'_yn cp, co chi matcell(table)
		
		*** Generate Counts
		matrix n_no_`decision'= table[1,2]
		matrix n_yes_`decision'= table[2,2]
		matrix n_total_`decision'=n_yes_`decision'[1,1]+n_no_`decision'[1,1]
		
		** Generate Percents
		matrix p_yes_`decision'=n_yes_`decision'[1,1]/n_total_`decision'[1,1]
		matrix p_no_`decision'=n_no_`decision'[1,1]/n_total_`decision'[1,1]
		
		** Chi2
		matrix chi2_`decision'=r(p)
				
	** PutExcel 
	putexcel `fp_col'`N'=   matrix(p_no_`decision')
		local N=`N'+1
	putexcel `fp_col'`N'=  matrix(p_yes_`decision')
	putexcel `chi2_col'`N'=matrix(chi2_`decision')
		local N=`N'+2
	}
	
* Decision Making - All Three
tab count_decision cp, co chi matcell(table)

	** Counts
	matrix n_count_0=table[1,2]
	matrix n_count_1=table[2,2]
	matrix n_count_2=table[3,2]
	matrix n_count_3=table[4,2]
	matrix n_count_total=n_count_0[1,1]+n_count_1[1,1]+n_count_2[1,1]+n_count_3[1,1]
	
	** Percents
	foreach count in 0 1 2 3 {
		matrix p_count_`count'=n_count_`count'[1,1]/n_count_total[1,1]
		}
	
	** Chi2
	matrix chi2_count=r(p)
	
	** PutExcel
	foreach count in 0 1 2 3 {
		putexcel `fp_col'`N'=matrix(p_count_`count')
			local N=`N'+1
		}
			local N=`N'-1
		putexcel `chi2_col'`N'=matrix(chi2_count)
		
	local N=`N'+2
	
* Exposed for FP Messaging
tab fp_messaging cp, co chi matcell(table)
		
	*** Counts
	matrix n_no= table[1,2]
	matrix n_yes= table[2,2]
	matrix n_total=n_yes[1,1]+n_no[1,1]
	
	** Generate Percents
	matrix p_yes=n_yes[1,1]/n_total[1,1]
	matrix p_no=n_no[1,1]/n_total[1,1]
	
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	putexcel `fp_col'`N'=matrix(p_no)
		local N=`N'+1
	putexcel `fp_col'`N'=matrix(p_yes)
	putexcel `chi2_col'`N'=matrix(chi2)
		local N=`N'+2
		
* Media
foreach item in radio tv cellphone {
	tab media_`item'_often cp, co chi matcell(table)
	
	*** Counts
	matrix n_no_`item'= table[1,2]
	matrix n_yes_`item'= table[2,2]
	matrix n_total_`item'=n_yes_`item'[1,1]+n_no_`item'[1,1]
	
	** Generate Percents
	matrix p_yes_`item'=n_yes_`item'[1,1]/n_total_`item'[1,1]
	matrix p_no_`item'=n_no_`item'[1,1]/n_total_`item'[1,1]
	
	** Chi2
	matrix chi2_`item'=r(p)
	
	** PutExcel
	putexcel `fp_col'`N'=matrix(p_no_`item')
		local N=`N'+1
	putexcel `fp_col'`N'=matrix(p_yes_`item')
	putexcel `chi2_col'`N'=matrix(chi2_`item')
		local N=`N'+2

	}
*/
********************************************************
*** Part 3 ***
********************************************************
use "1. Data/baseline_femaledata.dta", clear

* Set Put Excel
putexcel set "$putexcel_set", modify sheet("`bi_part_3'")
local N `cellnum3'

* Create counts
sum cp if cp==1
matrix n_cp=r(N)
local n_cp=r(N)

* Enter Ns
local celltext (N=`n_cp')
putexcel `chi2_col'`N'="`celltext'"

local N=`N'+4

* Age
tab age_cat cp, co chi matcell(table)
		
	*** Counts
	matrix n_14_16= table[1,2]
	matrix n_17_19= table[2,2]
	matrix n_total=n_14_16[1,1]+n_17_19[1,1]
	
	** Generate Percents
	matrix n_17_19=n_17_19[1,1]/n_total[1,1]
	matrix p_14_16=n_14_16[1,1]/n_total[1,1]
	
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	putexcel `fp_col'`N'=matrix(p_14_16)
		local N=`N'+1
	putexcel `fp_col'`N'=matrix(n_17_19)
	putexcel `chi2_col'`N'=matrix(chi2)
		local N=`N'+2
		
* Education
tab education cp, co chi matcell(table)
		
	*** Counts
	matrix n_primary= table[1,2]
	matrix n_secondary= table[2,2]
	matrix n_total=n_primary[1,1]+n_secondary[1,1]
	
	** Generate Percents
	foreach level in primary secondary {
		matrix p_`level'=n_`level'[1,1]/n_total[1,1]
		}
		
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	foreach level in primary secondary {	
		putexcel `fp_col'`N'=matrix(p_`level')
			local N=`N'+1
		}
			local N=`N'-1
		putexcel `chi2_col'`N'=matrix(chi2)
			local N=`N'+2
	
* Area of residence
tab ur cp, co chi matcell(table)
		
	*** Counts
	matrix n_urban= table[1,2]
	matrix n_rural= table[2,2]
	matrix n_total=n_urban[1,1]+n_rural[1,1]
	
	** Generate Percents
	matrix p_rural=n_rural[1,1]/n_total[1,1]
	matrix p_urban=n_urban[1,1]/n_total[1,1]
	
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	putexcel `fp_col'`N'=matrix(p_rural)
		local N=`N'+1
	putexcel `fp_col'`N'=matrix(p_urban)
	putexcel `chi2_col'`N'=matrix(chi2)
		local N=`N'+2		

* Religion
tab religion cp, co chi matcell(table)
		
	*** Counts
	matrix n_christian= table[1,2]
	matrix n_muslim= table[2,2]
	matrix n_other= table[3,2]
	matrix n_total=n_christian[1,1]+n_muslim[1,1]+n_other[1,1]
	
	** Generate Percents
	foreach level in christian muslim other {
		matrix p_`level'=n_`level'[1,1]/n_total[1,1]
		}
		
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	foreach level in christian muslim other {	
		putexcel `fp_col'`N'=matrix(p_`level')
			local N=`N'+1
		}
			local N=`N'-1
		putexcel `chi2_col'`N'=matrix(chi2)
			local N=`N'+2

* Marital Status
tab inunion cp, co chi matcell(table)
		
	*** Counts
	matrix n_inunion= table[1,2]
	matrix n_not= table[2,2]
	matrix n_total=n_inunion[1,1]+n_not[1,1]
	
	** Generate Percents
	matrix n_not=n_not[1,1]/n_total[1,1]
	matrix n_inunion=n_inunion[1,1]/n_total[1,1]
	
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	putexcel `fp_col'`N'=matrix(n_inunion)
		local N=`N'+1
	putexcel `fp_col'`N'=matrix(n_not)
	putexcel `chi2_col'`N'=matrix(chi2)
		local N=`N'+2	
		
* Given Birth
tab given_birth cp, co chi matcell(table)
		
	*** Counts
	matrix n_no= table[1,2]
	matrix n_yes= table[2,2]
	matrix n_total=n_yes[1,1]+n_no[1,1]
	
	** Generate Percents
	matrix p_yes=n_yes[1,1]/n_total[1,1]
	matrix p_no=n_no[1,1]/n_total[1,1]
	
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	putexcel `fp_col'`N'=matrix(p_no)
		local N=`N'+1
	putexcel `fp_col'`N'=matrix(p_yes)
	putexcel `chi2_col'`N'=matrix(chi2)
		local N=`N'+2

* Vulnerability and Standard of Living Indexes
foreach index in vbl sl {
	tab `index'_index cp, co chi matcell(table)
	
		*** Counts
		matrix n_low_`index'= table[1,2]
		matrix n_mid_`index'= table[2,2]
		matrix n_high_`index'=table[3,2]
		matrix n_total_`index'=n_low_`index'[1,1]+n_mid_`index'[1,1]+n_high_`index'[1,1]
		
		** Percents
		foreach level in low mid high {
			matrix p_`level'_`index'=n_`level'_`index'[1,1]/n_total_`index'[1,1]
			}
			
		** Chi2
		matrix chi2_`index'=r(p)
	
	** PutExcel 
	foreach level in low mid high {
		putexcel `fp_col'`N'= matrix(p_`level'_`index')
			local N=`N'+1
		}
			local N=`N'-1
		putexcel `chi2_col'`N'=matrix(chi2_`index')
	
	local N=`N'+2	
	}
	
* County of residence
tab county_id cp, co chi matcell(table)

	*** Counts
	matrix n_bomi= table[1,2]
	matrix n_bong= table[2,2]
	matrix n_gbarpolu= table[3,2]
	matrix n_total=n_bomi[1,1]+n_bong[1,1]+n_gbarpolu[1,1]
	
	** Generate Percents
	foreach level in bomi bong gbarpolu {
		matrix p_`level'=n_`level'[1,1]/n_total[1,1]
		}
		
	** Chi2
	matrix chi2=r(p)
	
	** PutExcel
	foreach level in bomi bong gbarpolu {	
		putexcel `fp_col'`N'=matrix(p_`level')
			local N=`N'+1
		}
			local N=`N'-1
		putexcel `chi2_col'`N'=matrix(chi2)




