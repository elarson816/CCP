** Maternal Health Tables
cd "$directory"

* Set macros for cells (General Tables)
local R $R
local rowlabel $rowlabel
local women_collabel $women_collabel
local women_intervention_col1 $women_intervention_col1
local women_intervention_col2 $women_intervention_col2
local women_intervention_col3 $women_intervention_col3
local women_control_col1 $women_control_col1
local men_collabel $men_collabel
local men_intervention_col1 $men_intervention_col1
local men_intervention_col2 $men_intervention_col2
local men_intervention_col3 $men_intervention_col3
local men_control_col1 $men_control_col1


********************************************************
*** Table 5.1: TIMING OF FIRST PRENATAL VISIT ***
********************************************************

putexcel set "$putexcel_set", modify sheet("Table5.1")

* Titles
putexcel `women_collabel'`R'="Women's (20-49 years)"
putexcel `men_collabel'`R'  ="Men/Partners (20-55 years)"
	local R=`R'+1
putexcel `women_intervention_col1'`R'="Intervention Counties"
putexcel `women_control_col1'`R'     ="Control"

putexcel `men_intervention_col1'`R'="Intervention Counties"
putexcel `men_control_col1'`R'     ="Control"
	local R=`R'+1
putexcel `rowlabel'`R'="Prenatal Care"

* Counts
sum one if county_id==23 & female==1 & mch_217_cat!=.
	local n_Bong_female=r(N)
sum one if county_id==22 & female==1 & mch_217_cat!=.
	local n_Bomi_female=r(N)
sum one if control==0 & female==1 & mch_217_cat!=.
	local n_Total_female=r(N)
sum one if county_id==24 & female==1 & mch_217_cat!=.
	local n_Gbarpolu_female=r(N)
sum one if county_id==23 & female==0 & mch_217_cat!=.
	local n_Bong_male=r(N)
sum one if county_id==22 & female==0 & mch_217_cat!=.
	local n_Bomi_male=r(N)
sum one if control==0 & female==0 & mch_217_cat!=.
	local n_Total_male=r(N)
sum one if county_id==24 & female==0 & mch_217_cat!=.
	local n_Gbarpolu_male=r(N)
	
	** Tiles
	local celltext Bong N=`n_Bong_female' %
		putexcel `women_intervention_col1'`R'="`celltext'"
	local celltext Bomi N=`n_Bomi_female' %
		putexcel `women_intervention_col2'`R'="`celltext'"
	local celltext Total N=`n_Total_female' %
		putexcel `women_intervention_col3'`R'="`celltext'"
	local celltext Gbarpolu N=`n_Gbarpolu_female' %
		putexcel `women_control_col1'`R'="`celltext'"
	
	local celltext Bong N=`n_Bong_male' %
		putexcel `men_intervention_col1'`R'="`celltext'"
	local celltext Bomi N=`n_Bomi_male' %
		putexcel `men_intervention_col2'`R'="`celltext'"
	local celltext Total N=`n_Total_male' %
		putexcel `men_intervention_col3'`R'="`celltext'"
	local celltext Gbarpolu N=`n_Gbarpolu_male' %
		putexcel `men_control_col1'`R'="`celltext'"	
		
local R=`R'+1

* Title
putexcel `rowlabel'`R'="First prenatal visit"
	local R=`R'+1
	
* Percents
tab mch_217_cat county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_3months_Bomi=table[1,1]
	matrix female_3months_Bong=table[1,2]
	matrix female_3months_Total=female_3months_Bomi[1,1]+female_3months_Bong[1,1]
	matrix female_4_5months_Bomi=table[2,1]
	matrix female_4_5months_Bong=table[2,2]
	matrix female_4_5months_Total=female_4_5months_Bomi[1,1]+female_4_5months_Bong[1,1]
	matrix female_6_plusmonths_Bomi=table[3,1]
	matrix female_6_plusmonths_Bong=table[3,2]
	matrix female_6_plusmonths_Total=female_6_plusmonths_Bomi[1,1]+female_6_plusmonths_Bong[1,1]
	
	matrix female_Bomi_Total= female_3months_Bomi[1,1]+ female_4_5months_Bomi[1,1]+ female_6_plusmonths_Bomi[1,1]
	matrix female_Bong_Total= female_3months_Bong[1,1]+ female_4_5months_Bong[1,1]+ female_6_plusmonths_Bong[1,1]
	matrix female_Total_Total=female_3months_Total[1,1]+female_4_5months_Total[1,1]+female_6_plusmonths_Total[1,1]
	
	matrix      p_3months_Bomi_female=     female_3months_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix    p_4_5months_Bomi_female=   female_4_5months_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_6_plusmonths_Bomi_female=female_6_plusmonths_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix      p_3months_Bong_female=     female_3months_Bong[1,1]/female_Bong_Total[1,1]
	matrix    p_4_5months_Bong_female=   female_4_5months_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_6_plusmonths_Bong_female=female_6_plusmonths_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix      p_3months_Total_female=     female_3months_Total[1,1]/female_Total_Total[1,1]
	matrix    p_4_5months_Total_female=   female_4_5months_Total[1,1]/female_Total_Total[1,1]
	matrix p_6_plusmonths_Total_female=female_6_plusmonths_Total[1,1]/female_Total_Total[1,1]
	
tab mch_217_cat if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_3months_Gbarpolu=table[1,1]
	matrix female_4_5months_Gbarpolu=table[2,1]
	matrix female_6_plusmonths_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_3months_Gbarpolu[1,1]+female_4_5months_Gbarpolu[1,1]+female_6_plusmonths_Gbarpolu[1,1]
	
	matrix      p_3months_Gbarpolu_female=     female_3months_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix    p_4_5months_Gbarpolu_female=   female_4_5months_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_6_plusmonths_Gbarpolu_female=female_6_plusmonths_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
tab mch_217_cat county_id if female==0 & control==0 & haschild2yrs==1, matcell(table)
	matrix male_3months_Bomi=table[1,1]
	matrix male_3months_Bong=table[1,2]
	matrix male_3months_Total=male_3months_Bomi[1,1]+male_3months_Bong[1,1]
	matrix male_4_5months_Bomi=table[2,1]
	matrix male_4_5months_Bong=table[2,2]
	matrix male_4_5months_Total=male_4_5months_Bomi[1,1]+male_4_5months_Bong[1,1]
	matrix male_6_plusmonths_Bomi=table[3,1]
	matrix male_6_plusmonths_Bong=table[3,2]
	matrix male_6_plusmonths_Total=male_6_plusmonths_Bomi[1,1]+male_6_plusmonths_Bong[1,1]
	
	matrix male_Bomi_Total= male_3months_Bomi[1,1]+ male_4_5months_Bomi[1,1]+ male_6_plusmonths_Bomi[1,1]
	matrix male_Bong_Total= male_3months_Bong[1,1]+ male_4_5months_Bong[1,1]+ male_6_plusmonths_Bong[1,1]
	matrix male_Total_Total=male_3months_Total[1,1]+male_4_5months_Total[1,1]+male_6_plusmonths_Total[1,1]
	
	matrix      p_3months_Bomi_male=     male_3months_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix    p_4_5months_Bomi_male=   male_4_5months_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_6_plusmonths_Bomi_male=male_6_plusmonths_Bomi[1,1]/male_Bomi_Total[1,1]
	
	matrix      p_3months_Bong_male=     male_3months_Bong[1,1]/male_Bong_Total[1,1]
	matrix    p_4_5months_Bong_male=   male_4_5months_Bong[1,1]/male_Bong_Total[1,1]
	matrix p_6_plusmonths_Bong_male=male_6_plusmonths_Bong[1,1]/male_Bong_Total[1,1]
	
	matrix      p_3months_Total_male=     male_3months_Total[1,1]/male_Total_Total[1,1]
	matrix    p_4_5months_Total_male=   male_4_5months_Total[1,1]/male_Total_Total[1,1]
	matrix p_6_plusmonths_Total_male=male_6_plusmonths_Total[1,1]/male_Total_Total[1,1]
	
tab mch_217_cat if female==0 & control==1 & haschild2yrs==1, matcell(table)
	matrix male_3months_Gbarpolu=table[1,1]
	matrix male_4_5months_Gbarpolu=table[2,1]
	matrix male_6_plusmonths_Gbarpolu=table[3,1]
	matrix male_Gbarpolu_Total=male_3months_Gbarpolu[1,1]+male_4_5months_Gbarpolu[1,1]+male_6_plusmonths_Gbarpolu[1,1]
	
	matrix      p_3months_Gbarpolu_male=     male_3months_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix    p_4_5months_Gbarpolu_male=   male_4_5months_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix p_6_plusmonths_Gbarpolu_male=male_6_plusmonths_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	
* PutExcel

	** ≤3 months
	putexcel `rowlabel'`R'="≤3 months"
	putexcel `women_intervention_col1'`R'=  matrix(p_3months_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_3months_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_3months_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_3months_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_3months_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_3months_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_3months_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_3months_Gbarpolu_male)

	local R=`R'+1
	
	** 4-5 months
	putexcel `rowlabel'`R'="4-5 months"
	putexcel `women_intervention_col1'`R'=  matrix(p_4_5months_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_4_5months_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_4_5months_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_4_5months_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_4_5months_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_4_5months_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_4_5months_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_4_5months_Gbarpolu_male)

	local R=`R'+1
	
	** ≥6 months
	putexcel `rowlabel'`R'="≥6 months"
	putexcel `women_intervention_col1'`R'=  matrix(p_6_plusmonths_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_6_plusmonths_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_6_plusmonths_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_6_plusmonths_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_6_plusmonths_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_6_plusmonths_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_6_plusmonths_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_6_plusmonths_Gbarpolu_male)


*********************************************************************
*** Table 5.2: FACTORS ASSOCIATED WITH USE OF EARLY PRENATAL CARE ***
*********************************************************************

putexcel set "$putexcel_set", modify sheet("Table5.2")
local R $R

* Titles
putexcel `women_collabel'`R'="Women's (20-49 years)"
	local R=`R'+1
	
	
** At Least One Birth (Total)
*******************************************

*Title
putexcel `rowlabel'`R'="At Least One Birth"


* Counts
sum one if county_id==23 & female==1 & haschild2yrs==1
	local n_Bong_female=r(N)
sum one if county_id==22 & female==1 & haschild2yrs==1
	local n_Bomi_female=r(N)
sum one if control==0 & female==1 & haschild2yrs==1
	local n_Total_female=r(N)
sum one if county_id==24 & female==1 & haschild2yrs==1
	local n_Gbarpolu_female=r(N)
	
	** Tiles
	local celltext Bong N=`n_Bong_female' %
		putexcel `women_intervention_col1'`R'="`celltext'"
	local celltext Bomi N=`n_Bomi_female' %
		putexcel `women_intervention_col2'`R'="`celltext'"
	local celltext Total N=`n_Total_female' %
		putexcel `women_intervention_col3'`R'="`celltext'"
	local celltext Gbarpolu N=`n_Gbarpolu_female' %
		putexcel `women_control_col1'`R'="`celltext'"
		
local R=`R'+1

* Title
putexcel `rowlabel'`R'="At Least One ANC (Total)"

* Percents
tab mch_217_cat county_id if female==1 & haschild2yrs==1, mi matcell(table)
	matrix female_None=table[4,1]+table[4,2]+table[4,3]
	matrix female_Yes=table[1,1]+table[1,2]+table[3,2]+table[2,1]+table[2,2]+table[2,2]+table[3,1]+table[3,2]+table[3,2]
	matrix female_Total=female_None[1,1]+female_Yes[1,1]
	
	matrix p_female_Yes=female_Yes[1,1]/female_Total[1,1]
	
* PutExcel
putexcel `women_intervention_col1'`R'=  matrix(p_female_Yes)

local R=`R'+1


** Early ANC (Total)
*******************************************

* Percents
tab mch_217_cat county_id if female==1 & haschild2yrs==1, matcell(table)
	matrix female_3month=table[1,1]+table[1,2]+table[1,3]
	matrix female_Total=female_3month[1,1]+table[2,1]+table[2,2]+table[2,2]+table[3,1]+table[3,2]+table[3,2]
	
	matrix p_female_3month=female_3month[1,1]/female_Total[1,1]
	
* Title
putexcel `rowlabel'`R'="Early ANC (Total)"
	
* PutExcel
putexcel `women_intervention_col1'`R'=  matrix(p_female_3month)


** Not Early ANC
*******************************************

* Percents
tab mch_217_cat county_id if female==1 & haschild2yrs==1, matcell(table)
	matrix female_Not_3mo=table[2,1]+table[2,2]+table[2,2]+table[3,1]+table[3,2]+table[3,2]
	matrix female_Total=female_Not_3mo[1,1]+table[1,1]+table[1,2]+table[1,2]
	
	matrix p_female_Not_3mo=female_Not_3mo[1,1]/female_Total[1,1]
	
* Title
putexcel `rowlabel'`R'="Early ANC (Total)"
	
* PutExcel
putexcel `women_intervention_col1'`R'=  matrix(p_female_Not_3mo)


** Reasons For No ANC
*******************************************



** Perceived Self-Efficacy to Use ANC
*******************************************

* Title
putexcel `rowlabel'`R'="Confidence to seek early ANC"
	local R=`R'+1

* Percents
tab mch_254_cat county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_low_Bomi=table[1,1]
	matrix female_low_Bong=table[1,2]
	matrix female_low_Total=female_low_Bomi[1,1]+female_low_Bong[1,1]
	matrix female_medium_Bomi=table[2,1]
	matrix female_medium_Bong=table[2,2]
	matrix female_medium_Total=female_medium_Bomi[1,1]+female_medium_Bong[1,1]
	matrix female_high_Bomi=table[3,1]
	matrix female_high_Bong=table[3,2]
	matrix female_high_Total=female_high_Bomi[1,1]+female_high_Bong[1,1]
	
	matrix female_Bomi_Total= female_low_Bomi[1,1]+ female_medium_Bomi[1,1]+ female_high_Bomi[1,1]
	matrix female_Bong_Total= female_low_Bong[1,1]+ female_medium_Bong[1,1]+ female_high_Bong[1,1]
	matrix female_Total_Total=female_low_Total[1,1]+female_medium_Total[1,1]+female_high_Total[1,1]
	
	matrix    p_low_Bomi_female=   female_low_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_medium_Bomi_female=female_medium_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix   p_high_Bomi_female=  female_high_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix    p_low_Bong_female=   female_low_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_medium_Bong_female=female_medium_Bong[1,1]/female_Bong_Total[1,1]
	matrix   p_high_Bong_female=  female_high_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix    p_low_Total_female=   female_low_Total[1,1]/female_Total_Total[1,1]
	matrix p_medium_Total_female=female_medium_Total[1,1]/female_Total_Total[1,1]
	matrix   p_high_Total_female=  female_high_Total[1,1]/female_Total_Total[1,1]
	
tab mch_254_cat if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** ≤3 months
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** 4-5 months
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** ≥6 months
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1

** Perceived Self-Efficacy for Multiple ANC
*******************************************

* Title
putexcel `rowlabel'`R'="Confidence to Seek Multipe ANC"
	local R=`R'+1

* Percents
tab mch_255_cat county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_low_Bomi=table[1,1]
	matrix female_low_Bong=table[1,2]
	matrix female_low_Total=female_low_Bomi[1,1]+female_low_Bong[1,1]
	matrix female_medium_Bomi=table[2,1]
	matrix female_medium_Bong=table[2,2]
	matrix female_medium_Total=female_medium_Bomi[1,1]+female_medium_Bong[1,1]
	matrix female_high_Bomi=table[3,1]
	matrix female_high_Bong=table[3,2]
	matrix female_high_Total=female_high_Bomi[1,1]+female_high_Bong[1,1]
	
	matrix female_Bomi_Total= female_low_Bomi[1,1]+ female_medium_Bomi[1,1]+ female_high_Bomi[1,1]
	matrix female_Bong_Total= female_low_Bong[1,1]+ female_medium_Bong[1,1]+ female_high_Bong[1,1]
	matrix female_Total_Total=female_low_Total[1,1]+female_medium_Total[1,1]+female_high_Total[1,1]
	
	matrix    p_low_Bomi_female=   female_low_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_medium_Bomi_female=female_medium_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix   p_high_Bomi_female=  female_high_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix    p_low_Bong_female=   female_low_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_medium_Bong_female=female_medium_Bong[1,1]/female_Bong_Total[1,1]
	matrix   p_high_Bong_female=  female_high_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix    p_low_Total_female=   female_low_Total[1,1]/female_Total_Total[1,1]
	matrix p_medium_Total_female=female_medium_Total[1,1]/female_Total_Total[1,1]
	matrix   p_high_Total_female=  female_high_Total[1,1]/female_Total_Total[1,1]
	
tab mch_255_cat if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** ≤3 months
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** 4-5 months
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** ≥6 months
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1


** Perceived Self-Efficacy for Spouse
*******************************************

* Title
putexcel `rowlabel'`R'="Confidence to Have Partner Attend ANC"
	local R=`R'+1

* Percents
tab mch_256_cat county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_low_Bomi=table[1,1]
	matrix female_low_Bong=table[1,2]
	matrix female_low_Total=female_low_Bomi[1,1]+female_low_Bong[1,1]
	matrix female_medium_Bomi=table[2,1]
	matrix female_medium_Bong=table[2,2]
	matrix female_medium_Total=female_medium_Bomi[1,1]+female_medium_Bong[1,1]
	matrix female_high_Bomi=table[3,1]
	matrix female_high_Bong=table[3,2]
	matrix female_high_Total=female_high_Bomi[1,1]+female_high_Bong[1,1]
	
	matrix female_Bomi_Total= female_low_Bomi[1,1]+ female_medium_Bomi[1,1]+ female_high_Bomi[1,1]
	matrix female_Bong_Total= female_low_Bong[1,1]+ female_medium_Bong[1,1]+ female_high_Bong[1,1]
	matrix female_Total_Total=female_low_Total[1,1]+female_medium_Total[1,1]+female_high_Total[1,1]
	
	matrix    p_low_Bomi_female=   female_low_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_medium_Bomi_female=female_medium_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix   p_high_Bomi_female=  female_high_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix    p_low_Bong_female=   female_low_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_medium_Bong_female=female_medium_Bong[1,1]/female_Bong_Total[1,1]
	matrix   p_high_Bong_female=  female_high_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix    p_low_Total_female=   female_low_Total[1,1]/female_Total_Total[1,1]
	matrix p_medium_Total_female=female_medium_Total[1,1]/female_Total_Total[1,1]
	matrix   p_high_Total_female=  female_high_Total[1,1]/female_Total_Total[1,1]
	
tab mch_256_cat if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** ≤3 months
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** 4-5 months
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** ≥6 months
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1


** Norms Around Early ANC
*******************************************
	
* Title
putexcel `rowlabel'`R'="Community Norms on Early ANC"
	local R=`R'+1

* Percents
tab mch_267 county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_low_Bomi=table[1,1]
	matrix female_low_Bong=table[1,2]
	matrix female_low_Total=female_low_Bomi[1,1]+female_low_Bong[1,1]
	matrix female_medium_Bomi=table[2,1]
	matrix female_medium_Bong=table[2,2]
	matrix female_medium_Total=female_medium_Bomi[1,1]+female_medium_Bong[1,1]
	matrix female_high_Bomi=table[3,1]
	matrix female_high_Bong=table[3,2]
	matrix female_high_Total=female_high_Bomi[1,1]+female_high_Bong[1,1]
	
	matrix female_Bomi_Total= female_low_Bomi[1,1]+ female_medium_Bomi[1,1]+ female_high_Bomi[1,1]
	matrix female_Bong_Total= female_low_Bong[1,1]+ female_medium_Bong[1,1]+ female_high_Bong[1,1]
	matrix female_Total_Total=female_low_Total[1,1]+female_medium_Total[1,1]+female_high_Total[1,1]
	
	matrix    p_low_Bomi_female=   female_low_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_medium_Bomi_female=female_medium_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix   p_high_Bomi_female=  female_high_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix    p_low_Bong_female=   female_low_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_medium_Bong_female=female_medium_Bong[1,1]/female_Bong_Total[1,1]
	matrix   p_high_Bong_female=  female_high_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix    p_low_Total_female=   female_low_Total[1,1]/female_Total_Total[1,1]
	matrix p_medium_Total_female=female_medium_Total[1,1]/female_Total_Total[1,1]
	matrix   p_high_Total_female=  female_high_Total[1,1]/female_Total_Total[1,1]
	
tab mch_267 if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** ≤3 months
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** 4-5 months
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** ≥6 months
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1	
	
	
** Norms Around Multiple ANC
*******************************************
	
* Title
putexcel `rowlabel'`R'="Community Norms on Women Attending At Least 4 ANC"
	local R=`R'+1

* Percents
tab mch_268 county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_low_Bomi=table[1,1]
	matrix female_low_Bong=table[1,2]
	matrix female_low_Total=female_low_Bomi[1,1]+female_low_Bong[1,1]
	matrix female_medium_Bomi=table[2,1]
	matrix female_medium_Bong=table[2,2]
	matrix female_medium_Total=female_medium_Bomi[1,1]+female_medium_Bong[1,1]
	matrix female_high_Bomi=table[3,1]
	matrix female_high_Bong=table[3,2]
	matrix female_high_Total=female_high_Bomi[1,1]+female_high_Bong[1,1]
	
	matrix female_Bomi_Total= female_low_Bomi[1,1]+ female_medium_Bomi[1,1]+ female_high_Bomi[1,1]
	matrix female_Bong_Total= female_low_Bong[1,1]+ female_medium_Bong[1,1]+ female_high_Bong[1,1]
	matrix female_Total_Total=female_low_Total[1,1]+female_medium_Total[1,1]+female_high_Total[1,1]
	
	matrix    p_low_Bomi_female=   female_low_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_medium_Bomi_female=female_medium_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix   p_high_Bomi_female=  female_high_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix    p_low_Bong_female=   female_low_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_medium_Bong_female=female_medium_Bong[1,1]/female_Bong_Total[1,1]
	matrix   p_high_Bong_female=  female_high_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix    p_low_Total_female=   female_low_Total[1,1]/female_Total_Total[1,1]
	matrix p_medium_Total_female=female_medium_Total[1,1]/female_Total_Total[1,1]
	matrix   p_high_Total_female=  female_high_Total[1,1]/female_Total_Total[1,1]
	
tab mch_268 if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** ≤3 months
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** 4-5 months
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** ≥6 months
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1		
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
