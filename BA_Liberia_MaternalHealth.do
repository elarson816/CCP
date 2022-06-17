** Maternal Health Tables
cd "$directory"
local user $user

* Set macros for cells
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
sum one if county_id==23 & female==1 & haschild2yrs==1
	local n_Bong_female=r(N)
sum one if county_id==22 & female==1 & haschild2yrs==1
	local n_Bomi_female=r(N)
sum one if control==0 & female==1 & haschild2yrs==1
	local n_Total_female=r(N)
sum one if county_id==24 & female==1 & haschild2yrs==1
	local n_Gbarpolu_female=r(N)
	
sum one if county_id==23 & female==0 & haschild2yrs==1
	local n_Bong_male=r(N)
sum one if county_id==22 & female==0 & haschild2yrs==1
	local n_Bomi_male=r(N)
sum one if control==0 & female==0 & haschild2yrs==1
	local n_Total_male=r(N)
sum one if county_id==24 & female==0 & haschild2yrs==1
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
	
	local R=`R'+1
	
* Sample
putexcel `rowlabel'`R'="Sample: Men and women who were not excluded, had a child in the last 2 years and responded to question 217"


*********************************************************************
*** Table 5.2: FACTORS ASSOCIATED WITH USE OF EARLY PRENATAL CARE ***
*********************************************************************

putexcel set "$putexcel_set", modify sheet("Table5.2")
local R $R

* Titles
putexcel `women_collabel'`R'="Women's (20-49 years)"
	local R=`R'+1
	
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
	
** At Least One ANC (Total)
*******************************************

* Title
putexcel `rowlabel'`R'="At Least One ANC"

* Percents
tab mch_215 county_id if female==1 & haschild2yrs==1, mi matcell(table)
	matrix Bong_yes=table[1,2]+table[2,2]+table[3,2]
	matrix Bong_total=table[1,2]+table[2,2]+table[3,2]+table[4,2]
	
	matrix Bomi_yes=table[1,1]+table[2,1]+table[3,1]
	matrix Bomi_total=table[1,1]+table[2,1]+table[3,1]+table[4,1]
	
	matrix Total_yes=Bong_yes[1,1]+Bomi_yes[1,1]
	matrix Total_total=Bong_total[1,1]+Bomi_total[1,1]

	matrix Gbarpolu_yes=table[1,3]+table[2,3]+table[3,3]
	matrix Gbarpolu_total=table[1,3]+table[2,3]+table[3,3]+table[4,3]
	
	
	matrix p_Bong_yes=Bong_yes[1,1]/Bong_total[1,1]
	matrix p_Bomi_yes=Bomi_yes[1,1]/Bomi_total[1,1]
	matrix p_Total_yes=Total_yes[1,1]/Total_total[1,1]
	matrix p_Gbarpolu_yes=Gbarpolu_yes[1,1]/Gbarpolu_total[1,1]
	
* PutExcel
putexcel `women_intervention_col1'`R'=matrix(p_Bong_yes)
putexcel `women_intervention_col2'`R'=matrix(p_Bomi_yes)
putexcel `women_intervention_col3'`R'=matrix(p_Total_yes)
putexcel `women_control_col1'`R'=matrix(p_Gbarpolu_yes)

local R=`R'+1


** Early ANC (Total)
*******************************************

* Title
putexcel `rowlabel'`R'="Early ANC"

* Percents
tab mch_217_cat county_id if female==1 & haschild2yrs==1, matcell(table)
	matrix Bomi_3mo=table[1,1]
	matrix Bomi_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix Bong_3mo=table[1,2]
	matrix Bong_total=table[1,2]+table[2,2]+table[3,2]
	
	matrix Total_3mo=Bomi_3mo[1,1]+Bong_3mo[1,1]
	matrix Total_total=Bomi_total[1,1]+Bong_total[1,1]
	
	matrix Gbarpolu_3mo=table[1,3]
	matrix Gbarpolu_total=table[1,3]+table[2,3]+table[3,3]
	
	
	matrix     p_Bong_3mo=    Bong_3mo[1,1]/Bong_total[1,1]
	matrix     p_Bomi_3mo=    Bomi_3mo[1,1]/Bomi_total[1,1]
	matrix    p_Total_3mo=   Total_3mo[1,1]/Total_total[1,1]
	matrix p_Gbarpolu_3mo=Gbarpolu_3mo[1,1]/Gbarpolu_total[1,1]
	
* PutExcel
putexcel `women_intervention_col1'`R'= matrix(p_Bong_3mo)
putexcel `women_intervention_col2'`R'= matrix(p_Bomi_3mo)
putexcel `women_intervention_col3'`R'=matrix(p_Total_3mo)
putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_3mo)

local R=`R'+1

** Not Early ANC
*******************************************

* Title
putexcel `rowlabel'`R'="Not Early ANC"

* Percents
tab mch_217_cat county_id if female==1 & haschild2yrs==1, matcell(table)
	matrix Bong_4plusmo=table[2,1]+table[3,1]
	matrix Bong_total=table[1,2]+table[2,2]+table[3,2]
	
	matrix Bomi_4plusmo=table[2,2]+table[3,2]
	matrix Bomi_total=table[1,2]+table[2,2]+table[3,2]
	
	matrix Total_4plusmo=Bong_4plusmo[1,1]+Bomi_4plusmo[1,1]
	matrix Total_total=Bong_total[1,1]+Bomi_total[1,1]
	
	matrix Gbarpolu_4plusmo=table[2,3]+table[3,3]
	matrix Gbarpolu_total=table[1,3]+table[2,3]+table[3,3]
	
	
	matrix     p_Bong_4plusmo=    Bong_4plusmo[1,1]/Bong_total[1,1]
	matrix     p_Bomi_4plusmo=    Bomi_4plusmo[1,1]/Bomi_total[1,1]
	matrix    p_Total_4plusmo=   Total_4plusmo[1,1]/Total_total[1,1]
	matrix p_Gbarpolu_4plusmo=Gbarpolu_4plusmo[1,1]/Gbarpolu_total[1,1]
	
* PutExcel
putexcel `women_intervention_col1'`R'= matrix(p_Bong_4plusmo)
putexcel `women_intervention_col2'`R'= matrix(p_Bomi_4plusmo)
putexcel `women_intervention_col3'`R'=matrix(p_Total_4plusmo)
putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_4plusmo)

local R=`R'+1

** Reasons For No ANC
*******************************************

* Title 
putexcel `rowlabel'`R'="Reasons for not seeking early ANC"
	local R=`R'+1

* Percents
foreach reason in 1 2 3 4 6 7 8 15 97 {
	tab mch_216_`reason' county_id if female==1 & haschild2yrs==1 & mch_215==1, matcell(table)
	matrix Bomi_Yes_`reason'=table[2,1]
	matrix Bomi_No_`reason'=table[1,1]
	matrix Bomi_Total_`reason'=Bomi_Yes_`reason'[1,1]+Bomi_No_`reason'[1,1]
	
	matrix Bong_Yes_`reason'=table[2,2]
	matrix Bong_No_`reason'=table[2,2]
	matrix Bong_Total_`reason'=Bong_Yes_`reason'[1,1]+Bomi_No_`reason'[1,1]
	
	matrix Total_Yes_`reason'=Bomi_Yes_`reason'[1,1]+Bong_Yes_`reason'[1,1]
	matrix Total_No_`reason'=Bomi_No_`reason'[1,1]+Bong_No_`reason'[1,1]
	matrix Total_Total_`reason'=Bomi_Total_`reason'[1,1]+Bong_Total_`reason'[1,1]
	
	matrix Gbarpolu_Yes_`reason'=table[2,3]
	matrix Gbarpolu_No_`reason'=table[1,3]
	matrix Gbarpolu_Total_`reason'=Gbarpolu_Yes_`reason'[1,1]+Gbarpolu_No_`reason'[1,1]
	
	matrix     p_Bomi_Yes_`reason'=    Bomi_Yes_`reason'[1,1]/    Bomi_Total_`reason'[1,1]
	matrix     p_Bong_Yes_`reason'=    Bong_Yes_`reason'[1,1]/    Bong_Total_`reason'[1,1]
	matrix    p_Total_Yes_`reason'=   Total_Yes_`reason'[1,1]/   Total_Total_`reason'[1,1]
	matrix p_Gbarpolu_Yes_`reason'=Gbarpolu_Yes_`reason'[1,1]/Gbarpolu_Total_`reason'[1,1]
	}
	
* PutExcel
	
	** Cost
	putexcel `rowlabel'`R'="Cost"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_1)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_1)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_1)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_1)
		local R=`R'+1
		
	** No Facility
	putexcel `rowlabel'`R'="No Facility"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_2)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_2)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_2)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_2)
		local R=`R'+1
		
	** Too far/No Transportation
	putexcel `rowlabel'`R'="Too far/No Transportation"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_3)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_3)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_3)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_3)
		local R=`R'+1
		
	** Don't trust the facility/Poor service quality
	putexcel `rowlabel'`R'="Don't trust the facility/poor service quality"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_4)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_4)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_4)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_4)
		local R=`R'+1
		
	** Not the first child
	putexcel `rowlabel'`R'="Not the first child"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_6)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_6)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_6)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_6)
		local R=`R'+1
		
	** Not necessary
	putexcel `rowlabel'`R'="Not Necessary"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_7)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_7)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_7)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_7)
		local R=`R'+1
		
	** Spouse/Partner didn't think it was necessary
	putexcel `rowlabel'`R'="Spouse/Partner didn't think it was necessary"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_8)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_8)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_8)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_8)
		local R=`R'+1
		
	** Afraid to go
	putexcel `rowlabel'`R'="Afraid to go"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_15)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_15)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_15)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_15)
		local R=`R'+1
		
	** Other
	putexcel `rowlabel'`R'="Other"
	putexcel `women_intervention_col1'`R'= matrix(p_Bong_Yes_97)
	putexcel `women_intervention_col2'`R'= matrix(p_Bomi_Yes_97)
	putexcel `women_intervention_col3'`R'=matrix(p_Total_Yes_97)
	putexcel `women_control_col1'`R'=  matrix(p_Gbarpolu_Yes_97)
		local R=`R'+1


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

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
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

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
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

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
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

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	**  high
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

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1		
	
	
	
** Household Support - Workload
*******************************************	
	
* Title
putexcel `rowlabel'`R'="Workload during pregnancy"
	local R=`R'+1

* Percents
tab ls_510 county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
	matrix female_less_Bomi=table[1,1]
	matrix female_less_Bong=table[1,2]
	matrix female_less_Total=female_less_Bomi[1,1]+female_less_Bong[1,1]
	matrix female_same_Bomi=table[2,1]
	matrix female_same_Bong=table[2,2]
	matrix female_same_Total=female_same_Bomi[1,1]+female_same_Bong[1,1]
	matrix female_more_Bomi=table[3,1]
	matrix female_more_Bong=table[3,2]
	matrix female_more_Total=female_more_Bomi[1,1]+female_more_Bong[1,1]
	
	matrix female_Bomi_Total= female_less_Bomi[1,1]+ female_same_Bomi[1,1]+ female_more_Bomi[1,1]
	matrix female_Bong_Total= female_less_Bong[1,1]+ female_same_Bong[1,1]+ female_more_Bong[1,1]
	matrix female_Total_Total=female_less_Total[1,1]+female_same_Total[1,1]+female_more_Total[1,1]
	
	matrix p_less_Bomi_female=female_less_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_same_Bomi_female=female_same_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_more_Bomi_female=female_more_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix p_less_Bong_female=female_less_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_same_Bong_female=female_same_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_more_Bong_female=female_more_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix p_less_Total_female=female_less_Total[1,1]/female_Total_Total[1,1]
	matrix p_same_Total_female=female_same_Total[1,1]/female_Total_Total[1,1]
	matrix p_more_Total_female=female_more_Total[1,1]/female_Total_Total[1,1]
	
tab ls_510 if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_less_Gbarpolu=table[1,1]
	matrix female_same_Gbarpolu=table[2,1]
	matrix female_more_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_less_Gbarpolu[1,1]+female_same_Gbarpolu[1,1]+female_more_Gbarpolu[1,1]
	
	matrix p_less_Gbarpolu_female=female_less_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_same_Gbarpolu_female=female_same_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_more_Gbarpolu_female=female_more_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** low
	putexcel `rowlabel'`R'="Less"
	putexcel `women_intervention_col1'`R'=  matrix(p_less_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_less_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_less_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_less_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Same"
	putexcel `women_intervention_col1'`R'=  matrix(p_same_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_same_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_same_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_same_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
	putexcel `rowlabel'`R'="More"
	putexcel `women_intervention_col1'`R'=  matrix(p_more_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_more_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_more_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_more_Gbarpolu_female)
	
	local R=`R'+1			
	
	
** Household Support - Support Provider
*******************************************	
	
* Title
putexcel `rowlabel'`R'="Person who provided support during pregnancy"
	local R=`R'+1
	
* Recode ls_512 -98 to missing
recode ls_512 (-98=.)

* Percents
tab ls_512 county_id if female==1 & control==0 & haschild2yrs==1 & ls_511==1, matcell(table)
	matrix female_partner_Bomi=table[1,1]
	matrix female_partner_Bong=table[1,2]
	matrix female_partner_Total=female_partner_Bomi[1,1]+female_partner_Bong[1,1]
	matrix female_mil_Bomi=table[2,1]
	matrix female_mil_Bong=table[2,2]
	matrix female_mil_Total=female_mil_Bomi[1,1]+female_mil_Bong[1,1]
	matrix female_fil_Bomi=table[3,1]
	matrix female_fil_Bong=table[3,2]
	matrix female_fil_Total=female_fil_Bomi[1,1]+female_fil_Bong[1,1]
	matrix female_children_Bomi=table[4,1]
	matrix female_children_Bong=table[4,2]
	matrix female_children_Total=female_children_Bomi[1,1]+female_children_Bong[1,1]
	matrix female_sibling_Bomi=table[5,1]
	matrix female_sibling_Bong=table[5,2]
	matrix female_sibling_Total=female_sibling_Bomi[1,1]+female_sibling_Bong[1,1]
	matrix female_other_Bomi=table[6,1]
	matrix female_other_Bong=table[6,2]
	matrix female_other_Total=female_other_Bomi[1,1]+female_other_Bong[1,1]	
	
	matrix female_Bomi_Total= female_partner_Bomi[1,1]+ female_mil_Bomi[1,1]+ female_fil_Bomi[1,1]+ female_children_Bomi[1,1]+ female_sibling_Bomi[1,1]+ female_other_Bomi[1,1]   
	matrix female_Bong_Total= female_partner_Bong[1,1]+ female_mil_Bong[1,1]+ female_fil_Bong[1,1]+ female_children_Bong[1,1]+ female_sibling_Bong[1,1]+ female_other_Bong[1,1]
	matrix female_Total_Total=female_partner_Total[1,1]+female_mil_Total[1,1]+female_fil_Total[1,1]+female_children_Total[1,1]+female_sibling_Total[1,1]+female_other_Total[1,1]
	
	matrix  p_partner_Bomi_female= female_partner_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix      p_mil_Bomi_female=     female_mil_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix      p_fil_Bomi_female=     female_fil_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix p_children_Bomi_female=female_children_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix  p_sibling_Bomi_female= female_sibling_Bomi[1,1]/female_Bomi_Total[1,1]
	matrix    p_other_Bomi_female=   female_other_Bomi[1,1]/female_Bomi_Total[1,1]
	
	matrix  p_partner_Bong_female= female_partner_Bong[1,1]/female_Bong_Total[1,1]
	matrix      p_mil_Bong_female=     female_mil_Bong[1,1]/female_Bong_Total[1,1]
	matrix      p_fil_Bong_female=     female_fil_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_children_Bong_female=female_children_Bong[1,1]/female_Bong_Total[1,1]
	matrix  p_sibling_Bong_female= female_sibling_Bong[1,1]/female_Bong_Total[1,1]
	matrix    p_other_Bong_female=   female_other_Bong[1,1]/female_Bong_Total[1,1]
	
	matrix  p_partner_Total_female= female_partner_Total[1,1]/female_Total_Total[1,1]
	matrix      p_mil_Total_female=     female_mil_Total[1,1]/female_Total_Total[1,1]
	matrix      p_fil_Total_female=     female_fil_Total[1,1]/female_Total_Total[1,1]
	matrix p_children_Total_female=female_children_Total[1,1]/female_Total_Total[1,1]
	matrix  p_sibling_Total_female= female_sibling_Total[1,1]/female_Total_Total[1,1]
	matrix    p_other_Total_female=   female_other_Total[1,1]/female_Total_Total[1,1]
	
tab ls_512 if female==1 & control==1 & haschild2yrs==1 & ls_511==1, matcell(table)
	matrix female_partner_Gbarpolu=table[1,1]
	matrix female_mil_Gbarpolu=table[2,1]
	matrix female_fil_Gbarpolu=table[3,1]
	matrix female_children_Gbarpolu=table[3,1]
	matrix female_sibling_Gbarpolu=table[3,1]
	matrix female_other_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_partner_Gbarpolu[1,1]+female_mil_Gbarpolu[1,1]+female_fil_Gbarpolu[1,1]+female_children_Gbarpolu[1,1]+female_sibling_Gbarpolu[1,1]+female_other_Gbarpolu[1,1]
	
	matrix  p_partner_Gbarpolu_female= female_partner_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix      p_mil_Gbarpolu_female=     female_mil_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix      p_fil_Gbarpolu_female=     female_fil_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_children_Gbarpolu_female=female_children_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix  p_sibling_Gbarpolu_female= female_sibling_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix    p_other_Gbarpolu_female=   female_other_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** Partner
	putexcel `rowlabel'`R'="Partner"
	putexcel `women_intervention_col1'`R'=  matrix(p_partner_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_partner_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_partner_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_partner_Gbarpolu_female)

	local R=`R'+1
	
	** Mother-in-Law
	putexcel `rowlabel'`R'="Mother-in-Law"
	putexcel `women_intervention_col1'`R'=  matrix(p_mil_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_mil_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_mil_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_mil_Gbarpolu_female)
	

	local R=`R'+1

	** Father-in-Law
	putexcel `rowlabel'`R'="Father-in-Law"
	putexcel `women_intervention_col1'`R'=  matrix(p_fil_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_fil_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_fil_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_fil_Gbarpolu_female)
	

	local R=`R'+1
	
	** Children
	putexcel `rowlabel'`R'="Children"
	putexcel `women_intervention_col1'`R'=  matrix(p_children_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_children_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_children_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_children_Gbarpolu_female)
	
	local R=`R'+1		
	
	** Sibling
	putexcel `rowlabel'`R'="Sibling"
	putexcel `women_intervention_col1'`R'=  matrix(p_sibling_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_sibling_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_sibling_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_sibling_Gbarpolu_female)
	
	local R=`R'+1		
		
	** Other
	putexcel `rowlabel'`R'="Other"
	putexcel `women_intervention_col1'`R'=  matrix(p_other_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_other_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_other_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_other_Gbarpolu_female)
	
	local R=`R'+1		
		
	
** Household Support - Spousal Suport
*******************************************
	
* Title
putexcel `rowlabel'`R'="Amount of support provided by parnter"
	local R=`R'+1
	
* Recode -98 to missing
recode ls_513 (-98=.)

* Percents
tab ls_513_cat county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
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
	
tab ls_513_cat if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1	
	
** Household Support - Family Suport
*******************************************
	
* Title
putexcel `rowlabel'`R'="Amount of support provided by family"
	local R=`R'+1


* Percents
tab ls_514_cat county_id if female==1 & control==0 & haschild2yrs==1, matcell(table)
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
	
tab ls_514_cat if female==1 & control==1 & haschild2yrs==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
* PutExcel

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	

	local R=`R'+1
	
	** high
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	local R=`R'+1
	
* Sample
putexcel `rowlabel'`R'="Sample: Women who were not excluded and had a child in the last 2 years"		
	
	
*********************************************************************
*** Table 5.3: EXPOSURE TO HEALTH-RELATED MESSAGES ***
*********************************************************************	
	
putexcel set "$putexcel_set", modify sheet("Table5.3")
local R $R

* Titles
putexcel `women_collabel'`R'="Women's (20-49 years)"
putexcel `men_collabel'`R'  ="Men/Partners (20-55 years)"
	local R=`R'+1
putexcel `women_intervention_col1'`R'="Intervention Counties"
putexcel `women_control_col1'`R'     ="Control"

putexcel `men_intervention_col1'`R'="Intervention Counties"
putexcel `men_control_col1'`R'     ="Control"
	local R=`R'+1


* Counts
sum one if county_id==23 & female==1
	local n_Bong_female=r(N)
sum one if county_id==22 & female==1
	local n_Bomi_female=r(N)
sum one if control==0 & female==1
	local n_Total_female=r(N)
sum one if county_id==24 & female==1
	local n_Gbarpolu_female=r(N)
sum one if county_id==23 & female==0
	local n_Bong_male=r(N)
sum one if county_id==22 & female==0
	local n_Bomi_male=r(N)
sum one if control==0 & female==0
	local n_Total_male=r(N)
sum one if county_id==24 & female==0
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
	
** Respectful Maternity Care
*******************************************
	
* Title
putexcel `rowlabel'`R'="Has heard of respectful maternity care"
	local R=`R'+1
	
* Recode -98 to missing
recode mch_236 (-98=.)

* Percents
tab mch_236 county_id if female==1 & control==0, matcell(table)
	matrix female_no_Bomi=table[1,1]
	matrix female_no_Bong=table[1,2]
	matrix female_no_Total=female_no_Bomi[1,1]+female_no_Bong[1,1]
	matrix female_yes_Bomi=table[2,1]
	matrix female_yes_Bong=table[2,2]
	matrix female_yes_Total=female_yes_Bomi[1,1]+female_yes_Bong[1,1]
	matrix male_no_Bomi=table[1,1]
	matrix male_no_Bong=table[1,2]
	matrix male_no_Total=male_no_Bomi[1,1]+male_no_Bong[1,1]
	matrix male_yes_Bomi=table[2,1]
	matrix male_yes_Bong=table[2,2]
	matrix male_yes_Total=male_yes_Bomi[1,1]+male_yes_Bong[1,1]
	
	matrix female_Bomi_Total= female_no_Bomi[1,1]+ female_yes_Bomi[1,1] 
	matrix female_Bong_Total= female_no_Bong[1,1]+ female_yes_Bong[1,1] 
	matrix female_Total_Total=female_no_Total[1,1]+female_yes_Total[1,1]
	matrix male_Bomi_Total= male_no_Bomi[1,1]+ male_yes_Bomi[1,1] 
	matrix male_Bong_Total= male_no_Bong[1,1]+ male_yes_Bong[1,1] 
	matrix male_Total_Total=male_no_Total[1,1]+male_yes_Total[1,1]
	
	matrix  p_no_Bomi_female= female_no_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_yes_Bomi_female=female_yes_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix  p_no_Bomi_male= male_no_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_yes_Bomi_male=male_yes_Bomi[1,1]/male_Bomi_Total[1,1]
	
	matrix  p_no_Bong_female= female_no_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_yes_Bong_female=female_yes_Bong[1,1]/female_Bong_Total[1,1]
	matrix  p_no_Bong_male= male_no_Bong[1,1]/male_Bong_Total[1,1]
	matrix p_yes_Bong_male=male_yes_Bong[1,1]/male_Bong_Total[1,1]
	
	matrix  p_no_Total_female= female_no_Total[1,1]/female_Total_Total[1,1]
	matrix p_yes_Total_female=female_yes_Total[1,1]/female_Total_Total[1,1]
	matrix  p_no_Total_male= male_no_Total[1,1]/male_Total_Total[1,1]
	matrix p_yes_Total_male=male_yes_Total[1,1]/male_Total_Total[1,1]
	
tab mch_236 if female==1 & control==1, matcell(table)
	matrix female_no_Gbarpolu=table[1,1]
	matrix female_yes_Gbarpolu=table[2,1]
	matrix female_Gbarpolu_Total=female_no_Gbarpolu[1,1]+female_yes_Gbarpolu[1,1]
	matrix male_no_Gbarpolu=table[1,1]
	matrix male_yes_Gbarpolu=table[2,1]
	matrix male_Gbarpolu_Total=male_no_Gbarpolu[1,1]+male_yes_Gbarpolu[1,1]
	
	matrix  p_no_Gbarpolu_female= female_no_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_yes_Gbarpolu_female=female_yes_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix  p_no_Gbarpolu_male= male_no_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix p_yes_Gbarpolu_male=male_yes_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	
* PutExcel

	** No
	putexcel `rowlabel'`R'="No"
	putexcel `women_intervention_col1'`R'=  matrix(p_no_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_no_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_no_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_no_Gbarpolu_female)
	putexcel `men_intervention_col1'`R'=  matrix(p_no_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_no_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_no_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_no_Gbarpolu_male)

	local R=`R'+1
	
	** Yes
	putexcel `rowlabel'`R'="Yes"
	putexcel `women_intervention_col1'`R'=  matrix(p_yes_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_yes_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_yes_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_yes_Gbarpolu_female)
	putexcel `men_intervention_col1'`R'=  matrix(p_yes_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_yes_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_yes_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_yes_Gbarpolu_male)
	
	local R=`R'+1

** Share It, Act It
*******************************************
	
* Title
putexcel `rowlabel'`R'="Has heard of 'Share It, Act It'"
	local R=`R'+1
	
* Recode -98 to missing
recode me_918 (-98=.)

* Percents
tab me_918 county_id if female==1 & control==0, matcell(table)
	matrix female_no_Bomi=table[1,1]
	matrix female_no_Bong=table[1,2]
	matrix female_no_Total=female_no_Bomi[1,1]+female_no_Bong[1,1]
	matrix female_yes_Bomi=table[2,1]
	matrix female_yes_Bong=table[2,2]
	matrix female_yes_Total=female_yes_Bomi[1,1]+female_yes_Bong[1,1]
	matrix male_no_Bomi=table[1,1]
	matrix male_no_Bong=table[1,2]
	matrix male_no_Total=male_no_Bomi[1,1]+male_no_Bong[1,1]
	matrix male_yes_Bomi=table[2,1]
	matrix male_yes_Bong=table[2,2]
	matrix male_yes_Total=male_yes_Bomi[1,1]+male_yes_Bong[1,1]
	
	matrix female_Bomi_Total= female_no_Bomi[1,1]+ female_yes_Bomi[1,1] 
	matrix female_Bong_Total= female_no_Bong[1,1]+ female_yes_Bong[1,1] 
	matrix female_Total_Total=female_no_Total[1,1]+female_yes_Total[1,1]
	matrix male_Bomi_Total= male_no_Bomi[1,1]+ male_yes_Bomi[1,1] 
	matrix male_Bong_Total= male_no_Bong[1,1]+ male_yes_Bong[1,1] 
	matrix male_Total_Total=male_no_Total[1,1]+male_yes_Total[1,1]
	
	matrix  p_no_Bomi_female= female_no_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_yes_Bomi_female=female_yes_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix  p_no_Bomi_male= male_no_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_yes_Bomi_male=male_yes_Bomi[1,1]/male_Bomi_Total[1,1]
	
	matrix  p_no_Bong_female= female_no_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_yes_Bong_female=female_yes_Bong[1,1]/female_Bong_Total[1,1]
	matrix  p_no_Bong_male= male_no_Bong[1,1]/male_Bong_Total[1,1]
	matrix p_yes_Bong_male=male_yes_Bong[1,1]/male_Bong_Total[1,1]
	
	matrix  p_no_Total_female= female_no_Total[1,1]/female_Total_Total[1,1]
	matrix p_yes_Total_female=female_yes_Total[1,1]/female_Total_Total[1,1]
	matrix  p_no_Total_male= male_no_Total[1,1]/male_Total_Total[1,1]
	matrix p_yes_Total_male=male_yes_Total[1,1]/male_Total_Total[1,1]
	
tab me_918 if female==1 & control==1, matcell(table)
	matrix female_no_Gbarpolu=table[1,1]
	matrix female_yes_Gbarpolu=table[2,1]
	matrix female_Gbarpolu_Total=female_no_Gbarpolu[1,1]+female_yes_Gbarpolu[1,1]
	matrix male_no_Gbarpolu=table[1,1]
	matrix male_yes_Gbarpolu=table[2,1]
	matrix male_Gbarpolu_Total=male_no_Gbarpolu[1,1]+male_yes_Gbarpolu[1,1]
	
	matrix  p_no_Gbarpolu_female= female_no_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_yes_Gbarpolu_female=female_yes_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix  p_no_Gbarpolu_male= male_no_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix p_yes_Gbarpolu_male=male_yes_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	
* PutExcel

	** No
	putexcel `rowlabel'`R'="No"
	putexcel `women_intervention_col1'`R'=  matrix(p_no_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_no_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_no_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_no_Gbarpolu_female)
	putexcel `men_intervention_col1'`R'=  matrix(p_no_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_no_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_no_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_no_Gbarpolu_male)

	local R=`R'+1
	
	** Yes
	putexcel `rowlabel'`R'="Yes"
	putexcel `women_intervention_col1'`R'=  matrix(p_yes_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_yes_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_yes_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_yes_Gbarpolu_female)
	putexcel `men_intervention_col1'`R'=  matrix(p_yes_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_yes_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_yes_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_yes_Gbarpolu_male)

	local R=`R'+1	
	
** Healthy Life
*******************************************
	
* Title
putexcel `rowlabel'`R'="Has heard of 'Healthy Life'"
	local R=`R'+1
	
* Recode -98 to missing
recode me_919 (-98=.)

* Percents
tab me_919 county_id if female==1 & control==0, matcell(table)
	matrix female_no_Bomi=table[1,1]
	matrix female_no_Bong=table[1,2]
	matrix female_no_Total=female_no_Bomi[1,1]+female_no_Bong[1,1]
	matrix female_yes_Bomi=table[2,1]
	matrix female_yes_Bong=table[2,2]
	matrix female_yes_Total=female_yes_Bomi[1,1]+female_yes_Bong[1,1]
	matrix male_no_Bomi=table[1,1]
	matrix male_no_Bong=table[1,2]
	matrix male_no_Total=male_no_Bomi[1,1]+male_no_Bong[1,1]
	matrix male_yes_Bomi=table[2,1]
	matrix male_yes_Bong=table[2,2]
	matrix male_yes_Total=male_yes_Bomi[1,1]+male_yes_Bong[1,1]
	
	matrix female_Bomi_Total= female_no_Bomi[1,1]+ female_yes_Bomi[1,1] 
	matrix female_Bong_Total= female_no_Bong[1,1]+ female_yes_Bong[1,1] 
	matrix female_Total_Total=female_no_Total[1,1]+female_yes_Total[1,1]
	matrix male_Bomi_Total= male_no_Bomi[1,1]+ male_yes_Bomi[1,1] 
	matrix male_Bong_Total= male_no_Bong[1,1]+ male_yes_Bong[1,1] 
	matrix male_Total_Total=male_no_Total[1,1]+male_yes_Total[1,1]
	
	matrix  p_no_Bomi_female= female_no_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_yes_Bomi_female=female_yes_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix  p_no_Bomi_male= male_no_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_yes_Bomi_male=male_yes_Bomi[1,1]/male_Bomi_Total[1,1]
	
	matrix  p_no_Bong_female= female_no_Bong[1,1]/female_Bong_Total[1,1]
	matrix p_yes_Bong_female=female_yes_Bong[1,1]/female_Bong_Total[1,1]
	matrix  p_no_Bong_male= male_no_Bong[1,1]/male_Bong_Total[1,1]
	matrix p_yes_Bong_male=male_yes_Bong[1,1]/male_Bong_Total[1,1]
	
	matrix  p_no_Total_female= female_no_Total[1,1]/female_Total_Total[1,1]
	matrix p_yes_Total_female=female_yes_Total[1,1]/female_Total_Total[1,1]
	matrix  p_no_Total_male= male_no_Total[1,1]/male_Total_Total[1,1]
	matrix p_yes_Total_male=male_yes_Total[1,1]/male_Total_Total[1,1]
	
tab me_919 if female==1 & control==1, matcell(table)
	matrix female_no_Gbarpolu=table[1,1]
	matrix female_yes_Gbarpolu=table[2,1]
	matrix female_Gbarpolu_Total=female_no_Gbarpolu[1,1]+female_yes_Gbarpolu[1,1]
	matrix male_no_Gbarpolu=table[1,1]
	matrix male_yes_Gbarpolu=table[2,1]
	matrix male_Gbarpolu_Total=male_no_Gbarpolu[1,1]+male_yes_Gbarpolu[1,1]
	
	matrix  p_no_Gbarpolu_female= female_no_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_yes_Gbarpolu_female=female_yes_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix  p_no_Gbarpolu_male= male_no_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix p_yes_Gbarpolu_male=male_yes_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	
* PutExcel

	** No
	putexcel `rowlabel'`R'="No"
	putexcel `women_intervention_col1'`R'=  matrix(p_no_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_no_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_no_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_no_Gbarpolu_female)
	putexcel `men_intervention_col1'`R'=  matrix(p_no_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_no_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_no_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_no_Gbarpolu_male)

	local R=`R'+1
	
	** Yes
	putexcel `rowlabel'`R'="Yes"
	putexcel `women_intervention_col1'`R'=  matrix(p_yes_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_yes_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_yes_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_yes_Gbarpolu_female)
	putexcel `men_intervention_col1'`R'=  matrix(p_yes_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_yes_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_yes_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_yes_Gbarpolu_male)

	local R=`R'+1		
	
putexcel `rowlabel'`R'="Sample: Men and women who were not excluded"
	
