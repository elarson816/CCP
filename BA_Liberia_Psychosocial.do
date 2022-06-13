** Psychosocial
cd "$directory"
local user $user

* Set macros for cells
local R $R
local rowlabel $rowlabel
local intervention_collab $intervention_collab
local intervention_men $intervention_men
local intervention_women $intervention_women
local control_collab $control_collab
local control_men $control_men
local control_women $control_women
local total_collab $total_collab
local total_men $total_men
local total_women $total_women

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

* Run GEM .do file
run "/Users/`user'/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_GEM_CoupleCommunication.do"

********************************************************
*** Table 4.1: COUPLE COMMUNICATION ***
********************************************************

putexcel set "$putexcel_set", modify sheet("Table4.1")

* Titles
putexcel `rowlabel'`R'="Health Topic"
putexcel `intervention_collab'`R'="Intervention"
putexcel `control_collab'`R' ="Control"
putexcel `total_collab'`R'="Total"
	local R=`R'+1
	
* Counts
sum one if control==0 & female==0 & marriedcohabiting==1
	local n_intervention_male=r(N)
sum one if control==0 & female==1 & marriedcohabiting==1
	local n_intervention_female=r(N)
sum one if control==1 & female==0 & marriedcohabiting==1
	local n_control_male=r(N)
sum one if control==1 & female==1 & marriedcohabiting==1
	local n_control_female=r(N)
sum one if female==0 & marriedcohabiting==1
	local n_male=r(N)
sum one if female==1 & marriedcohabiting==1
	local n_female=r(N)
	
	** Tiles
	local celltext Men (n=`n_intervention_male')
		putexcel `intervention_men'`R'="`celltext'"
	local celltext Women (n=`n_intervention_female')
		putexcel `intervention_women'`R'="`celltext'"
	local celltext Men (n=`n_control_male')
		putexcel `control_men'`R'="`celltext'"
	local celltext Women (n=`n_control_female')
		putexcel `control_women'`R'="`celltext'"
	local celltext Men (n=`n_male')
		putexcel `total_men'`R'="`celltext'"
	local celltext Women (n=`n_female')
		putexcel `total_women'`R'="`celltext'"

	local R=`R'+1
	
** Family Planning
*******************************************

* Title
putexcel `rowlabel'`R'="Family Planning"

* Recode -98 to missing
recode ls_501_1_cat (-98=.)

* Percents
tab ls_501_1_cat control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_int=table[3,1]
	matrix female_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_item_high_cntrl=table[3,2]
	matrix female_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_item_high_int=  female_item_high_int[1,1]/  female_item_total_int[1,1]
	matrix p_female_item_high_cntrl=female_item_high_cntrl[1,1]/female_item_total_cntrl[1,1]
	
tab ls_501_1_cat control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_int=table[3,1]
	matrix male_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_item_high_cntrl=table[3,2]
	matrix male_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_item_high_int=  male_item_high_int[1,1]/  male_item_total_int[1,1]
	matrix p_male_item_high_cntrl=male_item_high_cntrl[1,1]/male_item_total_cntrl[1,1]
	
tab ls_501_1_cat if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_total=table[3,1]
	matrix female_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_item_high_total=female_item_high_total[1,1]/female_item_total_total[1,1]
	
tab ls_501_1_cat if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_total=table[3,1]
	matrix male_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_item_high_total=male_item_high_total[1,1]/male_item_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_item_high_int)
	putexcel `intervention_women'`R'=matrix(p_female_item_high_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_item_high_cntrl)
	putexcel `control_women'`R'=matrix(p_female_item_high_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_item_high_total)
	putexcel `total_women'`R'=matrix(p_female_item_high_total)

local R=`R'+1	
	
** Pregnancy
*******************************************

* Title
putexcel `rowlabel'`R'="Pregnancy"

* Recode -98 to missing
recode ls_501_3_cat (-98=.)

* Percents
tab ls_501_3_cat control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_int=table[3,1]
	matrix female_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_item_high_cntrl=table[3,2]
	matrix female_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_item_high_int=  female_item_high_int[1,1]/  female_item_total_int[1,1]
	matrix p_female_item_high_cntrl=female_item_high_cntrl[1,1]/female_item_total_cntrl[1,1]
	
tab ls_501_3_cat control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_int=table[3,1]
	matrix male_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_item_high_cntrl=table[3,2]
	matrix male_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_item_high_int=  male_item_high_int[1,1]/  male_item_total_int[1,1]
	matrix p_male_item_high_cntrl=male_item_high_cntrl[1,1]/male_item_total_cntrl[1,1]
	
tab ls_501_3_cat if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_total=table[3,1]
	matrix female_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_item_high_total=female_item_high_total[1,1]/female_item_total_total[1,1]
	
tab ls_501_3_cat if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_total=table[3,1]
	matrix male_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_item_high_total=male_item_high_total[1,1]/male_item_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_item_high_int)
	putexcel `intervention_women'`R'=matrix(p_female_item_high_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_item_high_cntrl)
	putexcel `control_women'`R'=matrix(p_female_item_high_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_item_high_total)
	putexcel `total_women'`R'=matrix(p_female_item_high_total)

local R=`R'+1	

** Sanitation
*******************************************

* Title
putexcel `rowlabel'`R'="Sanitation"

* Recode -98 to missing
recode ls_501_5_cat (-98=.)

* Percents
tab ls_501_5_cat control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_int=table[3,1]
	matrix female_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_item_high_cntrl=table[3,2]
	matrix female_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_item_high_int=  female_item_high_int[1,1]/  female_item_total_int[1,1]
	matrix p_female_item_high_cntrl=female_item_high_cntrl[1,1]/female_item_total_cntrl[1,1]
	
tab ls_501_5_cat control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_int=table[3,1]
	matrix male_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_item_high_cntrl=table[3,2]
	matrix male_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_item_high_int=  male_item_high_int[1,1]/  male_item_total_int[1,1]
	matrix p_male_item_high_cntrl=male_item_high_cntrl[1,1]/male_item_total_cntrl[1,1]
	
tab ls_501_5_cat if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_total=table[3,1]
	matrix female_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_item_high_total=female_item_high_total[1,1]/female_item_total_total[1,1]
	
tab ls_501_5_cat if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_total=table[3,1]
	matrix male_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_item_high_total=male_item_high_total[1,1]/male_item_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_item_high_int)
	putexcel `intervention_women'`R'=matrix(p_female_item_high_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_item_high_cntrl)
	putexcel `control_women'`R'=matrix(p_female_item_high_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_item_high_total)
	putexcel `total_women'`R'=matrix(p_female_item_high_total)

local R=`R'+1	

** Nutrition
*******************************************

* Title
putexcel `rowlabel'`R'="Nutrition"

* Recode -98 to missing
recode ls_501_6_cat (-98=.)

* Percents
tab ls_501_6_cat control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_int=table[3,1]
	matrix female_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_item_high_cntrl=table[3,2]
	matrix female_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_item_high_int=  female_item_high_int[1,1]/  female_item_total_int[1,1]
	matrix p_female_item_high_cntrl=female_item_high_cntrl[1,1]/female_item_total_cntrl[1,1]
	
tab ls_501_6_cat control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_int=table[3,1]
	matrix male_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_item_high_cntrl=table[3,2]
	matrix male_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_item_high_int=  male_item_high_int[1,1]/  male_item_total_int[1,1]
	matrix p_male_item_high_cntrl=male_item_high_cntrl[1,1]/male_item_total_cntrl[1,1]
	
tab ls_501_6_cat if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_total=table[3,1]
	matrix female_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_item_high_total=female_item_high_total[1,1]/female_item_total_total[1,1]
	
tab ls_501_6_cat if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_total=table[3,1]
	matrix male_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_item_high_total=male_item_high_total[1,1]/male_item_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_item_high_int)
	putexcel `intervention_women'`R'=matrix(p_female_item_high_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_item_high_cntrl)
	putexcel `control_women'`R'=matrix(p_female_item_high_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_item_high_total)
	putexcel `total_women'`R'=matrix(p_female_item_high_total)

local R=`R'+1	


** Malaria
*******************************************

* Title
putexcel `rowlabel'`R'="Malaria"

* Recode -98 to missing
recode ls_501_7_cat (-98=.)

* Percents
tab ls_501_7_cat control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_int=table[3,1]
	matrix female_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_item_high_cntrl=table[3,2]
	matrix female_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_item_high_int=  female_item_high_int[1,1]/  female_item_total_int[1,1]
	matrix p_female_item_high_cntrl=female_item_high_cntrl[1,1]/female_item_total_cntrl[1,1]
	
tab ls_501_7_cat control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_int=table[3,1]
	matrix male_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_item_high_cntrl=table[3,2]
	matrix male_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_item_high_int=  male_item_high_int[1,1]/  male_item_total_int[1,1]
	matrix p_male_item_high_cntrl=male_item_high_cntrl[1,1]/male_item_total_cntrl[1,1]
	
tab ls_501_7_cat if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_total=table[3,1]
	matrix female_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_item_high_total=female_item_high_total[1,1]/female_item_total_total[1,1]
	
tab ls_501_7_cat if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_total=table[3,1]
	matrix male_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_item_high_total=male_item_high_total[1,1]/male_item_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_item_high_int)
	putexcel `intervention_women'`R'=matrix(p_female_item_high_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_item_high_cntrl)
	putexcel `control_women'`R'=matrix(p_female_item_high_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_item_high_total)
	putexcel `total_women'`R'=matrix(p_female_item_high_total)

local R=`R'+1

** Child Health
*******************************************

* Title
putexcel `rowlabel'`R'="Child Health"

* Recode -98 to missing
recode ls_501_11_cat (-98=.)

* Percents
tab ls_501_11_cat control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_int=table[3,1]
	matrix female_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_item_high_cntrl=table[3,2]
	matrix female_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_item_high_int=  female_item_high_int[1,1]/  female_item_total_int[1,1]
	matrix p_female_item_high_cntrl=female_item_high_cntrl[1,1]/female_item_total_cntrl[1,1]
	
tab ls_501_11_cat control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_int=table[3,1]
	matrix male_item_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_item_high_cntrl=table[3,2]
	matrix male_item_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_item_high_int=  male_item_high_int[1,1]/  male_item_total_int[1,1]
	matrix p_male_item_high_cntrl=male_item_high_cntrl[1,1]/male_item_total_cntrl[1,1]
	
tab ls_501_11_cat if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_item_high_total=table[3,1]
	matrix female_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_item_high_total=female_item_high_total[1,1]/female_item_total_total[1,1]
	
tab ls_501_11_cat if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_item_high_total=table[3,1]
	matrix male_item_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_item_high_total=male_item_high_total[1,1]/male_item_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_item_high_int)
	putexcel `intervention_women'`R'=matrix(p_female_item_high_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_item_high_cntrl)
	putexcel `control_women'`R'=matrix(p_female_item_high_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_item_high_total)
	putexcel `total_women'`R'=matrix(p_female_item_high_total)
	
* Sample
local R=`R'+1
putexcel `rowlabel'`R'="Sample: Men and women who were not excluded and were in union"
	

********************************************************
*** Table 4.2: DECISION MAKING ***
********************************************************

local R=$R

putexcel set "$putexcel_set", modify sheet("Table4.2")

* Titles
putexcel `rowlabel'`R'="Health Topic"
putexcel `intervention_collab'`R'="Intervention"
putexcel `control_collab'`R' ="Control"
putexcel `total_collab'`R'="Total"
	local R=`R'+1
	
* Counts
sum one if control==0 & female==0 & marriedcohabiting==1
	local n_intervention_male=r(N)
sum one if control==0 & female==1 & marriedcohabiting==1
	local n_intervention_female=r(N)
sum one if control==1 & female==0 & marriedcohabiting==1
	local n_control_male=r(N)
sum one if control==1 & female==1 & marriedcohabiting==1
	local n_control_female=r(N)
sum one if female==0 & marriedcohabiting==1
	local n_male=r(N)
sum one if female==1 & marriedcohabiting==1
	local n_female=r(N)
	
	** Tiles
	local celltext Men (n=`n_intervention_male')
		putexcel `intervention_men'`R'="`celltext'"
	local celltext Women (n=`n_intervention_female')
		putexcel `intervention_women'`R'="`celltext'"
	local celltext Men (n=`n_control_male')
		putexcel `control_men'`R'="`celltext'"
	local celltext Women (n=`n_control_female')
		putexcel `control_women'`R'="`celltext'"
	local celltext Men (n=`n_male')
		putexcel `total_men'`R'="`celltext'"
	local celltext Women (n=`n_female')
		putexcel `total_women'`R'="`celltext'"

	local R=`R'+1

** Buy Soap
*******************************************

* Title
putexcel `rowlabel'`R'="Buy Soap"

* Percents
tab _505_1p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_1p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_1p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_1p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1	


** Buy fish/meat/vegetables
*******************************************

* Title
putexcel `rowlabel'`R'="Buy fish/meat/vegetables"

* Percents
tab _505_2p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_2p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_2p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_2p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1	


** Clothes for children
*******************************************

* Title
putexcel `rowlabel'`R'="Clothes for children"

* Percents
tab _505_3p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_3p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_3p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_3p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1

** What to cook
*******************************************

* Title
putexcel `rowlabel'`R'="What to cook"

* Percents
tab _505_4p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_4p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_4p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_4p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1

** How many children to have
*******************************************

* Title
putexcel `rowlabel'`R'="How many children to have"

* Percents
tab _505_5p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_5p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_5p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_5p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1


** Whether to use contraceptives
*******************************************

* Title
putexcel `rowlabel'`R'="Whether to use contraceptives"

* Percents
tab _505_6p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_6p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_6p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_6p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1

** Seek care in a health facility when respondent is ill
*******************************************

* Title
putexcel `rowlabel'`R'="Seek care in a health facility when respondent is ill"

* Percents
tab _505_7p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_7p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_7p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_7p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)

local R=`R'+1



** Seek care in a health facility for a sick child
*******************************************

* Title
putexcel `rowlabel'`R'="Seek care in a health facility for a sick child"

* Percents
tab _505_8p control if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_int=table[1,1]
	matrix female_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix female_self_cntrl=table[1,2]
	matrix female_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_female_self_int=  female_self_int[1,1]/  female_total_int[1,1]
	matrix p_female_self_cntrl=female_self_cntrl[1,1]/female_total_cntrl[1,1]

tab _505_8p control if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_int=table[1,1]
	matrix male_total_int=table[1,1]+table[2,1]+table[3,1]
	matrix male_self_cntrl=table[1,2]
	matrix male_total_cntrl=table[1,2]+table[2,2]+table[3,2]
	
	matrix p_male_self_int=  male_self_int[1,1]/  male_total_int[1,1]
	matrix p_male_self_cntrl=male_self_cntrl[1,1]/male_total_cntrl[1,1]
	
tab _505_8p if female==1 & marriedcohabiting==1, matcell(table)
	matrix female_self_total=table[1,1]
	matrix female_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_female_self_total=female_self_total[1,1]/female_total_total[1,1]
	
tab _505_8p if female==0 & marriedcohabiting==1, matcell(table)
	matrix male_self_total=table[1,1]
	matrix male_total_total=table[1,1]+table[2,1]+table[3,1]
	
	matrix p_male_self_total=male_self_total[1,1]/male_total_total[1,1]
	
* PutExcel
	
	** Intervention
	putexcel `intervention_men'`R'=    matrix(p_male_self_int)
	putexcel `intervention_women'`R'=matrix(p_female_self_int)
	
	** Control
	putexcel `control_men'`R'=    matrix(p_male_self_cntrl)
	putexcel `control_women'`R'=matrix(p_female_self_cntrl)

	** Total
	putexcel `total_men'`R'=    matrix(p_male_self_total)
	putexcel `total_women'`R'=matrix(p_female_self_total)
	
* Sample
local R=`R'+1
putexcel `rowlabel'`R'="Sample: Men and women who were not excluded and were in union"

********************************************************
*** Table 4.3: DESCRIPTIVE NORMS ***
********************************************************

local R=$R

putexcel set "$putexcel_set", modify sheet("Table4.3")

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

** Physically Harmed - Overall
*******************************************

* Title
putexcel `rowlabel'`R'="Norms Around Overall Physical Harm"
	local R=`R'+1

* Percents
tab ls_507 county_id if female==1 & control==0, matcell(table)
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
	
tab ls_507 if female==1 & control==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]

tab ls_507 county_id if female==0 & control==0, matcell(table)
	matrix male_low_Bomi=table[1,1]
	matrix male_low_Bong=table[1,2]
	matrix male_low_Total=male_low_Bomi[1,1]+male_low_Bong[1,1]
	matrix male_medium_Bomi=table[2,1]
	matrix male_medium_Bong=table[2,2]
	matrix male_medium_Total=male_medium_Bomi[1,1]+male_medium_Bong[1,1]
	matrix male_high_Bomi=table[3,1]
	matrix male_high_Bong=table[3,2]
	matrix male_high_Total=male_high_Bomi[1,1]+male_high_Bong[1,1]
	
	matrix male_Bomi_Total= male_low_Bomi[1,1]+ male_medium_Bomi[1,1]+ male_high_Bomi[1,1]
	matrix male_Bong_Total= male_low_Bong[1,1]+ male_medium_Bong[1,1]+ male_high_Bong[1,1]
	matrix male_Total_Total=male_low_Total[1,1]+male_medium_Total[1,1]+male_high_Total[1,1]
	
	matrix    p_low_Bomi_male=   male_low_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_medium_Bomi_male=male_medium_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix   p_high_Bomi_male=  male_high_Bomi[1,1]/male_Bomi_Total[1,1]
	
	matrix    p_low_Bong_male=   male_low_Bong[1,1]/male_Bong_Total[1,1]
	matrix p_medium_Bong_male=male_medium_Bong[1,1]/male_Bong_Total[1,1]
	matrix   p_high_Bong_male=  male_high_Bong[1,1]/male_Bong_Total[1,1]
	
	matrix    p_low_Total_male=   male_low_Total[1,1]/male_Total_Total[1,1]
	matrix p_medium_Total_male=male_medium_Total[1,1]/male_Total_Total[1,1]
	matrix   p_high_Total_male=  male_high_Total[1,1]/male_Total_Total[1,1]
	
tab ls_507 if female==1 & control==0, matcell(table)
	matrix male_low_Gbarpolu=table[1,1]
	matrix male_medium_Gbarpolu=table[2,1]
	matrix male_high_Gbarpolu=table[3,1]
	matrix male_Gbarpolu_Total=male_low_Gbarpolu[1,1]+male_medium_Gbarpolu[1,1]+male_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_male=   male_low_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_male=male_medium_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_male=  male_high_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
		
	
* PutExcel

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_low_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_low_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_low_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_low_Gbarpolu_male)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_medium_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_medium_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_medium_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_medium_Gbarpolu_male)
	

	local R=`R'+1
	
	** high
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_high_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_high_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_high_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_high_Gbarpolu_male)
	
	local R=`R'+1
	
** Physically Harmed - Pregnancy
*******************************************

* Title
putexcel `rowlabel'`R'="Norms Around Physical Harm During Pregnancy"
	local R=`R'+1

* Percents
tab ls_508 county_id if female==1 & control==0, matcell(table)
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
	
tab ls_508 if female==1 & control==1, matcell(table)
	matrix female_low_Gbarpolu=table[1,1]
	matrix female_medium_Gbarpolu=table[2,1]
	matrix female_high_Gbarpolu=table[3,1]
	matrix female_Gbarpolu_Total=female_low_Gbarpolu[1,1]+female_medium_Gbarpolu[1,1]+female_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_female=   female_low_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_female=female_medium_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_female=  female_high_Gbarpolu[1,1]/female_Gbarpolu_Total[1,1]
	
tab ls_508 county_id if female==0 & control==0, matcell(table)
	matrix male_low_Bomi=table[1,1]
	matrix male_low_Bong=table[1,2]
	matrix male_low_Total=male_low_Bomi[1,1]+male_low_Bong[1,1]
	matrix male_medium_Bomi=table[2,1]
	matrix male_medium_Bong=table[2,2]
	matrix male_medium_Total=male_medium_Bomi[1,1]+male_medium_Bong[1,1]
	matrix male_high_Bomi=table[3,1]
	matrix male_high_Bong=table[3,2]
	matrix male_high_Total=male_high_Bomi[1,1]+male_high_Bong[1,1]
	
	matrix male_Bomi_Total= male_low_Bomi[1,1]+ male_medium_Bomi[1,1]+ male_high_Bomi[1,1]
	matrix male_Bong_Total= male_low_Bong[1,1]+ male_medium_Bong[1,1]+ male_high_Bong[1,1]
	matrix male_Total_Total=male_low_Total[1,1]+male_medium_Total[1,1]+male_high_Total[1,1]
	
	matrix    p_low_Bomi_male=   male_low_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix p_medium_Bomi_male=male_medium_Bomi[1,1]/male_Bomi_Total[1,1]
	matrix   p_high_Bomi_male=  male_high_Bomi[1,1]/male_Bomi_Total[1,1]
	
	matrix    p_low_Bong_male=   male_low_Bong[1,1]/male_Bong_Total[1,1]
	matrix p_medium_Bong_male=male_medium_Bong[1,1]/male_Bong_Total[1,1]
	matrix   p_high_Bong_male=  male_high_Bong[1,1]/male_Bong_Total[1,1]
	
	matrix    p_low_Total_male=   male_low_Total[1,1]/male_Total_Total[1,1]
	matrix p_medium_Total_male=male_medium_Total[1,1]/male_Total_Total[1,1]
	matrix   p_high_Total_male=  male_high_Total[1,1]/male_Total_Total[1,1]
	
tab ls_508 if female==1 & control==0, matcell(table)
	matrix male_low_Gbarpolu=table[1,1]
	matrix male_medium_Gbarpolu=table[2,1]
	matrix male_high_Gbarpolu=table[3,1]
	matrix male_Gbarpolu_Total=male_low_Gbarpolu[1,1]+male_medium_Gbarpolu[1,1]+male_high_Gbarpolu[1,1]
	
	matrix    p_low_Gbarpolu_male=   male_low_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix p_medium_Gbarpolu_male=male_medium_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	matrix   p_high_Gbarpolu_male=  male_high_Gbarpolu[1,1]/male_Gbarpolu_Total[1,1]
	
* PutExcel

	** low
	putexcel `rowlabel'`R'="Low"
	putexcel `women_intervention_col1'`R'=  matrix(p_low_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_low_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_low_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_low_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_low_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_low_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_low_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_low_Gbarpolu_male)

	local R=`R'+1
	
	** medium
	putexcel `rowlabel'`R'="Medium"
	putexcel `women_intervention_col1'`R'=  matrix(p_medium_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_medium_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_medium_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_medium_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_medium_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_medium_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_medium_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_medium_Gbarpolu_male)
	

	local R=`R'+1
	
	** high
	putexcel `rowlabel'`R'="High"
	putexcel `women_intervention_col1'`R'=  matrix(p_high_Bong_female)
	putexcel `women_intervention_col2'`R'=  matrix(p_high_Bomi_female)
	putexcel `women_intervention_col3'`R'=  matrix(p_high_Total_female)
	putexcel `women_control_col1'`R'=       matrix(p_high_Gbarpolu_female)
	
	putexcel `men_intervention_col1'`R'=  matrix(p_high_Bong_male)
	putexcel `men_intervention_col2'`R'=  matrix(p_high_Bomi_male)
	putexcel `men_intervention_col3'`R'=  matrix(p_high_Total_male)
	putexcel `men_control_col1'`R'=       matrix(p_high_Gbarpolu_male)
	
* Sample
local R=`R'+1
putexcel `rowlabel'`R'="Sample: Men and women who were not excluded"
