** Media Sources
cd "$directory"
local user $user

* Set macros for cells
local R $R
local rowlabel $rowlabel
local women_collab $women_collab
local women_intervention $women_intervention
local women_control $women_control
local men_collab $men_collab
local men_intervention $men_intervention
local men_control $men_control



********************************************************
*** Table 9.1: EXPOSURE TO MEDIA ***
********************************************************

putexcel set "$putexcel_set", modify sheet("Table9.1")
local R $R

* Titles
putexcel `women_collab'`R'="Women"
putexcel `men_collab'`R'="Men"
	local R=`R'+1
	
	
* Counts
sum one if control==0 & female==1
	local n_Intervention_female=r(N)
sum one if control==1 & female==1
	local n_Control_female=r(N)
sum one if control==0 & female==0
	local n_Intervention_male=r(N)
sum one if control==1 & female==0
	local n_Control_male=r(N)
	
	** Tiles
	local celltext Intervention N=`n_Intervention_female'
		putexcel `women_intervention'`R'="`celltext'"
	local celltext Control N=`n_Control_female' 
		putexcel `women_control'`R'="`celltext'"
	local celltext Intervention N=`n_Intervention_male' 
		putexcel `men_intervention'`R'="`celltext'"
	local celltext Control N=`n_Control_male' 
		putexcel `men_control'`R'="`celltext'"

local R=`R'+1	

** Listen to radio
*******************************************

* Title
putexcel `rowlabel'`R'="Frequency of listening to the radio"
local R=`R'+1

* Combine 901 values
gen me_901_cat=me_901
recode me_901_cat (1 2 3=1) (4=2) (5=3)

* Percents
tab me_901_cat control if female==1, matcell(table)
	matrix female_weekly_int= table[1,1]
	matrix female_monthly_int=table[2,1]
	matrix female_never_int=  table[3,1]
	matrix female_total_int=female_weekly_int[1,1]+female_monthly_int[1,1]+female_never_int[1,1]

	matrix female_weekly_cntrl =table[1,2]
	matrix female_monthly_cntrl=table[2,2]
	matrix female_never_cntrl  =table[3,2]
	matrix female_total_cntrl=female_weekly_cntrl[1,1]+female_monthly_cntrl[1,1]+female_never_cntrl[1,1]
	
	matrix  p_female_weekly_int= female_weekly_int[1,1]/female_total_int[1,1]
	matrix p_female_monthly_int=female_monthly_int[1,1]/female_total_int[1,1]
	matri    p_female_never_int=  female_never_int[1,1]/female_total_int[1,1]

	matrix   p_female_weekly_cntrl= female_weekly_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_monthly_cntrl=female_monthly_cntrl[1,1]/female_total_cntrl[1,1]
	matrix    p_female_never_cntrl=  female_never_cntrl[1,1]/female_total_cntrl[1,1]

tab me_901_cat control if female==0, matcell(table)
	matrix male_weekly_int= table[1,1]
	matrix male_monthly_int=table[2,1]
	matrix male_never_int=  table[3,1]
	matrix male_total_int=male_weekly_int[1,1]+male_monthly_int[1,1]+male_never_int[1,1]

	matrix male_weekly_cntrl =table[1,2]
	matrix male_monthly_cntrl=table[2,2]
	matrix male_never_cntrl  =table[3,2]
	matrix male_total_cntrl=male_weekly_cntrl[1,1]+male_monthly_cntrl[1,1]+male_never_cntrl[1,1]
	
	matrix p_male_weekly_int= male_weekly_int[1,1]/ male_total_int[1,1]
	matrix p_male_monthly_int=male_monthly_int[1,1]/male_total_int[1,1]
	matri  p_male_never_int=  male_never_int[1,1]/  male_total_int[1,1]

	matrix  p_male_weekly_cntrl= male_weekly_cntrl[1,1]/ male_total_cntrl[1,1]
	matrix  p_male_monthly_cntrl=male_monthly_cntrl[1,1]/male_total_cntrl[1,1]
	matrix  p_male_never_cntrl=  male_never_cntrl[1,1]/  male_total_cntrl[1,1]
	
* PutExcel

	* Weekly
	putexcel `rowlabel'`R'="Weekly"
	putexcel `women_intervention'`R'=matrix(p_female_weekly_int)
	putexcel `women_control'`R'=     matrix(p_female_weekly_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_weekly_int)
	putexcel `men_control'`R'=     matrix(p_male_weekly_cntrl)
		local R=`R'+1
	
	* Monthly
	putexcel `rowlabel'`R'="At least once a year"
	putexcel `women_intervention'`R'=matrix(p_female_monthly_int)
	putexcel `women_control'`R'=     matrix(p_female_monthly_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_monthly_int)
	putexcel `men_control'`R'=     matrix(p_male_monthly_cntrl)
		local R=`R'+1
		
	* Never
	putexcel `rowlabel'`R'="Never"
	putexcel `women_intervention'`R'=matrix(p_female_never_int)
	putexcel `women_control'`R'=     matrix(p_female_never_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_never_int)
	putexcel `men_control'`R'=     matrix(p_male_never_cntrl)
		local R=`R'+1

** Watch TV
*******************************************

* Title
putexcel `rowlabel'`R'="Frequency of watching TV"
local R=`R'+1

* Combine 901 values
gen me_902_cat=me_902
recode me_902_cat (1 2 3=1) (4=2) (5=3)

* Percents
tab me_902_cat control if female==1, matcell(table)
	matrix female_weekly_int= table[1,1]
	matrix female_monthly_int=table[2,1]
	matrix female_never_int=  table[3,1]
	matrix female_total_int=female_weekly_int[1,1]+female_monthly_int[1,1]+female_never_int[1,1]

	matrix female_weekly_cntrl =table[1,2]
	matrix female_monthly_cntrl=table[2,2]
	matrix female_never_cntrl  =table[3,2]
	matrix female_total_cntrl=female_weekly_cntrl[1,1]+female_monthly_cntrl[1,1]+female_never_cntrl[1,1]
	
	matrix  p_female_weekly_int= female_weekly_int[1,1]/female_total_int[1,1]
	matrix p_female_monthly_int=female_monthly_int[1,1]/female_total_int[1,1]
	matri    p_female_never_int=  female_never_int[1,1]/female_total_int[1,1]

	matrix   p_female_weekly_cntrl= female_weekly_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_monthly_cntrl=female_monthly_cntrl[1,1]/female_total_cntrl[1,1]
	matrix    p_female_never_cntrl=  female_never_cntrl[1,1]/female_total_cntrl[1,1]

tab me_902_cat control if female==0, matcell(table)
	matrix male_weekly_int= table[1,1]
	matrix male_monthly_int=table[2,1]
	matrix male_never_int=  table[3,1]
	matrix male_total_int=male_weekly_int[1,1]+male_monthly_int[1,1]+male_never_int[1,1]

	matrix male_weekly_cntrl =table[1,2]
	matrix male_monthly_cntrl=table[2,2]
	matrix male_never_cntrl  =table[3,2]
	matrix male_total_cntrl=male_weekly_cntrl[1,1]+male_monthly_cntrl[1,1]+male_never_cntrl[1,1]
	
	matrix p_male_weekly_int= male_weekly_int[1,1]/ male_total_int[1,1]
	matrix p_male_monthly_int=male_monthly_int[1,1]/male_total_int[1,1]
	matri  p_male_never_int=  male_never_int[1,1]/  male_total_int[1,1]

	matrix  p_male_weekly_cntrl= male_weekly_cntrl[1,1]/ male_total_cntrl[1,1]
	matrix  p_male_monthly_cntrl=male_monthly_cntrl[1,1]/male_total_cntrl[1,1]
	matrix  p_male_never_cntrl=  male_never_cntrl[1,1]/  male_total_cntrl[1,1]
	
* PutExcel

	* Weekly
	putexcel `rowlabel'`R'="Weekly"
	putexcel `women_intervention'`R'=matrix(p_female_weekly_int)
	putexcel `women_control'`R'=     matrix(p_female_weekly_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_weekly_int)
	putexcel `men_control'`R'=     matrix(p_male_weekly_cntrl)
		local R=`R'+1
	
	* Monthly
	putexcel `rowlabel'`R'="At least once a year"
	putexcel `women_intervention'`R'=matrix(p_female_monthly_int)
	putexcel `women_control'`R'=     matrix(p_female_monthly_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_monthly_int)
	putexcel `men_control'`R'=     matrix(p_male_monthly_cntrl)
		local R=`R'+1
		
	* Never
	putexcel `rowlabel'`R'="Never"
	putexcel `women_intervention'`R'=matrix(p_female_never_int)
	putexcel `women_control'`R'=     matrix(p_female_never_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_never_int)
	putexcel `men_control'`R'=     matrix(p_male_never_cntrl)
		local R=`R'+1	
		
		
** Cellphone Ownership
*******************************************

* Title
putexcel `rowlabel'`R'="Owns at least one (1) cellphone"
local R=`R'+1

* Combine 901 values
gen me_903_cat=me_903
recode me_903_cat (0=0) (1/9=1)

* Percents
tab me_903_cat control if female==1, matcell(table)
	matrix  female_none_int= table[1,1]
	matrix   female_one_int=table[2,1]
	matrix female_total_int=female_none_int[1,1]+female_one_int[1,1]

	matrix  female_none_cntrl= table[1,2]
	matrix   female_one_cntrl=table[2,2]
	matrix female_total_cntrl=female_none_cntrl[1,1]+female_one_cntrl[1,1]
	
	matrix p_female_none_int= female_none_int[1,1]/female_total_int[1,1]
	matrix  p_female_one_int=female_one_int[1,1]/female_total_int[1,1]

	matrix p_female_none_cntrl=female_none_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_one_cntrl= female_one_cntrl[1,1]/female_total_cntrl[1,1]

tab me_903_cat control if female==0, matcell(table)
	matrix male_none_int= table[1,1]
	matrix male_one_int=table[2,1]
	matrix male_total_int=male_none_int[1,1]+male_one_int[1,1]

	matrix male_none_cntrl= table[1,2]
	matrix male_one_cntrl=table[2,2]
	matrix male_total_cntrl=male_none_cntrl[1,1]+male_one_cntrl[1,1]
	
	matrix p_male_none_int=male_none_int[1,1]/male_total_int[1,1]
	matrix p_male_one_int= male_one_int[1,1]/ male_total_int[1,1]

	matrix p_male_none_cntrl=male_none_cntrl[1,1]/male_total_cntrl[1,1]
	matrix p_male_one_cntrl= male_one_cntrl[1,1]/ male_total_cntrl[1,1]
	
* PutExcel

	* None
	putexcel `rowlabel'`R'="None"
	putexcel `women_intervention'`R'=matrix(p_female_none_int)
	putexcel `women_control'`R'=     matrix(p_female_none_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_none_int)
	putexcel `men_control'`R'=     matrix(p_male_none_cntrl)
		local R=`R'+1
	
	* At least one
	putexcel `rowlabel'`R'="At least one (1)"
	putexcel `women_intervention'`R'=matrix(p_female_one_int)
	putexcel `women_control'`R'=     matrix(p_female_one_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_one_int)
	putexcel `men_control'`R'=     matrix(p_male_one_cntrl)
		local R=`R'+1

** Smartphone Ownership
*******************************************		

* Title
putexcel `rowlabel'`R'="Owns at least one (1) smartphone"
local R=`R'+1

* Combine 901 values
gen me_904_3_cat_v2=me_904_3
recode me_904_3_cat_v2 (0=0) (1/10=1)

* Percents
tab me_904_3_cat_v2 control if female==1, matcell(table)
	matrix  female_none_int= table[1,1]
	matrix   female_one_int=table[2,1]
	matrix female_total_int=female_none_int[1,1]+female_one_int[1,1]

	matrix  female_none_cntrl= table[1,2]
	matrix   female_one_cntrl=table[2,2]
	matrix female_total_cntrl=female_none_cntrl[1,1]+female_one_cntrl[1,1]
	
	matrix p_female_none_int= female_none_int[1,1]/female_total_int[1,1]
	matrix  p_female_one_int=female_one_int[1,1]/female_total_int[1,1]

	matrix p_female_none_cntrl=female_none_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_one_cntrl= female_one_cntrl[1,1]/female_total_cntrl[1,1]

tab me_904_3_cat_v2 control if female==0, matcell(table)
	matrix male_none_int= table[1,1]
	matrix male_one_int=table[2,1]
	matrix male_total_int=male_none_int[1,1]+male_one_int[1,1]

	matrix male_none_cntrl= table[1,2]
	matrix male_one_cntrl=table[2,2]
	matrix male_total_cntrl=male_none_cntrl[1,1]+male_one_cntrl[1,1]
	
	matrix p_male_none_int=male_none_int[1,1]/male_total_int[1,1]
	matrix p_male_one_int= male_one_int[1,1]/ male_total_int[1,1]

	matrix p_male_none_cntrl=male_none_cntrl[1,1]/male_total_cntrl[1,1]
	matrix p_male_one_cntrl= male_one_cntrl[1,1]/ male_total_cntrl[1,1]
	
* PutExcel

	* None
	putexcel `rowlabel'`R'="None"
	putexcel `women_intervention'`R'=matrix(p_female_none_int)
	putexcel `women_control'`R'=     matrix(p_female_none_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_none_int)
	putexcel `men_control'`R'=     matrix(p_male_none_cntrl)
		local R=`R'+1
	
	* At least one
	putexcel `rowlabel'`R'="At least one (1)"
	putexcel `women_intervention'`R'=matrix(p_female_one_int)
	putexcel `women_control'`R'=     matrix(p_female_one_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_one_int)
	putexcel `men_control'`R'=     matrix(p_male_one_cntrl)
		local R=`R'+1

	
** Women & Girls Access to Cellphones
*******************************************		

* Title
putexcel `rowlabel'`R'="Women & Girls need to ask permission to use cellphones"
local R=`R'+1


* Percents
tab me_906 control if female==1, matcell(table)
	matrix  female_no_int= table[1,1]
	matrix   female_yes_int=table[2,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix  female_no_cntrl= table[1,2]
	matrix   female_yes_cntrl=table[2,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix p_female_no_int= female_no_int[1,1]/female_total_int[1,1]
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix p_female_no_cntrl=female_no_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_906 control if female==0, matcell(table)
	matrix male_no_int= table[1,1]
	matrix male_yes_int=table[2,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix male_no_cntrl= table[1,2]
	matrix male_yes_cntrl=table[2,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_no_int= male_no_int[1,1]/ male_total_int[1,1]
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_no_cntrl= male_no_cntrl[1,1]/ male_total_cntrl[1,1]
	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* No
	putexcel `rowlabel'`R'="No"
	putexcel `women_intervention'`R'=matrix(p_female_no_int)
	putexcel `women_control'`R'=     matrix(p_female_no_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_no_int)
	putexcel `men_control'`R'=     matrix(p_male_no_cntrl)
		local R=`R'+1
	
	* Yes
	putexcel `rowlabel'`R'="Yes"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

** Reason for cellphone use
*******************************************		

* Title
putexcel `rowlabel'`R'="Reason for cellphone use"
local R=`R'+1

********** Send and Receive SMS **********
* Percents
tab me_908_2 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_908_2 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Send and Receive SMS"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

********** View media/listen to radio **********
* Percents
tab me_908_3 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_908_3 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="View media/listen to radio"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

** Ever Used Social Media
*******************************************	

* Percents
tab me_909 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_909 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Has ever used social media"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1


** Been to a health facility in the last 6 months
*************************************************	

* Title
putexcel `rowlabel'`R'="Been to a health facility in the last 6 months"
local R=`R'+1

* Percents
tab me_914 control if female==1, matcell(table)
	matrix  female_none_int= table[1,1]
	matrix   female_one_int=table[2,1]
	matrix female_total_int=female_none_int[1,1]+female_one_int[1,1]

	matrix  female_none_cntrl= table[1,2]
	matrix   female_one_cntrl=table[2,2]
	matrix female_total_cntrl=female_none_cntrl[1,1]+female_one_cntrl[1,1]
	
	matrix p_female_none_int= female_none_int[1,1]/female_total_int[1,1]
	matrix  p_female_one_int=female_one_int[1,1]/female_total_int[1,1]

	matrix p_female_none_cntrl=female_none_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_one_cntrl= female_one_cntrl[1,1]/female_total_cntrl[1,1]

tab me_914 control if female==0, matcell(table)
	matrix male_none_int= table[1,1]
	matrix male_one_int=table[2,1]
	matrix male_total_int=male_none_int[1,1]+male_one_int[1,1]

	matrix male_none_cntrl= table[1,2]
	matrix male_one_cntrl=table[2,2]
	matrix male_total_cntrl=male_none_cntrl[1,1]+male_one_cntrl[1,1]
	
	matrix p_male_none_int=male_none_int[1,1]/male_total_int[1,1]
	matrix p_male_one_int= male_one_int[1,1]/ male_total_int[1,1]

	matrix p_male_none_cntrl=male_none_cntrl[1,1]/male_total_cntrl[1,1]
	matrix p_male_one_cntrl= male_one_cntrl[1,1]/ male_total_cntrl[1,1]
	
* PutExcel

	* None
	putexcel `rowlabel'`R'="Not Visited"
	putexcel `women_intervention'`R'=matrix(p_female_none_int)
	putexcel `women_control'`R'=     matrix(p_female_none_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_none_int)
	putexcel `men_control'`R'=     matrix(p_male_none_cntrl)
		local R=`R'+1
	
	* At least one
	putexcel `rowlabel'`R'="Visited"
	putexcel `women_intervention'`R'=matrix(p_female_one_int)
	putexcel `women_control'`R'=     matrix(p_female_one_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_one_int)
	putexcel `men_control'`R'=     matrix(p_male_one_cntrl)
		local R=`R'+1
		
** Visisted by CHA/CHV/CHSS in the last 6 months
*************************************************	

* Title
putexcel `rowlabel'`R'="Visisted by CHA/CHV/CHSS in the last 6 months"
local R=`R'+1

* Combine 901 values
gen me_911_cat_v2=me_911_cat
recode me_911_cat_v2 (0=0) (1/2=1)

* Percents
tab me_911_cat_v2 control if female==1, matcell(table)
	matrix  female_none_int= table[1,1]
	matrix   female_one_int=table[2,1]
	matrix female_total_int=female_none_int[1,1]+female_one_int[1,1]

	matrix  female_none_cntrl= table[1,2]
	matrix   female_one_cntrl=table[2,2]
	matrix female_total_cntrl=female_none_cntrl[1,1]+female_one_cntrl[1,1]
	
	matrix p_female_none_int= female_none_int[1,1]/female_total_int[1,1]
	matrix  p_female_one_int=female_one_int[1,1]/female_total_int[1,1]

	matrix p_female_none_cntrl=female_none_cntrl[1,1]/female_total_cntrl[1,1]
	matrix  p_female_one_cntrl= female_one_cntrl[1,1]/female_total_cntrl[1,1]

tab me_911_cat_v2 control if female==0, matcell(table)
	matrix male_none_int= table[1,1]
	matrix male_one_int=table[2,1]
	matrix male_total_int=male_none_int[1,1]+male_one_int[1,1]

	matrix male_none_cntrl= table[1,2]
	matrix male_one_cntrl=table[2,2]
	matrix male_total_cntrl=male_none_cntrl[1,1]+male_one_cntrl[1,1]
	
	matrix p_male_none_int=male_none_int[1,1]/male_total_int[1,1]
	matrix p_male_one_int= male_one_int[1,1]/ male_total_int[1,1]

	matrix p_male_none_cntrl=male_none_cntrl[1,1]/male_total_cntrl[1,1]
	matrix p_male_one_cntrl= male_one_cntrl[1,1]/ male_total_cntrl[1,1]
	
* PutExcel

	* None
	putexcel `rowlabel'`R'="Not Visited"
	putexcel `women_intervention'`R'=matrix(p_female_none_int)
	putexcel `women_control'`R'=     matrix(p_female_none_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_none_int)
	putexcel `men_control'`R'=     matrix(p_male_none_cntrl)
		local R=`R'+1
	
	* At least one
	putexcel `rowlabel'`R'="Visited"
	putexcel `women_intervention'`R'=matrix(p_female_one_int)
	putexcel `women_control'`R'=     matrix(p_female_one_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_one_int)
	putexcel `men_control'`R'=     matrix(p_male_one_cntrl)
		local R=`R'+1


********************************************************
*** Table 9.2: Satisfaction with Health Providers ***
********************************************************

putexcel set "$putexcel_set", modify sheet("Table9.2")
local R $R

* Titles
putexcel `women_collab'`R'="Women"
putexcel `men_collab'`R'="Men"
	local R=`R'+1
	
	
* Counts
sum one if control==0 & female==1
	local n_Intervention_female=r(N)
sum one if control==1 & female==1
	local n_Control_female=r(N)
sum one if control==0 & female==0
	local n_Intervention_male=r(N)
sum one if control==1 & female==0
	local n_Control_male=r(N)
	
	** Tiles
	local celltext Intervention N=`n_Intervention_female'
		putexcel `women_intervention'`R'="`celltext'"
	local celltext Control N=`n_Control_female' 
		putexcel `women_control'`R'="`celltext'"
	local celltext Intervention N=`n_Intervention_male' 
		putexcel `men_intervention'`R'="`celltext'"
	local celltext Control N=`n_Control_male' 
		putexcel `men_control'`R'="`celltext'"

local R=`R'+1	


** Satisfaction with CHW
*************************************************

* Recode Variable
recode me_913 (-98=.)
gen me_913_cat_v2=me_913/10
recode me_913_cat_v2 (0/6=0) (7/10=1)

* Percents
tab me_913_cat_v2 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_913_cat_v2 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Very satisfied with CHA/CHV Services"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

** Satisfaction with CHSS
*************************************************

* Recode Variable
recode me_915 (-98=.)
gen me_915_cat_v2=me_915/10
recode me_915_cat_v2 (0/6=0) (7/10=1)

* Percents
tab me_915_cat_v2 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_915_cat_v2 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Very satisfied with CHSS services"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

** Satisfaction with CHSS
*************************************************

* Recode Variable
recode me_916 (-98=.)
gen me_916_cat_v2=me_916/10
recode me_916_cat_v2 (0/6=0) (7/10=1)

* Percents
tab me_916_cat_v2 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_916_cat_v2 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Very satisfied with Doctor or Nurse's Services"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

** Has Heard about `Share It, Act It'
*************************************************

* Recode Variable
recode me_918 (-98=.)

* Percents
tab me_918 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix    female_no_int=table[1,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,2]
	matrix    female_no_cntrl=table[1,2]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_918 control if female==0, matcell(table)
	matrix   male_yes_int=table[2,1]
	matrix    male_no_int=table[1,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix  male_yes_cntrl=table[2,2]
	matrix   male_no_cntrl=table[1,2]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Has heard of Share It, Act It"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1

** Has Heard about the `Healthy Life' Campaign
*************************************************

* Recode Variable
recode me_919 (-98=.)

* Percents
tab me_919 control if female==1, matcell(table)
	matrix   female_yes_int=table[2,1]
	matrix female_total_int=female_no_int[1,1]+female_yes_int[1,1]

	matrix   female_yes_cntrl=table[2,1]
	matrix female_total_cntrl=female_no_cntrl[1,1]+female_yes_cntrl[1,1]
	
	matrix  p_female_yes_int=female_yes_int[1,1]/female_total_int[1,1]

	matrix  p_female_yes_cntrl= female_yes_cntrl[1,1]/female_total_cntrl[1,1]

tab me_919 control if female==0, matcell(table)
	matrix male_yes_int=table[2,1]
	matrix male_total_int=male_no_int[1,1]+male_yes_int[1,1]

	matrix male_yes_cntrl=table[2,1]
	matrix male_total_cntrl=male_no_cntrl[1,1]+male_yes_cntrl[1,1]
	
	matrix p_male_yes_int=male_yes_int[1,1]/male_total_int[1,1]

	matrix p_male_yes_cntrl=male_yes_cntrl[1,1]/male_total_cntrl[1,1]
	
* PutExcel

	* Yes
	putexcel `rowlabel'`R'="Has heard about the Healthy Life Campaign"
	putexcel `women_intervention'`R'=matrix(p_female_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_yes_cntrl)
	putexcel `men_intervention'`R'=matrix(p_male_yes_int)
	putexcel `men_control'`R'=     matrix(p_male_yes_cntrl)
		local R=`R'+1		


********************************************************
*** Table 9.3: MESSAGE EXPOSURE - SOURCES ***
********************************************************

putexcel set "$putexcel_set", modify sheet("Table9.3")
local R $R

* Titles
putexcel `women_collab'`R'="Women"
putexcel `men_collab'`R'="Men"
	local R=`R'+1
	
	
* Counts
sum one if control==0 & female==1
	local n_Intervention_female=r(N)
sum one if control==1 & female==1
	local n_Control_female=r(N)
sum one if control==0 & female==0
	local n_Intervention_male=r(N)
sum one if control==1 & female==0
	local n_Control_male=r(N)
	
	** Tiles
	local celltext Intervention N=`n_Intervention_female'
		putexcel `women_intervention'`R'="`celltext'"
	local celltext Control N=`n_Control_female' 
		putexcel `women_control'`R'="`celltext'"
	local celltext Intervention N=`n_Intervention_male' 
		putexcel `men_intervention'`R'="`celltext'"
	local celltext Control N=`n_Control_male' 
		putexcel `men_control'`R'="`celltext'"

local R=`R'+1	

** Generate Variables
*************************************************

foreach num in 1 2 3 5 6 7 8 9 {

	* Facility-Based
	gen me_920_`num'_FBHW=0
		replace   me_920_`num'_FBHW=1 if me_920_`num'_1==1 | me_920_`num'_2==1 | me_920_`num'_4
		label var me_920_`num'_FBHW "Information from facility-based health worker"
		
	* Community Health Worker
	gen me_920_`num'_CHW=0
		replace   me_920_`num'_CHW=1 if me_920_`num'_18==1
		label var me_920_`num'_CHW "Information from CHA/CHV/CHSS"	
		
	* Friends/Relatives
	gen me_920_`num'_FR=0
		replace   me_920_`num'_FR=1 if me_920_`num'_15==1 | me_920_`num'_16==1 
		label var me_920_`num'_FR "Information from friends or relatives"
		
	* Radio
	gen me_920_`num'_R=0
		replace   me_920_`num'_R=1 if me_920_`num'_12==1 
		label var me_920_`num'_R "Information from the radio"

	* Community Source (Church/Mosque, CBO, Women's Org)s
	gen me_920_`num'_CS=0
		capture replace me_920_`num'_CS=1 if me_920_`num'_11==1 
		replace me_920_`num'_CS=1 if me_920_`num'_10==1 | me_920_`num'_9==1
		label var me_920_`num'_CS "Information from Community Source (Church/Mosque, CBO, Women's Org)s"		
		
	}

** Generate Macros
*************************************************

foreach info in 1 2 3 5 6 7 8 9 {
	foreach source in FBHW CHW FR R CS {
	
	* Percents
	tab me_920_`info'_`source' control if female==1, matcell(table)
		matrix   female_`info'_`source'_yes_int=table[2,1]
		matrix    female_`info'_`source'_no_int=table[1,1]
		matrix female_`info'_`source'_total_int=female_`info'_`source'_no_int[1,1]+female_`info'_`source'_yes_int[1,1]
	
		matrix   female_`info'_`source'_yes_cntrl=table[2,2]
		matrix    female_`info'_`source'_no_cntrl=table[1,2]
		matrix female_`info'_`source'_total_cntrl=female_`info'_`source'_no_cntrl[1,1]+female_`info'_`source'_yes_cntrl[1,1]
		
		matrix  p_female_`info'_`source'_yes_int=female_`info'_`source'_yes_int[1,1]/female_`info'_`source'_total_int[1,1]
	
		matrix  p_female_`info'_`source'_yes_cntrl= female_`info'_`source'_yes_cntrl[1,1]/female_`info'_`source'_total_cntrl[1,1]
	
	tab me_920_`info'_`source' control if female==0, matcell(table)
		matrix   male_`info'_`source'_yes_int=table[2,1]
		matrix    male_`info'_`source'_no_int=table[1,1]
		matrix male_`info'_`source'_total_int=male_`info'_`source'_no_int[1,1]+male_`info'_`source'_yes_int[1,1]
	
		matrix  male_`info'_`source'_yes_cntrl=table[2,2]
		matrix   male_`info'_`source'_no_cntrl=table[1,2]
		matrix male_`info'_`source'_total_cntrl=male_`info'_`source'_no_cntrl[1,1]+male_`info'_`source'_yes_cntrl[1,1]
		
		matrix p_male_`info'_`source'_yes_int=male_`info'_`source'_yes_int[1,1]/male_`info'_`source'_total_int[1,1]
	
		matrix p_male_`info'_`source'_yes_cntrl=male_`info'_`source'_yes_cntrl[1,1]/male_`info'_`source'_total_cntrl[1,1]
		
		}
	}
	
** Antenatal Care
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on antenatal care from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_1_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_1_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_1_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_1_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_1_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_1_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_1_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_1_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_1_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_1_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_1_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_1_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_1_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_1_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_1_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_1_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_1_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_1_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_1_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_1_CS_yes_cntrl)
		local R=`R'+1	

** Postnatal Care
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on postnatal care from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_2_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_2_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_2_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_2_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_2_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_2_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_2_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_2_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_2_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_2_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_2_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_2_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_2_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_2_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_2_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_2_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_2_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_2_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_2_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_2_CS_yes_cntrl)
		local R=`R'+1	

** Respectful Maternity Care
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on respectful maternity care from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_3_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_3_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_3_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_3_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_3_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_3_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_3_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_3_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_3_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_3_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_3_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_3_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_3_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_3_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_3_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_3_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_3_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_3_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_3_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_3_CS_yes_cntrl)
		local R=`R'+1	

** Nutrition
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on nutrition from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_5_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_5_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_5_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_5_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_5_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_5_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_5_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_5_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_5_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_5_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_5_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_5_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_5_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_5_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_5_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_5_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_5_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_5_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_5_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_5_CS_yes_cntrl)
		local R=`R'+1	

** WASH
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on WASH from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_6_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_6_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_6_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_6_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_6_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_6_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_6_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_6_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_6_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_6_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_6_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_6_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_6_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_6_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_6_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_6_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_6_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_6_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_6_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_6_CS_yes_cntrl)
		local R=`R'+1	

** COVID
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on COVID from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_7_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_7_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_7_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_7_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_7_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_7_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_7_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_7_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_7_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_7_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_7_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_7_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_7_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_7_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_7_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_7_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_7_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_7_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_7_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_7_CS_yes_cntrl)
		local R=`R'+1			
		
		
** Share it, Act it
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on 'Share It, Act It' from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_8_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_8_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_8_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_8_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_8_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_8_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_8_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_8_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_8_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_8_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_8_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_8_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_8_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_8_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_8_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_8_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_8_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_8_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_8_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_8_CS_yes_cntrl)
		local R=`R'+1			
		
		
** Health Life Campaign
*************************************************

* Title
putexcel `rowlabel'`R'="Received information on the 'Healthy Life Campaign' from source"
local R=`R'+1
	
* PutExcel

	* Facility Based Health Worker
	putexcel `rowlabel'`R'="Facility Based Health Worker (Doctor, Nurse, Midwife)"
	putexcel `women_intervention'`R'=matrix(p_female_9_FBHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_9_FBHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_9_FBHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_9_FBHW_yes_cntrl)
		local R=`R'+1	

	* Community Health Worker
	putexcel `rowlabel'`R'="Community Health Worker (CHA, CHV, CHSS)"
	putexcel `women_intervention'`R'=matrix(p_female_9_CHW_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_9_CHW_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_9_CHW_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_9_CHW_yes_cntrl)
		local R=`R'+1	

	* Friend or Relative
	putexcel `rowlabel'`R'="Friend or Relative"
	putexcel `women_intervention'`R'=matrix(p_female_9_FR_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_9_FR_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_9_FR_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_9_FR_yes_cntrl)
		local R=`R'+1	

	* Radio
	putexcel `rowlabel'`R'="Radio"
	putexcel `women_intervention'`R'=matrix(p_female_9_R_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_9_R_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_9_R_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_9_R_yes_cntrl)
		local R=`R'+1	

	* Community Source (Church/Mosque, CBO, Women's Org)
	putexcel `rowlabel'`R'="Community Source (Church/Mosque, CBO, Women's Org)"
	putexcel `women_intervention'`R'=matrix(p_female_9_CS_yes_int)
	putexcel `women_control'`R'=     matrix(p_female_9_CS_yes_cntrl)
	putexcel `men_intervention'`R'=    matrix(p_male_9_CS_yes_int)
	putexcel `men_control'`R'=         matrix(p_male_9_CS_yes_cntrl)
		local R=`R'+1			
		
		
