*********************************************************************
*** Table 9.1: INFORMATION SOURCES ***
*********************************************************************
	
* Title
putexcel `rowlabel'`R'="Top Five Maternal Health Information Sources (Total)"
	local R=`R'+1	

foreach n in 1 2 3 4 5 {
			
	** Item
	local top5_`n'_me_920_1_female=top5_`n'_me_920_1_1_item[1]
	local top5_`n'_me_920_1_male=  top5_`n'_me_920_1_0_item[1]
	
	** Percent
	sum    top5_`n'_me_920_1_1
	matrix top5_`n'_me_920_1_female=r(mean)
	
	sum    top5_`n'_me_920_1_0
	matrix top5_`n'_me_920_1_male=r(mean)
	}
	
	
* Put Excel
		
	** Items and Percents
	foreach n in 1 2 3 4 5 {
		putexcel `women_intervention_col1'`R'=     "`top5_`n'_me_920_1_female'"
		putexcel `women_intervention_col2'`R'=matrix(top5_`n'_me_920_1_female)
		putexcel `men_intervention_col1'`R'=     "`top5_`n'_me_920_1_male'"
		putexcel `men_intervention_col2'`R'=matrix(top5_`n'_me_920_1_male)
			local R=`R'+1
	}


