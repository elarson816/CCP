local var1 $var1
local var2 $var2
local first $first
local last $last

** Top 5
foreach list in $list {
	
	* Recode for counting
	recode  `var1'_`list' 0 = .
	
	* Count of people who heard of method
	egen n_`var1'_`list'=count(`var1'_`list')
		
	* Percent of people who heard of method
	gen p_`var2'_`list'=n_`var1'_`list'/total_n

	* Recode back
	recode `var1'_`list' . = 0
	
	}
	
	* Top
	order *, sequential
	egen top5_1_`var1'=rowmax(p_`var2'_`first'-p_`var2'_`last')
	
	gen top5_1_`var1'_item=""
	
		foreach list in $list {
			replace top5_1_`var1'_item="`list'" if top5_1_`var1'==p_`var2'_`list'
			replace p_`var2'_`list'=0 if top5_1_`var1'==p_`var2'_`list'
		}
			
	label var top5_1_`var1' "Most common % (`var1')"
	label var top5_1_`var1'_item "Most common item (`var1')"
	
	* Second 
	order *, sequential
	egen top5_2_`var1'=rowmax(p_`var2'_`first'-p_`var2'_`last')
	
	gen top5_2_`var1'_item=""
	
		foreach list in $list {		
			replace top5_2_`var1'_item="`list'" if top5_2_`var1'==p_`var2'_`list'
			replace p_`var2'_`list'=0 if top5_2_`var1'==p_`var2'_`list'
		}

	label var top5_2_`var1' "Second most common % (`var1')"
	label var top5_2_`var1'_item "Second most common item (`var1')"
	
	* Third
	order *, sequential
	egen top5_3_`var1'=rowmax(p_`var2'_`first'-p_`var2'_`last')

	gen top5_3_`var1'_item=""	
	
		foreach list in $list {
			replace top5_3_`var1'_item="`list'" if top5_3_`var1'==p_`var2'_`list'
			replace p_`var2'_`list'=0 if top5_3_`var1'==p_`var2'_`list'
		}

	label var top5_3_`var1' "Third most common % (`var1')"
	label var top5_3_`var1'_item "Third most common item (`var1')"
			
	* Fourth
	order *, sequential
	egen top5_4_`var1'=rowmax(p_`var2'_`first'-p_`var2'_`last')

	gen top5_4_`var1'_item=""	
	
		foreach list in $list {
			replace top5_4_`var1'_item="`list'" if top5_4_`var1'==p_`var2'_`list'
			replace p_`var2'_`list'=0 if top5_4_`var1'==p_`var2'_`list'
		}

	label var top5_4_`var1' "Fourth most common % (`var1')"
	label var top5_4_`var1'_item "Fourth most common item (`var1')"
	
	* Fifth
	order *, sequential
	egen top5_5_`var1'=rowmax(p_`var2'_`first'-p_`var2'_`last')	
	
	gen top5_5_`var1'_item=""
	
		foreach list in $list {
			replace top5_5_`var1'_item="`list'" if top5_5_`var1'==p_`var2'_`list'
			replace p_`var2'_`list'=0 if top5_5_`var1'==p_`var2'_`list'
		}
				
	label var top5_5_`var1' "Fifth most common % (`var1')"
	label var top5_5_`var1'_item "Fifth most common item (`var1')"
