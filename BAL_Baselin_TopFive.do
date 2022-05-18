local var1 $var1
local var2 $var2
local first $first
local last $last

** Top 5
foreach list in $list {
	
	* Recode for counting
	recode  `var1'_`list' 0 = .
	
	* Count of people who heard of method in region
	egen n_`var1'_`list'_reg=count(`var1'_`list'), by(county_id)
	
	
	* Percent of people in the region who heard of method
	gen p_`var2'_`list'_reg=n_`var1'_`list'_reg/N_region
	
	* Recode back
	recode `var1'_`list' . = 0
	
	}
	
	* Top
	order *, sequential
	egen top5_1_`var1'=rowmax(p_`var2'_`first'_reg-p_`var2'_`last'_reg)
	
	gen top5_1_`var1'_item=""
	
		foreach list in $list {
			
			* Change top percent to 0 and create variable indicating the top method
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_1_`var1' & county_id==22
				replace top5_1_`var1'_item="`list'" if county_id==22 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_1_`var1'_item=="`list'" & county_id==22
			
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_1_`var1' & county_id==23
				replace top5_1_`var1'_item="`list'" if county_id==23 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_1_`var1'_item=="`list'" & county_id==23
			
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_1_`var1' & county_id==24
				replace top5_1_`var1'_item="`list'" if county_id==24 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_1_`var1'_item=="`list'" & county_id==24
			}
	
	* Second 
	order *, sequential
	egen top5_2_`var1'=rowmax(p_`var2'_`first'_reg-p_`var2'_`last'_reg)
	
	gen top5_2_`var1'_item=""
	
		foreach list in $list {
			
			* Change top percent to 0 and create variable indicating the top method
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_2_`var1' & county_id==22
				replace top5_2_`var1'_item="`list'" if county_id==22 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_2_`var1'_item=="`list'" & county_id==22
			
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_2_`var1' & county_id==23
				replace top5_2_`var1'_item="`list'" if county_id==23 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_2_`var1'_item=="`list'" & county_id==23
			
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_2_`var1' & county_id==24
				replace top5_2_`var1'_item="`list'" if county_id==24 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_2_`var1'_item=="`list'" & county_id==24
			}
	
	* Third
	order *, sequential
	egen top5_3_`var1'=rowmax(p_`var2'_`first'_reg-p_`var2'_`last'_reg)

	gen top5_3_`var1'_item=""	
	
		foreach list in $list {
			
			* Change top percent to 0 and create variable indicating the top method
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_3_`var1' & county_id==22
				replace top5_3_`var1'_item="`list'" if county_id==22 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_3_`var1'_item=="`list'" & county_id==22
				
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_3_`var1' & county_id==23
				replace top5_3_`var1'_item="`list'" if county_id==23 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_3_`var1'_item=="`list'" & county_id==23
				
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_3_`var1' & county_id==24
				replace top5_3_`var1'_item="`list'" if county_id==24 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_3_`var1'_item=="`list'" & county_id==24
			}
			
	* Fourth
	order *, sequential
	egen top5_4_`var1'=rowmax(p_`var2'_`first'_reg-p_`var2'_`last'_reg)

	gen top5_4_`var1'_item=""	
	
		foreach list in $list {
			
			* Change top percent to 0 and create variable indicating the top method
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_4_`var1' & county_id==22
				replace top5_4_`var1'_item="`list'" if county_id==22 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_4_`var1'_item=="`list'" & county_id==22
				
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_4_`var1' & county_id==23
				replace top5_4_`var1'_item="`list'" if county_id==23 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_4_`var1'_item=="`list'" & county_id==23
				
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_4_`var1' & county_id==24
				replace top5_4_`var1'_item="`list'" if county_id==24 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_4_`var1'_item=="`list'" & county_id==24
			}
	
	* Fifth
	order *, sequential
	egen top5_5_`var1'=rowmax(p_`var2'_`first'_reg-p_`var2'_`last'_reg)	
	
	gen top5_5_`var1'_item=""
	
		foreach list in $list {
			
			* Change top percent to 0 and create variable indicating the top method
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_5_`var1' & county_id==22
				replace top5_5_`var1'_item="`list'" if county_id==22 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_5_`var1'_item=="`list'" & county_id==22
				
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_5_`var1' & county_id==23
				replace top5_5_`var1'_item="`list'" if county_id==23 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_4_`var1'_item=="`list'" & county_id==23
				
			replace p_`var2'_`list'_reg=0 if p_`var2'_`list'_reg==top5_5_`var1' & county_id==24
				replace top5_5_`var1'_item="`list'" if county_id==24 & p_`var2'_`list'_reg==0
				replace p_`var2'_`list'_reg=. if top5_4_`var1'_item=="`list'" & county_id==25
			}
