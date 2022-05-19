*install if you do not have (remove * below to run installation)
ssc install fre //package that shows values and labels together for tabulation

/////GEM SCALE - REDONE/////

	drop gem_*
	
	**Partner Violence (PV) sub-scale (ls_506_1 to ls_506_6)
	gen gem_pv_score = (ls_506_1 + ls_506_2 + ls_506_3 + ls_506_4 + ls_506_5 + ls_506_6)
	ta gem_pv_score
	
	xtile gem_pv_index = gem_pv_score, nq(3) //partner violence subscale - tertiles
	ta gem_pv_index //distribution looks good
	ta gem_pv_score gem_pv_index  //cross-reference looks good too
	
	label variable gem_pv_index "Partner Violence subscale"
	
	
	**Sexual Relationships (SR) sub-scale (ls_506_7 to ls_506_14)
	gen gem_sr_score = (ls_506_7 + ls_506_8 + ls_506_9 + ls_506_10 + ls_506_11 + ls_506_12 + ls_506_13 + ls_506_14)
	ta gem_sr_score
	
	xtile gem_sr_index = gem_sr_score, nq(3) //sexual relationships subscale - tertiles
	ta gem_sr_index // distribution looks good
	ta gem_sr_score gem_sr_index //good
	
	label variable gem_sr_index "Sexual Relationships subscale"
	
	
	**Reproductive Health (RH) sub-scale (ls_506_15 to ls_506_19)
	gen gem_rh_score = (ls_506_15 + ls_506_16 + ls_506_17 + ls_506_18 + ls_506_19)
	ta gem_rh_score
	
	xtile gem_rh_index=gem_rh_score, nq(3) //RH subsacle - tertiles
	ta gem_rh_index //looks good
	ta gem_rh_score gem_rh_index //overlap is correct
	
	label variable gem_rh_index "Reproductive Health subscale"
	
	
	**Domestic Chores and Daily Life (DCDL) sub-scale (ls_506_20 to ls_506_24)
	gen gem_dcdl_score = (ls_506_20 + ls_506_21 + ls_506_22 + ls_506_23 + ls_506_24)
	ta gem_dcdl_score
	
	xtile gem_dcdl_index=gem_dcdl_score, nq(3)
	ta gem_dcdl_index
	ta gem_dcdl_score gem_dcdl_index
	
	label variable gem_dcdl_index "Domestic Chores and Daily Life subscale"
	
	label define tertile 1 "Low Gender Inequity" 2 "Moderate Gender Inequity" 3 "High Gender Inequity"
	label values gem_pv_index gem_dcdl_index gem_rh_index gem_sr_index tertile
	
	*check labels are good for all
	foreach i of varlist gem_*index{
		ta `i'
	}
	
	
/////COUPLE COMMUNICATION - REDONE/////

*total should be married and cohabitating total
fre marital_status

di 467+619 //total should 1086 in this example



*look at values
egen couple_communication_total=rowtotal(ls_501_1 ls_501_3 ls_501_5 ls_501_6 ls_501_7 ls_501_10)
replace couple_communication_total=. if marital_status==3 | marital_status==. // get rid of couple communication variable for single and missing marital status

ta couple_comm* //1086 count, so this looks good

xtile couple_communication_index=couple_communication_total, nq(3)

label define comm 1 "Low Couple Communication" 2 "Moderate Couple Communication" 3 "Frequent Couple Communication"
label values couple_communication_index comm
ta couple_communication_t* couple_communication_i*  //looks good
label variable couple_communication_index "couple communication composite index"

ta couple_communication_index
