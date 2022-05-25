** ParentFile

clear
capture log close

/*** Breathrough Action Liberia - Adolescent Family Planning ***

* Input: BAL_baseline_femaleadol_cleaned_dataset_v3.dta 
		 BAL_baseline_male adolescent_cleaned_dataset_v2.dta
* Output: BAL Baseline Dummy Tables Adolescent 02.xls		 
*/

********************************************************
*** SET UP DO FILE ***
********************************************************

* Directory
cd "/Users/ealarson/Documents/CCP/BA_Adolsecents_Liberia"
global directory "/Users/ealarson/Documents/CCP/BA_Adolsecents_Liberia"

* Datasets
global male_data "1. Data/BAL_baseline_maleadolescent_cleaned_dataset_v4.dta"
global male_total_data "1. Data/BAL_baseline_maleadolescent_total_cleaned_dataset_v4.dta"
global female_ums_data "1. Data/BAL_baseline_femaleadol_umsexactive_cleaned_dataset_v4.dta"
global female_ums_total_data "1. Data/BAL_baseline_femaleadol_umsexactive_total_cleaned_dataset_v4.dta"
global female_iu_data "1. Data/BAL_baseline_femaleadol_inunion_cleaned_dataset_v4.dta.dta"
global female_iu_total_data "1. Data/BAL_baseline_femaleadol_inunion_total_cleaned_dataset_v4.dta.dta"

* Date 
local today=c(current_date)
local c_today= "`today'"
global date=subinstr("`c_today'", " ", "",.)

di %td_CY-N-D  date("$S_DATE", "DMY")
local newdate: di %td_CY-N-D date("$S_DATE","DMY")
di "`newdate'"
di trim("`newdate'")
global todaysdate: di trim("`newdate'")

* Set Put Excel
global putexcel_set "2. Analysis/BAL Baseline Dummy Tables Adolescent 02.xlsx"

* Set TopFive Program
global topfive "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_TopFive.do"

* Set GEM Program
global gem "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_GEM_CoupleCommunication.do"

* Set macros for cells (General Tables)
global cellnum1 3
global female_umscol "B"
global female_iucol "C"
global malecol "D"

* Set macros for cells (Demographic Table)
global cellnum2 2
global female_totalcol "B"

* Set macros for cells (Bivariate Analysis)
global cellnum3 1
global fp_col B
global chi2_col C

* Set macros for sheets
global part_1 "Part 1"
global part_2 "Part 2"
global part_3 "Part 3"
global part_4 "Part 4"
global demographic "Demographic Tables"
global bi_part_1 "Bivariate Part 1"
global bi_part_2 "Bivariate Part 2"
global bi_part_3 "Bivariate Part 3"
global multivariate "Multivariate"
	   
* Create log
log using "2. Analysis/BAL_Baseline_Adolescent_FamilyPlanning_$date.log", replace

********************************************************
*** READ IN DO FILES ***
********************************************************

do "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_Cleaning.do"
do "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_Tables.do"
do "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_Bivariate.do"
do "/Users/ealarson/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_Multivariate.do"
