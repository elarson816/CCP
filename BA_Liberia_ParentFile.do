** ParentFile

clear
capture log close

/*** Breathrough Action Liberia - Adolescent Family Planning ***

* Input: BAL_baseline_women&men_cleaned_dataset_v4.dta 
* Output: BA_Liberia_Men_Women_Analysis.xls		 
*/

********************************************************
*** SET UP DO FILE ***
********************************************************

* Directory
cd "/Users/Beth/Documents/CCP/BA_Liberia/Men and Women"
global directory "/Users/Beth/Documents/CCP/BA_Liberia/Men and Women"

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
global putexcel_set "/Users/Beth/Documents/CCP/BA_Liberia/Men and Women/2. Analysis/Men_Women_Report.xlsx"

* Set GEM Program
global gem "/Users/Beth/PMA_GitKraken/GitHub_Personal/CCP/BAL_Baseline_GEM_CoupleCommunication.do"

* Set macros for cells (General Tables)
global R "1"
global rowlabel "A"
global women_collabel "B"
global women_intervention_col1 "B"
global women_intervention_col2 "C"
global women_intervention_col3 "D"
global women_control_col1 "E"
global men_collabel "F"
global men_intervention_col1 "F"
global men_intervention_col2 "G"
global men_intervention_col3 "H"
global men_control_col1 "I"

	   
* Create log
log using "2. Analysis/BA_Liberia_Men_Women_Analysis_$date.log", replace

* Call in data
use "/Users/Beth/Documents/CCP/BA_Liberia/Men and Women/1. Data/BAL_baseline_women&men_cleaned_dataset_v4.dta", clear

* Prep data
drop if exclude==1
gen one=1

save "BA_Liberia_Men_Women_Analysis_Data.dta", replace

********************************************************
*** READ IN DO FILES ***
********************************************************

do "/Users/Beth/PMA_GitKraken/GitHub_Personal/CCP/BA_Liberia_MaternalHealth.do"
