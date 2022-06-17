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
* User
global user ealarson
local user $user

* Directory
cd "/Users/`user'/Documents/CCP/BA_Liberia/Men and Women"
global directory "/Users/`user'/Documents/CCP/BA_Liberia/Men and Women"

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
global putexcel_set "/Users/`user'/Documents/CCP/BA_Liberia/Men and Women/2. Analysis/Men_Women_Report.xlsx"


* Set macros for cells (Maternal Health)
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

* Set macros for cells (Psychosocial)
global R "1"
global rowlabel "A"
global intervention_collab "B"
global intervention_men "B"
global intervention_women "C"
global control_collab "D"
global control_men "D"
global control_women "E"
global total_collab "F"
global total_men "F"
global total_women "G"

* Set macros for cells (Media)
global R "1"
global rowlabel "A"
global women_collab "B"
global women_intervention "B"
global women_control "C"
global men_collab "D"
global men_intervention "D"
global men_control "E"
	   
* Create log
log using "2. Analysis/BA_Liberia_Men_Women_Analysis_$date.log", replace

* Call in data
use "/Users/`user'/Documents/CCP/BA_Liberia/Men and Women/1. Data/BAL_baseline_women&men_cleaned_dataset_v4.dta", clear

* Prep data
drop if exclude==1
gen one=1

save "/Users/`user'/Documents/CCP/BA_Liberia/Men and Women/1. Data/BA_Liberia_Men_Women_Analysis_Data.dta", replace

********************************************************
*** READ IN DO FILES ***
********************************************************

*do "/Users/`user'/PMA_GitKraken/GitHub_Personal/CCP/BA_Liberia_MaternalHealth.do"
*do "/Users/`user'/PMA_GitKraken/GitHub_Personal/CCP/BA_Liberia_Psychosocial.do"
do "/Users/`user'/PMA_GitKraken/GitHub_Personal/CCP/BA_Liberia_MediaSources.do"
