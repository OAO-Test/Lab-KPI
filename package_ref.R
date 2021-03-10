#Install packages only the first time you run the code
# install.packages("timeDate")
# install.packages("lubridate")
# install.packages("readxl")
# install.packages("dplyr")
# install.packages("reshape2")
# install.packages("knitr")
# install.packages("gdtools")
# install.packages("kableExtra")
# install.packages("formattable")
# install.packages("bizdays")
# install.packages("rmarkdown")
# install.packages("stringr")
# install.packages("writexl")

#-------------------------------Required packages------------------------------#

#Required packages: run these every time you run the code
library(timeDate)
library(readxl)
library(bizdays)
library(dplyr)
library(lubridate)
library(reshape2)
library(knitr)
# library(gdtools)
library(kableExtra)
library(formattable)
library(rmarkdown)
library(stringr)
library(writexl)


#Clear existing history
rm(list = ls())
#-------------------------------holiday/weekend-------------------------------#
# Get today and yesterday's date

today <- as.timeDate(format(Sys.Date(), "%m/%d/%Y"))
#today <- as.timeDate(as.Date("03/04/2021", format = "%m/%d/%Y"))

#Determine if yesterday was a holiday/weekend
#get yesterday's DOW
yesterday <- as.timeDate(format(Sys.Date() - 1, "%m/%d/%Y"))
#yesterday <- as.timeDate(as.Date("03/03/2021", format = "%m/%d/%Y"))

#Get yesterday's DOW
yesterday_day <- dayOfWeek(yesterday)

#Remove Good Friday from MSHS Holidays
nyse_holidays <- as.Date(holidayNYSE(year = 1990:2100))
good_friday <- as.Date(GoodFriday())
mshs_holiday <- nyse_holidays[good_friday != nyse_holidays]

#Determine whether yesterday was a holiday/weekend
holiday_det <- isHoliday(yesterday, holidays = mshs_holiday)

#Set up a calendar for collect to received TAT calculations for Path & Cyto
create.calendar("MSHS_working_days", mshs_holiday,
                weekdays = c("saturday", "sunday"))
bizdays.options$set(default.calendar = "MSHS_working_days")


# Select file/folder path for easier file selection and navigation

if (list.files("J://") == "Presidents") {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab Kpi/Data")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab Kpi/Data")
}

# Import analysis reference data
reference_file <- paste0(user_directory,
                         "/Code Reference/",
                         "Analysis Reference 2021-02-23.xlsx")

# CP and Micro --------------------------------

test_code <- read_excel(reference_file, sheet = "TestNames")

tat_targets <- read_excel(reference_file, sheet = "Turnaround Targets")
#
# Add a column concatenating test, priority, and setting for matching later
tat_targets <- tat_targets %>%
  mutate(Concate = ifelse(
    Priority == "All" & `Pt Setting` == "All", Test,
    ifelse(Priority != "All" & `Pt Setting` == "All", paste(Test, Priority),
           paste(Test, Priority, `Pt Setting`))))

scc_icu <- read_excel(reference_file, sheet = "SCC_ICU")
scc_setting <- read_excel(reference_file, sheet = "SCC_ClinicType")
sun_icu <- read_excel(reference_file, sheet = "SUN_ICU")
sun_setting <- read_excel(reference_file, sheet = "SUN_LocType")

mshs_site <- read_excel(reference_file, sheet = "SiteNames")

###code to be moved to the CP code
#scc_wday <- scc_weekday
#sun_wday <- sq_weekday

cp_micro_lab_order <- c("Troponin",
                        "Lactate WB",
                        "BUN",
                        "HGB",
                        "PT",
                        "Rapid Flu",
                        "C. diff")

site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM", "MSSN")
city_sites <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM")

pt_setting_order <- c("ED", "ICU", "IP Non-ICU", "Amb", "Other")
pt_setting_order2 <- c("ED & ICU", "IP Non-ICU", "Amb", "Other")
dashboard_pt_setting <- c("ED & ICU", "IP Non-ICU", "Amb")

dashboard_priority_order <- c("All", "Stat", "Routine")

#-----------Patient Setting Excel File-----------#
#Using Rev Center to determine patient setting
patient_setting <- data.frame(read_excel(reference_file,
                                         sheet = "AP_Patient Setting"),
                              stringsAsFactors = FALSE)

#-----------Anatomic Pathology Targets Excel File-----------#
tat_targets_ap <- data.frame(read_excel(reference_file,
                                        sheet = "AP_TAT Targets"),
                             stringsAsFactors = FALSE)

#-----------GI Codes Excel File-----------#
#Upload the exclusion vs inclusion criteria associated with the GI codes
gi_codes <- data.frame(read_excel(reference_file, sheet = "GI_Codes"),
                       stringsAsFactors = FALSE)

#-----------Create table template for Cyto/Path-----------#
#The reason behind the table templates is to make sure all the variables and
#patient settings are included

#Cyto
#this template for cyto is with an assumption that received to result is not
#centralized
table_temp_cyto <- data.frame(matrix(ncol = 19, nrow = 4))

colnames(table_temp_cyto) <- c("spec_group", "Patient.Setting",
                               "no_cases_signed",
                               "MSH.x", "BIMC.x", "MSQ.x", "NYEE.x",
                               "PACC.x", "R.x", "SL.x", "KH.x", "BIMC.y",
                               "MSH.y", "MSQ.y", "NYEE.y", "PACC.y", "R.y",
                               "SL.y", "KH.y")

table_temp_cyto[1] <- c("CYTO GYN", "CYTO GYN", "CYTO NONGYN", "CYTO NONGYN")
table_temp_cyto[2] <- c("IP", "Amb")

#this template for cyto is with an assumption that received to result is
#centralized
table_temp_cyto_v2 <- data.frame(matrix(ncol = 12, nrow = 4))

colnames(table_temp_cyto_v2) <- c("spec_group", "Patient.Setting",
                                  "no_cases_signed",
                                  "received_to_signed_out_within_target",
                                  "BIMC", "MSH", "MSQ", "NYEE", "PACC",
                                  "R", "SL", "KH")

table_temp_cyto_v2[1] <- c("CYTO GYN", "CYTO GYN", "CYTO NONGYN", "CYTO NONGYN")
table_temp_cyto_v2[2] <- c("IP", "Amb")

#this table template is for cytology volume
table_temp_cyto_vol <- data.frame(matrix(ncol = 10, nrow = 4))

colnames(table_temp_cyto_vol) <- c("spec_group", "Patient.Setting",
                                   "BIMC", "MSH", "MSQ", "NYEE", "PACC",
                                   "R", "SL", "KH")

table_temp_cyto_vol[1] <- c("CYTO GYN", "CYTO GYN",
                            "CYTO NONGYN", "CYTO NONGYN")
table_temp_cyto_vol[2] <- c("IP", "Amb")

#Patho
#this template for patho (sp) is with an assumption that received to result is
#not centralized
table_temp_patho <- data.frame(matrix(ncol = 17, nrow = 4))
colnames(table_temp_patho) <- c("spec_group", "Patient.Setting",
                                "no_cases_signed",
                                "MSH.x", "BIMC.x", "MSQ.x", "PACC.x",
                                "R.x", "SL.x", "KH.x", "BIMC.y", "MSH.y",
                                "MSQ.y", "PACC.y", "R.y", "SL.y", "KH.y")

table_temp_patho[1] <- c("Breast", "Breast", "GI", "GI")
table_temp_patho[2] <- c("IP", "Amb")

#this table template is for surgical pathology volume
table_temp_patho_vol <- data.frame(matrix(ncol = 9, nrow = 4))
colnames(table_temp_patho_vol) <- c("spec_group", "Patient.Setting",
                                    "BIMC", "MSH", "MSQ", "PACC",
                                    "R", "SL", "KH")

table_temp_patho_vol[1] <- c("Breast", "Breast", "GI", "GI")
table_temp_patho_vol[2] <- c("IP", "Amb")
