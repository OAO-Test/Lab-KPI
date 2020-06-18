#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("lubridate")
#install.packages("readxl")
#install.packages("dplyr")
#install.packages("reshape2")
#install.packages("knitr")
#install.packages("kableExtra")
#install.packages("formattable")
#install.packages("bizdays")
#install.packages("rmarkdown")
#install.packages("stringr")
#install.packages("writexl")

#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(bizdays)
library(dplyr)
library(lubridate)
library(reshape2)
library(knitr)
library(kableExtra)
library(formattable)
library(rmarkdown)
library(stringr)
library(writexl)

rm(list = ls())

# Set working directory -------------------------------
# reference_file <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data\\Code Reference\\Analysis Reference 2020-01-22.xlsx"
user_wd <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data"
user_path <- paste0(user_wd, "\\*.*")
setwd(user_wd)

# Import data for two scenarious - first time compiling repository and updating repository -----------------------
initial_run <- TRUE

if (initial_run == TRUE) {
  # Find list of data reports from 2020
  file_list_scc <- list.files(path = paste0(user_wd, "\\SCC CP Reports"), pattern = "^(Doc){1}.+(2020)\\-[0-9]{2}-[0-9]{2}.xlsx")
  
  # Pattern for daily Sunquest reports
  sun_daily_pattern = c("^(KPI_Daily_TAT_Report ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls", "^(KPI_Daily_TAT_Report_Updated ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls")
  # file_list_sun_daily <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = "^(KPI_Daily_TAT_Report ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls")
  file_list_sun_daily <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = paste0(sun_daily_pattern, collapse = "|"))
  file_list_sun_monthly <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = "^(KPI_TAT Report_){1}[A-z]+\\s(2020.xlsx)")
  # Read in data reports from possible date range
  # scc_list <- lapply(file_list_scc, function(x) read_excel(path = paste0(user_wd, "\\SCC CP Reports\\", x)))
  sun_daily_list <- lapply(file_list_sun_daily, function(x) (read_excel(path = paste0(user_wd, "\\SUN CP Reports\\", x),
                                                                        col_types = c("text", "text", "text", "text", "text", 
                                                                                      "text", "text", "text", "text", 
                                                                                      "numeric", "numeric", "numeric", "numeric", "numeric",
                                                                                      "text", "text", "text", "text", "text", 
                                                                                      "text", "text", "text", "text", "text", 
                                                                                      "text", "text", "text", "text", "text", 
                                                                                      "text", "text", "text", "text", "text", "text"))))
  sun_monthly_list <- lapply(file_list_sun_monthly, function(x) (read_excel(path = paste0(user_wd, "\\SUN CP Reports\\", x),
                                                                                           col_types = c("text", "text", "text", "text", "text", 
                                                                                                         "date", "date", "date", "date", 
                                                                                                         "numeric", "numeric", "numeric", "numeric", "numeric",
                                                                                                         "text", "text", "text", "text", "text", 
                                                                                                         "text", "text", "text", "text", "text", 
                                                                                                         "text", "text", "text", "text", "text", 
                                                                                                         "text", "text", "text", "text", "text", "text"))))
} else {
  # Import existing historical repository
  existing_repo <- read_excel(choose.files(default = user_path, caption = "Select Historical Repository"), sheet = 1, col_names = TRUE)
  #
  # Find last date of resulted lab data in historical repository
  last_date <- as.Date(max(existing_repo$ResultDate), format = "%Y-%m-%d")
  # Determine today's date to determine last possible data report
  todays_date <- as.Date(Sys.Date(), format = "%Y-%m-%d")
  # Create vector with possible data report dates
  date_range <- seq(from = last_date + 2, to = todays_date, by = "day")
  # Find list of data reports from date range
  file_list_scc <- list.files(path = paste0(user_wd, "\\SCC CP Reports"), pattern = paste0("^(Doc){1}.+", date_range, ".xlsx", collapse = "|"))
  file_list_sun <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = paste0("^(KPI_Daily_TAT_Report ){1}", date_range, ".xls", collapse = "|"))
  # Read in data reports from possible date range
  scc_list <- lapply(file_list_scc, function(x) read_excel(path = paste0(user_wd, "\\SCC CP Reports\\", x)))
  sun_list <- lapply(file_list_sun, function(x) suppressWarnings(read_excel(path = paste0(user_wd, "\\SUN CP Reports\\", x))))
}

# file_list_scc <- list.files(path = paste0(user_wd, "\\SCC CP Reports"), pattern = "^(Doc){1}.+(2020)\\-[0-9]{2}-[0-9]{2}.xlsx")
# file_list_sun <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = "^(KPI_Daily_TAT_Report ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls")
# 
# scc_list <- lapply(file_list_scc, function(x) read_excel(path = paste0(user_wd, "\\SCC CP Reports\\", x)))
# sun_list <- lapply(file_list_sun, function(x) suppressWarnings(read_excel(path = paste0(user_wd, "\\SUN CP Reports\\", x))))

# Import analysis reference data starting with test codes for SCC and Sunquest --------------------------------------
reference_file <- choose.files(default = user_path, caption = "Select analysis reference file")

test_code <- read_excel(reference_file, sheet = "TestNames")

tat_targets <- read_excel(reference_file, sheet = "Turnaround Targets")
tat_targets$Concate <- ifelse(tat_targets$Priority == "All" & tat_targets$`Pt Setting` == "All", tat_targets$Test,
                              ifelse(tat_targets$Priority != "All" & tat_targets$`Pt Setting` == "All", paste(tat_targets$Test, tat_targets$Priority),
                                     paste(tat_targets$Test, tat_targets$Priority, tat_targets$`Pt Setting`)))

scc_icu <- read_excel(reference_file, sheet = "SCC_ICU")
scc_setting <- read_excel(reference_file, sheet = "SCC_ClinicType")
sun_icu <- read_excel(reference_file, sheet = "SUN_ICU")
sun_setting <- read_excel(reference_file, sheet = "SUN_LocType")

mshs_site <- read_excel(reference_file, sheet = "SiteNames")

cp_micro_lab_order <- c("Troponin", "Lactate WB", "BUN", "HGB", "PT", "Rapid Flu", "C. diff")

site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM")
pt_setting_order <- c("ED", "ICU", "IP Non-ICU", "Amb", "Other")
pt_setting_order2 <- c("ED & ICU", "IP Non-ICU", "Amb", "Other")
dashboard_pt_setting <- c("ED & ICU", "IP Non-ICU", "Amb")

dashboard_priority_order <- c("All", "Stat", "Routine")

# # Preprocess new data --------------------------------
# preprocess_scc_sun <- function(raw_scc, raw_sun)  {
#   # SCC DATA PROCESSING --------------------------
#   # SCC lookup references ----------------------------------------------
#   # Crosswalk labs included and remove out of scope labs
#   raw_scc <- left_join(raw_scc, test_code[ , c("Test", "SCC_TestID", "Division")], by = c("TEST_ID" = "SCC_TestID"))
#   raw_scc$TEST_ID <- as.factor(raw_scc$TEST_ID)
#   raw_scc$Division <- as.factor(raw_scc$Division)
#   raw_scc$TestIncl <- ifelse(is.na(raw_scc$Test), FALSE, TRUE)
#   raw_scc <- raw_scc[raw_scc$TestIncl == TRUE, ]
#   # Crosswalk units and identify ICUs
#   raw_scc$WardandName <- paste(raw_scc$Ward, raw_scc$WARD_NAME)
#   raw_scc <- left_join(raw_scc, scc_icu[ , c("Concatenate", "ICU")], by = c("WardandName" = "Concatenate"))
#   raw_scc[is.na(raw_scc$ICU), "ICU"] <- FALSE
#   # Crosswalk unit type
#   raw_scc <- left_join(raw_scc, scc_setting, by = c("CLINIC_TYPE" = "Clinic_Type"))
#   # Crosswalk site name
#   raw_scc <- left_join(raw_scc, mshs_site, by = c("SITE" = "DataSite"))
#   
#   # SCC data formatting ----------------------------------------------
#   raw_scc[c("Ward", "WARD_NAME", 
#             "REQUESTING_DOC", 
#             "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "Test",
#             "COLLECT_CENTER_ID", "SITE", "Site",
#             "CLINIC_TYPE", "Setting", "SettingRollUp")] <- lapply(raw_scc[c("Ward", "WARD_NAME", 
#                                                                             "REQUESTING_DOC", 
#                                                                             "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "Test",
#                                                                             "COLLECT_CENTER_ID", "SITE", "Site",
#                                                                             "CLINIC_TYPE", "Setting", "SettingRollUp")], as.factor)
#   
#   # Fix any timestamps that weren't imported correctly and then format as date/time
#   raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")], function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, str_replace(x, "\\*.*\\*", ""), x))
#   
#   raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")],
#                                                                                             as.POSIXlt, tz = "", format = "%Y-%m-%d %H:%M:%OS", options(digits.sec = 1))
#   
#   # Add a column for Resulted date for later use in repository
#   raw_scc$ResultedDate <- as.Date(raw_scc$VERIFIED_DATE, format, "%m/%d/%Y")
#   
#   # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
#   raw_scc$MasterSetting <- ifelse(raw_scc$CLINIC_TYPE == "E", "ED", 
#                                   ifelse(raw_scc$CLINIC_TYPE == "O", "Amb", 
#                                          ifelse(raw_scc$CLINIC_TYPE == "I" & raw_scc$ICU == TRUE, "ICU", 
#                                                 ifelse(raw_scc$CLINIC_TYPE == "I" & raw_scc$ICU != TRUE, "IP Non-ICU", "Other"))))
#   raw_scc$DashboardSetting <- ifelse(raw_scc$MasterSetting == "ED" | raw_scc$MasterSetting == "ICU", "ED & ICU", raw_scc$MasterSetting)
#   
#   
#   # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
#   raw_scc$AdjPriority <- ifelse(raw_scc$MasterSetting == "ED" | raw_scc$MasterSetting == "ICU" | raw_scc$PRIORITY == "S", "Stat", "Routine")
#   raw_scc$DashboardPriority <- ifelse(tat_targets$Priority[match(raw_scc$Test, tat_targets$Test)] == "All", "All", raw_scc$AdjPriority)
#   
#   # Calculate turnaround times
#   raw_scc$CollectToReceive <- raw_scc$RECEIVE_DATE - raw_scc$COLLECTION_DATE
#   raw_scc$ReceiveToResult <- raw_scc$VERIFIED_DATE - raw_scc$RECEIVE_DATE
#   raw_scc$CollectToResult <- raw_scc$VERIFIED_DATE - raw_scc$COLLECTION_DATE
#   raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
#   
#   # Identify add on orders as orders placed more than 5 min after specimen received
#   raw_scc$AddOnMaster <- ifelse(difftime(raw_scc$ORDERING_DATE, raw_scc$RECEIVE_DATE, units = "mins") > 5, "AddOn", "Original")
#   
#   # Identify specimens with missing collections times as those with collection time defaulted to receive time
#   raw_scc$MissingCollect <- ifelse(raw_scc$CollectToReceive == 0, TRUE, FALSE)
#   
#   # Determine target TAT based on test, priority, and patient setting
#   raw_scc$Concate1 <- paste(raw_scc$Test, raw_scc$DashboardPriority)
#   raw_scc$Concate2 <- paste(raw_scc$Test, raw_scc$DashboardPriority, raw_scc$MasterSetting)
#   
#   raw_scc$ReceiveResultTarget <- ifelse(!is.na(match(raw_scc$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_scc$Concate2, tat_targets$Concate)], 
#                                         ifelse(!is.na(match(raw_scc$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_scc$Concate1, tat_targets$Concate)],
#                                                tat_targets$ReceiveToResultTarget[match(raw_scc$Test, tat_targets$Concate)]))
#   
#   raw_scc$CollectResultTarget <- ifelse(!is.na(match(raw_scc$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_scc$Concate2, tat_targets$Concate)], 
#                                         ifelse(!is.na(match(raw_scc$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_scc$Concate1, tat_targets$Concate)],
#                                                tat_targets$CollectToResultTarget[match(raw_scc$Test, tat_targets$Concate)]))
#   
#   raw_scc$ReceiveResultInTarget <- ifelse(raw_scc$ReceiveToResult <= raw_scc$ReceiveResultTarget, TRUE, FALSE)
#   raw_scc$CollectResultInTarget <- ifelse(raw_scc$CollectToResult <= raw_scc$CollectResultTarget, TRUE, FALSE)
#   
#   # Identify and remove duplicate tests
#   raw_scc$Concate3 <- paste(raw_scc$LAST_NAME, raw_scc$FIRST_NAME, 
#                             raw_scc$ORDER_ID, raw_scc$TEST_NAME,
#                             raw_scc$COLLECTION_DATE, raw_scc$RECEIVE_DATE, raw_scc$VERIFIED_DATE)
#   
#   raw_scc <- raw_scc[!duplicated(raw_scc$Concate3), ]
#   
#   # Identify which labs to include in TAT analysis
#   # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
#   raw_scc$TATInclude <- ifelse(raw_scc$AddOnMaster != "Original" | raw_scc$MasterSetting == "Other" | raw_scc$CollectToResult < 0 | raw_scc$ReceiveToResult < 0 | is.na(raw_scc$CollectToResult) | is.na(raw_scc$ReceiveToResult), FALSE, TRUE)
#   
#   scc_master <- raw_scc[ , c("Ward", "WARD_NAME", "WardandName",
#                              "ORDER_ID", "REQUESTING_DOC NAME", "MPI", "WORK SHIFT",
#                              "TEST_NAME", "Test", "Division", "PRIORITY", 
#                              "Site", "ICU", "CLINIC_TYPE", 
#                              "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting",
#                              "AdjPriority", "DashboardPriority",
#                              "ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE",
#                              "ResultedDate", 
#                              "CollectToReceive", "ReceiveToResult", "CollectToResult", 
#                              "AddOnMaster", "MissingCollect", 
#                              "ReceiveResultTarget", "CollectResultTarget", 
#                              "ReceiveResultInTarget", "CollectResultInTarget", 
#                              "TATInclude")]
#   colnames(scc_master) <- c("LocCode", "LocName", "LocConcat",
#                             "OrderID", "RequestMD", "MSMRN", "WorkShift", 
#                             "TestName", "Test", "Division", "OrderPriority", 
#                             "Site", "ICU", "LocType", 
#                             "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting",
#                             "AdjPriority", "DashboardPriority",
#                             "OrderTime", "CollectTime", "ReceiveTime", "ResultTime", 
#                             "ResultDate", 
#                             "CollectToReceiveTAT", "ReceiveToResultTAT", "CollectToResultTAT", 
#                             "AddOnMaster", "MissingCollect", 
#                             "ReceiveResultTarget", "CollectResultTarget", 
#                             "ReceiveResultInTarget", "CollectResultInTarget", 
#                             "TATInclude")
#   
#   ## SUNQUEST DATA PROCESSING ----------
#   # Sunquest lookup references ----------------------------------------------
#   # Crosswalk labs included and remove out of scope labs
#   raw_sun <- left_join(raw_sun, test_code[ , c("Test", "SUN_TestCode", "Division")], by = c("TestCode" = "SUN_TestCode"))
#   raw_sun$TestCode <- as.factor(raw_sun$TestCode)
#   raw_sun$Division <- as.factor(raw_sun$Division)
#   raw_sun$TestIncl <- ifelse(is.na(raw_sun$Test), FALSE, TRUE)
#   raw_sun <- raw_sun[raw_sun$TestIncl == TRUE, ]
#   # Crosswalk units and identify ICUs
#   raw_sun$LocandName <- paste(raw_sun$LocCode, raw_sun$LocName)
#   raw_sun <- left_join(raw_sun, sun_icu[ , c("Concatenate", "ICU")], by = c("LocandName" = "Concatenate"))
#   raw_sun[is.na(raw_sun$ICU), "ICU"] <- FALSE
#   # Crosswalk unit type
#   raw_sun <- left_join(raw_sun, sun_setting, by = c("LocType" = "LocType"))
#   # Crosswalk site name
#   raw_sun <- left_join(raw_sun, mshs_site, by = c("HospCode" = "DataSite"))
#   
#   # Sunquest data formatting --------------------------------------------
#   raw_sun[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
#             "LocType", "LocCode", "LocName", 
#             "PhysName", "SHIFT",
#             "ReceiveTech", "ResultTech", "PerformingLabCode",
#             "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- lapply(raw_sun[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
#                                                                                            "LocType", "LocCode", "LocName", 
#                                                                                            "PhysName", "SHIFT",
#                                                                                            "ReceiveTech", "ResultTech", "PerformingLabCode",
#                                                                                            "Test", "LocandName", "Setting", "SettingRollUp", "Site")], as.factor)
#   
#   # Add NA as factor level for SpecimenPriority since many specimens are submitted without a priority
#   raw_sun$SpecimenPriority <- addNA(raw_sun$SpecimenPriority)
#   
#   # Fix any timestamps that weren't imported correctly and then format as date/time
#   raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, str_replace(x, "\\*.*\\*", ""), x))
#   
#   raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], 
#                                                                                                 as.POSIXlt, tz = "", format = "%m/%d/%Y %H:%M:%S")
#   
#   # Add a column for Resulted date for later use in repository
#   raw_sun$ResultedDate <- as.Date(raw_sun$ResultDateTime, format = "%m/%d/%Y")
#   
#   
#   
#   # Sunquest data preprocessing --------------------------------------------------
#   # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
#   raw_sun$MasterSetting <- ifelse(raw_sun$SettingRollUp == "ED", "ED", 
#                                   ifelse(raw_sun$SettingRollUp == "Amb", "Amb", 
#                                          ifelse(raw_sun$SettingRollUp == "IP" & raw_sun$ICU == TRUE, "ICU",
#                                                 ifelse(raw_sun$SettingRollUp == "IP" & raw_sun$ICU == FALSE, "IP Non-ICU", "Other"))))
#   raw_sun$DashboardSetting <- ifelse(raw_sun$MasterSetting == "ED" | raw_sun$MasterSetting == "ICU", "ED & ICU", raw_sun$MasterSetting)
#   
#   # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
#   raw_sun$AdjPriority <- ifelse(raw_sun$MasterSetting != "ED" & raw_sun$MasterSetting != "ICU" & is.na(raw_sun$SpecimenPriority), "Routine",
#                                 ifelse(raw_sun$MasterSetting == "ED" | raw_sun$MasterSetting == "ICU" | raw_sun$SpecimenPriority == "S", "Stat", "Routine"))
#   raw_sun$DashboardPriority <- ifelse(tat_targets$Priority[match(raw_sun$Test, tat_targets$Test)] == "All", "All", raw_sun$AdjPriority)
#   
#   # Calculate turnaround times
#   raw_sun$CollectToReceive <- raw_sun$ReceiveDateTime - raw_sun$CollectDateTime
#   raw_sun$ReceiveToResult <- raw_sun$ResultDateTime - raw_sun$ReceiveDateTime
#   raw_sun$CollectToResult <- raw_sun$ResultDateTime - raw_sun$CollectDateTime
#   raw_sun[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(raw_sun[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
#   
#   # Identify add on orders as orders placed more than 5 min after specimen received
#   raw_sun$AddOnMaster <- ifelse(difftime(raw_sun$OrderDateTime, raw_sun$ReceiveDateTime, units = "mins") > 5, "AddOn", "Original")
#   
#   # Identify specimens with missing collections times as those with collection time defaulted to order time
#   raw_sun$MissingCollect <- ifelse(raw_sun$CollectDateTime == raw_sun$OrderDateTime, TRUE, FALSE)
#   
#   # Determine target TAT based on test, priority, and patient setting
#   raw_sun$Concate1 <- paste(raw_sun$Test, raw_sun$DashboardPriority)
#   raw_sun$Concate2 <- paste(raw_sun$Test, raw_sun$DashboardPriority, raw_sun$MasterSetting)
#   
#   raw_sun$ReceiveResultTarget <- ifelse(!is.na(match(raw_sun$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_sun$Concate2, tat_targets$Concate)], 
#                                         ifelse(!is.na(match(raw_sun$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_sun$Concate1, tat_targets$Concate)],
#                                                tat_targets$ReceiveToResultTarget[match(raw_sun$Test, tat_targets$Concate)]))
#   
#   raw_sun$CollectResultTarget <- ifelse(!is.na(match(raw_sun$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_sun$Concate2, tat_targets$Concate)], 
#                                         ifelse(!is.na(match(raw_sun$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_sun$Concate1, tat_targets$Concate)],
#                                                tat_targets$CollectToResultTarget[match(raw_sun$Test, tat_targets$Concate)]))
#   
#   raw_sun$ReceiveResultInTarget <- ifelse(raw_sun$ReceiveToResult <= raw_sun$ReceiveResultTarget, TRUE, FALSE)
#   raw_sun$CollectResultInTarget <- ifelse(raw_sun$CollectToResult <= raw_sun$CollectResultTarget, TRUE, FALSE)
#   
#   # Identify and remove duplicate tests
#   raw_sun$Concate3 <- paste(raw_sun$PtNumber, 
#                             raw_sun$HISOrderNumber, raw_sun$TSTName,
#                             raw_sun$CollectDateTime, raw_sun$ReceiveDateTime, raw_sun$ResultDateTime)
#   
#   raw_sun <- raw_sun[!duplicated(raw_sun$Concate3), ]
#   
#   
#   # Identify which labs to include in TAT analysis
#   # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
#   raw_sun$TATInclude <- ifelse(raw_sun$AddOnMaster != "Original" | raw_sun$MasterSetting == "Other" | raw_sun$CollectToResult < 0 | raw_sun$ReceiveToResult < 0 | is.na(raw_sun$CollectToResult) | is.na(raw_sun$ReceiveToResult), FALSE, TRUE)
#   
#   sun_master <- raw_sun[ ,c("LocCode", "LocName", "LocandName", 
#                             "HISOrderNumber", "PhysName", "PtNumber", "SHIFT",            
#                             "TSTName", "Test", "Division", "SpecimenPriority", 
#                             "Site", "ICU", "LocType",               
#                             "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting", 
#                             "AdjPriority", "DashboardPriority", 
#                             "OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime", 
#                             "ResultedDate", 
#                             "CollecttoReceive", "ReceivetoResult", "CollecttoResult", 
#                             "AddOnMaster", "MissingCollect", 
#                             "ReceiveResultTarget", "CollectResultTarget", 
#                             "ReceiveResultInTarget", "CollectResultInTarget", 
#                             "TATInclude")]
#   
#   colnames(sun_master) <- c("LocCode", "LocName", "LocConcat", 
#                             "OrderID", "RequestMD", "MSMRN", "WorkShift", 
#                             "TestName", "Test", "Division", "OrderPriority", 
#                             "Site", "ICU", "LocType", 
#                             "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting", 
#                             "AdjPriority", "DashboardPriority",
#                             "OrderTime", "CollectTime", "ReceiveTime", "ResultTime", 
#                             "ResultDate", 
#                             "CollectToReceiveTAT", "ReceiveToResultTAT", "CollectToResultTAT", 
#                             "AddOnMaster", "MissingCollect", 
#                             "ReceiveResultTarget", "CollectResultTarget", 
#                             "ReceiveResultInTarget", "CollectResultInTarget", 
#                             "TATInclude")
#   
#   
#   scc_sun_master <- rbind(scc_master, sun_master)
#   
#   scc_sun_master[c("LocConcat", "RequestMD", "WorkShift", "AdjPriority", "AddOnMaster")] <-
#     lapply(scc_sun_master[c("LocConcat", "RequestMD", "WorkShift", "AdjPriority", "AddOnMaster")], as.factor)
#   
#   scc_sun_master$Site <- factor(scc_sun_master$Site, levels = site_order)
#   scc_sun_master$Test <- factor(scc_sun_master$Test, levels = cp_micro_lab_order)
#   scc_sun_master$MasterSetting <- factor(scc_sun_master$MasterSetting, levels = pt_setting_order)
#   scc_sun_master$DashboardSetting <- factor(scc_sun_master$DashboardSetting, levels = pt_setting_order2)
#   scc_sun_master$DashboardPriority <- factor(scc_sun_master$DashboardPriority, levels = dashboard_priority_order)
#   
#   scc_sun_list <- list(raw_scc, raw_sun, scc_sun_master)
#   
# }
# 
# # Preprocess all SCC and Sunquest files --------------------------------------------
# preprocess_all_files <- mapply(preprocess_scc_sun, scc_list, sun_daily_list)
# 
# # Custom function to determine resulted lab date from preprocessed data (SCC data often has a few labs with incorrect result date) --------------------
# correct_result_dates <- function(data, number_days) {
#   all_resulted_dates_vol <- data %>%
#     group_by(ResultDate) %>%
#     summarize(VolLabs = n())
#   
#   all_resulted_dates_vol <- all_resulted_dates_vol[order(all_resulted_dates_vol$VolLabs, decreasing = TRUE), ]
#   
#   correct_dates <- all_resulted_dates_vol$ResultDate[1:number_days]
#   
#   new_data <- data[data$ResultDate %in% correct_dates, ]
#   return(new_data)
# }


# Simplify preprocessing code ---------
sun_monthly_raw2 <- sun_monthly_list[[1]]

# Remove duplicates
sun_monthly_raw2 <- unique(sun_monthly_raw2)

# Crosswalk labs included and remove out of scope labs
sun_monthly_raw2 <- left_join(sun_monthly_raw2, test_code[ , c("Test", "SUN_TestCode", "Division")], 
                             by = c("TestCode" = "SUN_TestCode"))

# Set test codes and lab divisions to factors. Determine if test should be included in dashboard
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(TestCode = as.factor(TestCode),
         Division = as.factor(Division),
         TestIncl = ifelse(is.na(Test), FALSE, TRUE))

# Remove tests that are not included in dashboard
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  filter(TestIncl == TRUE)

# Crosswalk units and identify ICUs
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(LocandName = paste(LocCode, LocName))
sun_monthly_raw2 <- left_join(sun_monthly_raw2, sun_icu[ , c("Concatenate", "ICU")], by = c("LocandName" = "Concatenate"))

sun_monthly_raw2[is.na(sun_monthly_raw2$ICU), "ICU"] <- FALSE
# Crosswalk unit type
sun_monthly_raw2 <- left_join(sun_monthly_raw2, sun_setting, by = c("LocType" = "LocType"))
# Crosswalk site name
sun_monthly_raw2 <- left_join(sun_monthly_raw2, mshs_site, by = c("HospCode" = "DataSite"))

# Sunquest data formatting --------------------------------------------
sun_monthly_raw2[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
                        "LocType", "LocCode", "LocName", 
                        "PhysName", "SHIFT",
                        "ReceiveTech", "ResultTech", "PerformingLabCode",
                        "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- 
  lapply(sun_monthly_raw2[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName", 
                            "LocType", "LocCode", "LocName", "PhysName", "SHIFT", "ReceiveTech", 
                            "ResultTech", "PerformingLabCode", "Test", "LocandName", "Setting", "SettingRollUp", "Site")], as.factor)

# Add NA as factor level for SpecimenPriority since many specimens are submitted without a priority
# Add a column for resulted date for later use in repository
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(SpecimenPriority = addNA(SpecimenPriority),
         ResultedDate = as.Date(ResultDateTime, format = "%m/%d/%Y"))

# Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(MasterSetting = ifelse(SettingRollUp == "ED", "ED", 
                               ifelse(SettingRollUp == "Amb", "Amb", 
                                      ifelse(SettingRollUp == "IP" & ICU == TRUE, "ICU",
                                             ifelse(SettingRollUp == "IP" & ICU == FALSE, "IP Non-ICU", "Other")))),
         DashboardSetting = ifelse(MasterSetting == "ED" | MasterSetting == "ICU", "ED & ICU", MasterSetting))
         

# Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(AdjPriority = ifelse(MasterSetting != "ED" & MasterSetting != "ICU" & is.na(SpecimenPriority), "Routine", 
                              ifelse(MasterSetting == "ED" | MasterSetting == "ICU" | SpecimenPriority == "S", "Stat", "Routine")),
         DashboardPriority = ifelse(tat_targets$Priority[match(Test, tat_targets$Test)] == "All", "All", AdjPriority))

# Calculate turnaround times
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(CollectToReceive = ReceiveDateTime - CollectDateTime,
         ReceiveToResult = ResultDateTime - ReceiveDateTime,
         CollectToResult = ResultDateTime - CollectDateTime)

sun_monthly_raw2[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(sun_monthly_raw2[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")

# Identify add on orders as orders placed more than 5 min after specimen received
# Identify specimens with missing collections times as those with collection time defaulted to order time
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(AddOnMaster = ifelse(difftime(OrderDateTime, ReceiveDateTime, units = "mins") > 5, "AddOn", "Original"),
         MissingCollect = CollectDateTime == OrderDateTime)

# Determine target TAT based on test, priority, and patient setting
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(Concate1 = paste(Test, DashboardPriority),
         Concate2 = paste(Test, DashboardPriority, MasterSetting),
         ReceiveResultTarget = ifelse(!is.na(match(Concate2, tat_targets$Concate)), 
                                      tat_targets$ReceiveToResultTarget[match(Concate2, tat_targets$Concate)], 
                                              ifelse(!is.na(match(Concate1, tat_targets$Concate)), 
                                                     tat_targets$ReceiveToResultTarget[match(Concate1, tat_targets$Concate)],
                                                     tat_targets$ReceiveToResultTarget[match(Test, tat_targets$Concate)])),
         CollectResultTarget = ifelse(!is.na(match(Concate2, tat_targets$Concate)), 
                                      tat_targets$CollectToResultTarget[match(Concate2, tat_targets$Concate)],
                                      ifelse(!is.na(match(Concate1, tat_targets$Concate)), 
                                             tat_targets$CollectToResultTarget[match(Concate1, tat_targets$Concate)],
                                             tat_targets$CollectToResultTarget[match(Test, tat_targets$Concate)])),
         ReceiveResultInTarget = ReceiveToResult <= ReceiveResultTarget,
         CollectResultInTarget = CollectToResult <= CollectResultTarget)

# Identify and remove duplicate tests
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(Concate3 = paste(PtNumber, HISOrderNumber, TSTName, CollectDateTime, 
                           ReceiveDateTime, ResultDateTime))

sun_monthly_raw2 <- sun_monthly_raw2[!duplicated(sun_monthly_raw2$Concate3), ]

# Identify which labs to include in TAT analysis
# Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
sun_monthly_raw2 <- sun_monthly_raw2 %>%
  mutate(TATInclude = ifelse(AddOnMaster != "Original" | MasterSetting == "Other" | CollectToResult < 0 | 
                               ReceiveToResult < 0 | is.na(CollectToResult) | is.na(ReceiveToResult), FALSE, TRUE))

sun_monthly_master2 <- sun_monthly_raw2[ ,c("LocCode", "LocName", "LocandName", 
                                          "HISOrderNumber", "PhysName", "PtNumber", "SHIFT",            
                                          "TSTName", "Test", "Division", "SpecimenPriority", 
                                          "Site", "ICU", "LocType",               
                                          "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting", 
                                          "AdjPriority", "DashboardPriority", 
                                          "OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime", 
                                          "ResultedDate", 
                                          "CollecttoReceive", "ReceivetoResult", "CollecttoResult", 
                                          "AddOnMaster", "MissingCollect", 
                                          "ReceiveResultTarget", "CollectResultTarget", 
                                          "ReceiveResultInTarget", "CollectResultInTarget", 
                                          "TATInclude")]

colnames(sun_monthly_master2) <- c("LocCode", "LocName", "LocConcat", 
                                  "OrderID", "RequestMD", "MSMRN", "WorkShift", 
                                  "TestName", "Test", "Division", "OrderPriority", 
                                  "Site", "ICU", "LocType", 
                                  "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting", 
                                  "AdjPriority", "DashboardPriority",
                                  "OrderTime", "CollectTime", "ReceiveTime", "ResultTime", 
                                  "ResultDate", 
                                  "CollectToReceiveTAT", "ReceiveToResultTAT", "CollectToResultTAT", 
                                  "AddOnMaster", "MissingCollect", 
                                  "ReceiveResultTarget", "CollectResultTarget", 
                                  "ReceiveResultInTarget", "CollectResultInTarget", 
                                  "TATInclude")









# # Count total rows in all reports to compare after incorrect dates are removed
# total_row <- 0
# for (i in seq(from = 3, to = length(preprocess_all_files), by = 3)) {
#   total_row <- total_row + nrow(preprocess_all_files[[i]])
# }

# Remove any data with incorrect date from master files individually and then combine all reports into one dataframe ------------------------
bind_all_data <- NULL
for (i in seq(from = 3, to = length(preprocess_all_files), by = 3)) {
  preprocess_all_files[[i]] <- correct_result_dates(preprocess_all_files[[i]], 1)
  bind_all_data <- rbind(bind_all_data, preprocess_all_files[[i]])
}
# new_total_row <- nrow(bind_all_data)

# Remove duplicate entries across days
bind_all_data <- unique(bind_all_data)


# Summarize data for export to historical repo ---------------------------------
scc_sun_all_days_subset <- bind_all_data %>%
  group_by(Site, ResultDate, Test, Division, 
           Setting, SettingRollUp, MasterSetting, DashboardSetting,
           OrderPriority, AdjPriority, DashboardPriority,
           ReceiveResultTarget, CollectResultTarget) %>%
  summarize(TotalResulted = n(), TotalResultedTAT = sum(TATInclude), TotalReceiveResultInTarget = sum(ReceiveResultInTarget[TATInclude == TRUE]), TotalCollectResultInTarget = sum(CollectResultInTarget[TATInclude == TRUE]), TotalAddOnOrder = sum(AddOnMaster == "AddOn"), TotalMissingCollections = sum(MissingCollect)) %>%
  ungroup()

if (initial_run == TRUE) {
  scc_sun_repo <- scc_sun_all_days_subset
} else {
  # Format existing repository for binding
  existing_repo$Site <- factor(existing_repo$Site, levels = site_order)
  existing_repo$Test <- factor(existing_repo$Test, levels = cp_micro_lab_order)
  existing_repo$MasterSetting <- factor(existing_repo$MasterSetting, levels = pt_setting_order)
  existing_repo$DashboardSetting <- factor(existing_repo$DashboardSetting, levels = pt_setting_order2)
  existing_repo$DashboardPriority <- factor(existing_repo$DashboardPriority, levels = dashboard_priority_order)
  # 
  # Bind new data with repository
  scc_sun_repo <- rbind(existing_repo, scc_sun_all_days_subset)
}

start_date <- format(min(scc_sun_repo$ResultDate), "%m-%d-%y")
end_date <- format(max(scc_sun_repo$ResultDate), "%m-%d-%y")

write_xlsx(scc_sun_repo, path = paste0(user_wd, "\\SCC Sunquest Script Repo", "\\Hist Repo Test ", start_date, " to ", end_date, " Created ", Sys.Date(), ".xlsx"))

