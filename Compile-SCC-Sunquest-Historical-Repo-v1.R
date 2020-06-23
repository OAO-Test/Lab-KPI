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
initial_run <- FALSE

if (initial_run == TRUE) {
  # Find list of data reports from 2020
  file_list_scc <- list.files(path = paste0(user_wd, "\\SCC CP Reports"), pattern = "^(Doc){1}.+(2020)\\-[0-9]{2}-[0-9]{2}.xlsx")
  
  # Pattern for daily Sunquest reports
  sun_daily_pattern = c("^(KPI_Daily_TAT_Report ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls", "^(KPI_Daily_TAT_Report_Updated ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls")

  # file_list_sun_daily <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = "^(KPI_Daily_TAT_Report ){1}(2020)\\-[0-9]{2}-[0-9]{2}.xls")
  file_list_sun_daily <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = paste0(sun_daily_pattern, collapse = "|"))
  file_list_sun_monthly <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = "^(KPI_TAT Report_){1}[A-z]+\\s(2020.xlsx)")

  # Read in data reports from possible date range
  scc_list <- lapply(file_list_scc, function(x) read_excel(path = paste0(user_wd, "\\SCC CP Reports\\", x)))
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
  
  # Pattern for daily Sunquest reports
  sun_daily_pattern = c(paste0("^(KPI_Daily_TAT_Report ){1}", date_range, ".xls", collapse = "|"), 
                        paste0("^(KPI_Daily_TAT_Report_Updated ){1}", date_range, ".xls", collapse = "|"))
  
  file_list_sun_daily <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = paste0(sun_daily_pattern, collapse = "|"))
  # Read in data reports from possible date range
  scc_list <- lapply(file_list_scc, function(x) read_excel(path = paste0(user_wd, "\\SCC CP Reports\\", x)))
  sun_daily_list <- lapply(file_list_sun_daily, function(x) (read_excel(path = paste0(user_wd, "\\SUN CP Reports\\", x),
                                                                  col_types = c("text", "text", "text", "text", "text", 
                                                                                "text", "text", "text", "text", 
                                                                                "numeric", "numeric", "numeric", "numeric", "numeric",
                                                                                "text", "text", "text", "text", "text", 
                                                                                "text", "text", "text", "text", "text", 
                                                                                "text", "text", "text", "text", "text", 
                                                                                "text", "text", "text", "text", "text", "text"))))
  
  # Create empty list for Sunquest monthly report
  sun_monthly_list <- NULL
}


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


# Custom function for preprocessing daily Sunquest data -----------------
preprocess_sun_daily <- function(raw_sun_daily) {
  
  # Remove duplicates
  raw_sun_daily <- unique(raw_sun_daily)
  
  # Sunquest lookup references ----------------------------------------------
  # Crosswalk labs included and remove out of scope labs
  raw_sun_daily <- left_join(raw_sun_daily, test_code[ , c("Test", "SUN_TestCode", "Division")], 
                             by = c("TestCode" = "SUN_TestCode"))
  raw_sun_daily <- raw_sun_daily %>%
    mutate(TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  # mutate(TestCode = as.factor(TestCode),
  #        Division = as.factor(Division),
  #        TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  
  # Remove out of scope labs
  raw_sun_daily <- raw_sun_daily %>%
    filter(TestIncl)
  
  # Crosswalk units and identify ICUs
  raw_sun_daily <- raw_sun_daily %>%
    mutate(LocandName = paste(LocCode, LocName))
  raw_sun_daily <- left_join(raw_sun_daily, sun_icu[ , c("Concatenate", "ICU")], by = c("LocandName" = "Concatenate"))
  raw_sun_daily[is.na(raw_sun_daily$ICU), "ICU"] <- FALSE
  
  # Crosswalk unit type
  raw_sun_daily <- left_join(raw_sun_daily, sun_setting, by = c("LocType" = "LocType"))
  # Crosswalk site name
  raw_sun_daily <- left_join(raw_sun_daily, mshs_site, by = c("HospCode" = "DataSite"))
  
  # Sunquest data formatting --------------------------------------------
  # raw_sun_daily[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
  #            "LocType", "LocCode", "LocName",
  #            "PhysName", "SHIFT",
  #            "ReceiveTech", "ResultTech", "PerformingLabCode",
  #            "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- lapply(raw_sun_daily[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
  #                                                                                            "LocType", "LocCode", "LocName",
  #                                                                                            "PhysName", "SHIFT",
  #                                                                                            "ReceiveTech", "ResultTech", "PerformingLabCode",
  #                                                                                            "Test", "LocandName", "Setting", "SettingRollUp", "Site")], as.factor)
  
  # Add NA as factor level for SpecimenPriority since many specimens are submitted without a priority
  # raw_sun_daily <- raw_sun_daily %>%
  #   mutate(SpecimenPriority = addNA(SpecimenPriority))
  
  # Fix any timestamps that weren't imported correctly and then format as date/time
  raw_sun_daily[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun_daily[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, str_replace(x, "\\*.*\\*", ""), x))
  
  raw_sun_daily[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun_daily[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")],
                                                                                                      as.POSIXct, tz = "UTC", format = "%m/%d/%Y %H:%M:%S")
  
  # Add a column for Resulted date for later use in repository
  raw_sun_daily <- raw_sun_daily %>%
    mutate(ResultedDate = as.Date(ResultDateTime, format = "%m/%d/%Y"))
  
  
  
  # Sunquest data preprocessing --------------------------------------------------
  # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
  raw_sun_daily <- raw_sun_daily %>%
    mutate(MasterSetting = ifelse(SettingRollUp == "ED", "ED",
                                  ifelse(SettingRollUp == "Amb", "Amb",
                                         ifelse(SettingRollUp == "IP" & ICU == TRUE, "ICU",
                                                ifelse(SettingRollUp == "IP" & ICU == FALSE, "IP Non-ICU", "Other")))),
           DashboardSetting = ifelse(MasterSetting == "ED" | MasterSetting == "ICU", "ED & ICU", MasterSetting))
  
  # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
  raw_sun_daily <- raw_sun_daily %>%
    mutate(AdjPriority = ifelse(MasterSetting != "ED" & MasterSetting != "ICU" & is.na(SpecimenPriority), "Routine",
                                ifelse(MasterSetting == "ED" | MasterSetting == "ICU" | SpecimenPriority == "S", "Stat", "Routine")),
           DashboardPriority = ifelse(tat_targets$Priority[match(Test, tat_targets$Test)] == "All", "All", AdjPriority))
  
  # Calculate turnaround times
  raw_sun_daily <- raw_sun_daily %>%
    mutate(CollectToReceive = ReceiveDateTime - CollectDateTime,
           ReceiveToResult = ResultDateTime - ReceiveDateTime,
           CollectToResult = ResultDateTime - CollectDateTime)
  
  raw_sun_daily[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(raw_sun_daily[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min after specimen received
  # Identify specimens with missing collections times as those with collection time defaulted to order time
  raw_sun_daily <- raw_sun_daily %>%
    mutate(AddOnMaster = ifelse(difftime(OrderDateTime, ReceiveDateTime, units = "mins") > 5, "AddOn", "Original"),
           MissingCollect = ifelse(CollectDateTime == OrderDateTime, TRUE, FALSE))
  
  # Determine target TAT based on test, priority, and patient setting
  raw_sun_daily <- raw_sun_daily %>%
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
  raw_sun_daily <- raw_sun_daily %>%
    mutate(Concate3 = paste(PtNumber, HISOrderNumber, TSTName,
                            CollectDateTime, ReceiveDateTime, ResultDateTime))
  
  raw_sun_daily <- raw_sun_daily[!duplicated(raw_sun_daily$Concate3), ]
  
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
  raw_sun_daily <- raw_sun_daily %>%
    mutate(TATInclude = ifelse(AddOnMaster != "Original" | MasterSetting == "Other" | CollectToResult < 0 | 
                                 ReceiveToResult < 0 | is.na(CollectToResult) | is.na(ReceiveToResult), FALSE, TRUE))
  
  sun_daily_master <- raw_sun_daily[ ,c("LocCode", "LocName", "LocandName",
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
  
  colnames(sun_daily_master) <- c("LocCode", "LocName", "LocConcat",
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
  
  sun_daily_list <- list(raw_sun_daily, sun_daily_master)
  
}

# Custom function for preprocessing monthly Sunquest data -----------------------
preprocess_sun_monthly <- function(raw_sun_monthly) {
  
  # Remove duplicates
  raw_sun_monthly <- unique(raw_sun_monthly)
  
  # Crosswalk labs included and remove out of scope labs
  raw_sun_monthly <- left_join(raw_sun_monthly, test_code[ , c("Test", "SUN_TestCode", "Division")], 
                               by = c("TestCode" = "SUN_TestCode"))
  
  # Set test codes and lab divisions to factors. Determine if test should be included in dashboard
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  # mutate(TestCode = as.factor(TestCode),
  #        Division = as.factor(Division),
  #        TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  
  # Remove tests that are not included in dashboard
  raw_sun_monthly <- raw_sun_monthly %>%
    filter(TestIncl == TRUE)
  
  # Crosswalk units and identify ICUs
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(LocandName = paste(LocCode, LocName))
  raw_sun_monthly <- left_join(raw_sun_monthly, sun_icu[ , c("Concatenate", "ICU")], by = c("LocandName" = "Concatenate"))
  
  raw_sun_monthly[is.na(raw_sun_monthly$ICU), "ICU"] <- FALSE
  # Crosswalk unit type
  raw_sun_monthly <- left_join(raw_sun_monthly, sun_setting, by = c("LocType" = "LocType"))
  # Crosswalk site name
  raw_sun_monthly <- left_join(raw_sun_monthly, mshs_site, by = c("HospCode" = "DataSite"))
  
  # Sunquest data formatting --------------------------------------------
  # raw_sun_monthly[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
  #                    "LocType", "LocCode", "LocName", 
  #                    "PhysName", "SHIFT",
  #                    "ReceiveTech", "ResultTech", "PerformingLabCode",
  #                    "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- 
  #   lapply(raw_sun_monthly[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName", 
  #                             "LocType", "LocCode", "LocName", "PhysName", "SHIFT", "ReceiveTech", 
  #                             "ResultTech", "PerformingLabCode", "Test", "LocandName", "Setting", "SettingRollUp", "Site")], as.factor)
  
  # Add NA as factor level for SpecimenPriority since many specimens are submitted without a priority
  # Add a column for resulted date for later use in repository
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(ResultedDate = as.Date(ResultDateTime, format = "%m/%d/%Y"))
  # mutate(SpecimenPriority = addNA(SpecimenPriority),
  #        ResultedDate = as.Date(ResultDateTime, format = "%m/%d/%Y"))
  
  # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(MasterSetting = ifelse(SettingRollUp == "ED", "ED", 
                                  ifelse(SettingRollUp == "Amb", "Amb", 
                                         ifelse(SettingRollUp == "IP" & ICU == TRUE, "ICU",
                                                ifelse(SettingRollUp == "IP" & ICU == FALSE, "IP Non-ICU", "Other")))),
           DashboardSetting = ifelse(MasterSetting == "ED" | MasterSetting == "ICU", "ED & ICU", MasterSetting))
  
  
  # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(AdjPriority = ifelse(MasterSetting != "ED" & MasterSetting != "ICU" & is.na(SpecimenPriority), "Routine", 
                                ifelse(MasterSetting == "ED" | MasterSetting == "ICU" | SpecimenPriority == "S", "Stat", "Routine")),
           DashboardPriority = ifelse(tat_targets$Priority[match(Test, tat_targets$Test)] == "All", "All", AdjPriority))
  
  # Calculate turnaround times
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(CollectToReceive = ReceiveDateTime - CollectDateTime,
           ReceiveToResult = ResultDateTime - ReceiveDateTime,
           CollectToResult = ResultDateTime - CollectDateTime)
  
  raw_sun_monthly[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(raw_sun_monthly[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min after specimen received
  # Identify specimens with missing collections times as those with collection time defaulted to order time
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(AddOnMaster = ifelse(difftime(OrderDateTime, ReceiveDateTime, units = "mins") > 5, "AddOn", "Original"),
           MissingCollect = CollectDateTime == OrderDateTime)
  
  # Determine target TAT based on test, priority, and patient setting
  raw_sun_monthly <- raw_sun_monthly %>%
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
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(Concate3 = paste(PtNumber, HISOrderNumber, TSTName, CollectDateTime, 
                            ReceiveDateTime, ResultDateTime))
  
  raw_sun_monthly <- raw_sun_monthly[!duplicated(raw_sun_monthly$Concate3), ]
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
  raw_sun_monthly <- raw_sun_monthly %>%
    mutate(TATInclude = ifelse(AddOnMaster != "Original" | MasterSetting == "Other" | CollectToResult < 0 | 
                                 ReceiveToResult < 0 | is.na(CollectToResult) | is.na(ReceiveToResult), FALSE, TRUE))
  
  sun_monthly_master <- raw_sun_monthly[ ,c("LocCode", "LocName", "LocandName", 
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
  
  colnames(sun_monthly_master) <- c("LocCode", "LocName", "LocConcat", 
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
  
  # sun_monthly_master[c("LocConcat", "RequestMD", "WorkShift", "AdjPriority", "AddOnMaster")] <-
  #   lapply(sun_monthly_master[c("LocConcat", "RequestMD", "WorkShift", "AdjPriority", "AddOnMaster")], as.factor)
  # 
  # sun_monthly_master$Site <- factor(sun_monthly_master$Site, levels = site_order)
  # sun_monthly_master$Test <- factor(sun_monthly_master$Test, levels = cp_micro_lab_order)
  # sun_monthly_master$MasterSetting <- factor(sun_monthly_master$MasterSetting, levels = pt_setting_order)
  # sun_monthly_master$DashboardSetting <- factor(sun_monthly_master$DashboardSetting, levels = pt_setting_order2)
  # sun_monthly_master$DashboardPriority <- factor(sun_monthly_master$DashboardPriority, levels = dashboard_priority_order)
  
  sun_monthly_list <- list(raw_sun_monthly, sun_monthly_master)
  
  
}

# Custom function for preprocessing SCC data ---------------------------------
preprocess_scc <- function(raw_scc)  {
  # SCC DATA PROCESSING --------------------------
  # SCC lookup references ----------------------------------------------
  # Crosswalk labs included and remove out of scope labs
  raw_scc <- left_join(raw_scc, test_code[ , c("Test", "SCC_TestID", "Division")], by = c("TEST_ID" = "SCC_TestID"))
  raw_scc <- raw_scc %>%
    mutate(TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  # mutate(TEST_ID = as.factor(TEST_ID),
  #        Division = as.factor(Division),
  #        TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  
  raw_scc <- raw_scc %>%
    filter(TestIncl)
  
  # Crosswalk units and identify ICUs
  raw_scc <- raw_scc %>%
    mutate(WardandName = paste(Ward, WARD_NAME))
  
  raw_scc <- left_join(raw_scc, scc_icu[ , c("Concatenate", "ICU")], by = c("WardandName" = "Concatenate"))
  raw_scc[is.na(raw_scc$ICU), "ICU"] <- FALSE
  
  # Crosswalk unit type
  raw_scc <- left_join(raw_scc, scc_setting, by = c("CLINIC_TYPE" = "Clinic_Type"))
  # Crosswalk site name
  raw_scc <- left_join(raw_scc, mshs_site, by = c("SITE" = "DataSite"))
  
  # SCC data formatting ----------------------------------------------
  # raw_scc[c("Ward", "WARD_NAME",
  #           "REQUESTING_DOC",
  #           "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "Test",
  #           "COLLECT_CENTER_ID", "SITE", "Site",
  #           "CLINIC_TYPE", "Setting", "SettingRollUp")] <- lapply(raw_scc[c("Ward", "WARD_NAME",
  #                                                                           "REQUESTING_DOC",
  #                                                                           "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "Test",
  #                                                                           "COLLECT_CENTER_ID", "SITE", "Site",
  #                                                                           "CLINIC_TYPE", "Setting", "SettingRollUp")], as.factor)
  
  # Fix any timestamps that weren't imported correctly and then format as date/time
  raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")], function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, str_replace(x, "\\*.*\\*", ""), x))
  
  raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")],
                                                                                            as.POSIXct, tz = "UTC", format = "%Y-%m-%d %H:%M:%OS", options(digits.sec = 1))
  
  # Add a column for Resulted date for later use in repository
  raw_scc <- raw_scc %>%
    mutate(ResultedDate = as.Date(VERIFIED_DATE, format = "%m/%d/%Y"))
  
  # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
  raw_scc <- raw_scc %>%
    mutate(MasterSetting = ifelse(CLINIC_TYPE == "E", "ED",
                                  ifelse(CLINIC_TYPE == "O", "Amb",
                                         ifelse(CLINIC_TYPE == "I" & ICU == TRUE, "ICU",
                                                ifelse(CLINIC_TYPE == "I" & ICU != TRUE, "IP Non-ICU", "Other")))),
           DashboardSetting = ifelse(MasterSetting == "ED" | MasterSetting == "ICU", "ED & ICU", MasterSetting))
  
  # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
  raw_scc <- raw_scc %>%
    mutate(AdjPriority = ifelse(MasterSetting == "ED" | MasterSetting == "ICU" | PRIORITY == "S", "Stat", "Routine"),
           DashboardPriority = ifelse(tat_targets$Priority[match(Test, tat_targets$Test)] == "All", "All", AdjPriority))
  
  # Calculate turnaround times
  raw_scc <- raw_scc %>%
    mutate(CollectToReceive = RECEIVE_DATE - COLLECTION_DATE,
           ReceiveToResult = VERIFIED_DATE - RECEIVE_DATE,
           CollectToResult = VERIFIED_DATE - COLLECTION_DATE)
  
  raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- 
    lapply(raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min after specimen received
  # Identify specimens with missing collections times as those with collection time defaulted to receive time
  raw_scc <- raw_scc %>%
    mutate(AddOnMaster = ifelse(difftime(ORDERING_DATE, RECEIVE_DATE, units = "mins") > 5, "AddOn", "Original"),
           MissingCollect = ifelse(CollectToReceive == 0, TRUE, FALSE))
  
  # Determine target TAT based on test, priority, and patient setting
  raw_scc <- raw_scc %>%
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
  raw_scc <- raw_scc %>%
    mutate(Concate3 = paste(LAST_NAME, FIRST_NAME,
                            ORDER_ID, TEST_NAME,
                            COLLECTION_DATE, RECEIVE_DATE, VERIFIED_DATE))
  
  raw_scc <- raw_scc[!duplicated(raw_scc$Concate3), ]
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
  raw_scc <- raw_scc %>%
    mutate(TATInclude = ifelse(AddOnMaster != "Original" | MasterSetting == "Other" | 
                                 CollectToResult < 0 | ReceiveToResult < 0 | is.na(CollectToResult) | 
                                 is.na(ReceiveToResult), FALSE, TRUE))  
  scc_master <- raw_scc[ , c("Ward", "WARD_NAME", "WardandName",
                             "ORDER_ID", "REQUESTING_DOC NAME", "MPI", "WORK SHIFT",
                             "TEST_NAME", "Test", "Division", "PRIORITY",
                             "Site", "ICU", "CLINIC_TYPE",
                             "Setting", "SettingRollUp", "MasterSetting", "DashboardSetting",
                             "AdjPriority", "DashboardPriority",
                             "ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE",
                             "ResultedDate",
                             "CollectToReceive", "ReceiveToResult", "CollectToResult",
                             "AddOnMaster", "MissingCollect",
                             "ReceiveResultTarget", "CollectResultTarget",
                             "ReceiveResultInTarget", "CollectResultInTarget",
                             "TATInclude")]
  colnames(scc_master) <- c("LocCode", "LocName", "LocConcat",
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
  
  
  scc_daily_list <- list(raw_scc, scc_master)
  
}

# Custom function to determine resulted lab date from preprocessed data (SCC data often has a few labs with incorrect result date) --------------------
correct_result_dates <- function(data, number_days) {
  all_resulted_dates_vol <- data %>%
    group_by(ResultDate) %>%
    summarize(VolLabs = n()) %>%
    arrange(desc(VolLabs)) %>%
    ungroup()
  
  
  correct_dates <- all_resulted_dates_vol$ResultDate[1:number_days]
  
  new_data <- data %>%
    filter(ResultDate %in% correct_dates)
  return(new_data)
}


# Add logic for different scenarios where there may be empty data sets ----------------

sun_daily_bind <- NULL

sun_monthly_bind <- NULL

scc_daily_bind <- NULL

# Compile preprocessed Sunquest daily data, if any exists -----------------------------
if (length(sun_daily_list) != 0) {
  # Preprocess Sunquest daily raw data using custom function
  sun_daily_preprocessed <- mapply(preprocess_sun_daily, sun_daily_list)
  
  # Bind daily reports into one data frame
  for (i in seq(from = 2, to = length(sun_daily_preprocessed), by = 2)) {
    sun_daily_bind <- rbind(sun_daily_bind, sun_daily_preprocessed[[i]])
  }
  
} else {
  sun_daily_preprocessed <- NULL
}

# Compile preprocessed Sunquest monthly data, if any exists ----------------------------
if (length(sun_monthly_list) != 0) {
  # Preprocess Sunquest monthly raw data using custom function
  sun_monthly_preprocessed <- mapply(preprocess_sun_monthly, sun_monthly_list)
  
  # Bind monthly reports into one data frame
  for (i in seq(from = 2, to = length(sun_monthly_preprocessed), by = 2)) {
    sun_monthly_bind <- rbind(sun_monthly_bind, sun_monthly_preprocessed[[i]])
  }
  
} else {
  sun_monthly_preprocessed <- NULL
}

# Compile processed SCC daily data, if any exists --------------------------------------
if (length(scc_list) != 0) {
  # Preprocess SCC daily raw data using custom function
  scc_daily_preprocessed <- mapply(preprocess_scc, scc_list)
  
  # Remove any labs with incorrect dates then bind daily reports into one data frame
  for (i in seq(from = 2, to = length(scc_daily_preprocessed), by = 2)) {
    # Remove any labs from incorrect dates
    updated_data <- correct_result_dates(scc_daily_preprocessed[[i]], 1)
    
    # Compile all dates
    scc_daily_bind <- rbind(scc_daily_bind, updated_data)
  }
  
} else {
  scc_daily_preprocessed <- NULL
}

# Bind together all SCC and Sunquest data --------------------------------------
bind_all_data <- rbind(sun_daily_bind, sun_monthly_bind, scc_daily_bind)


# Summarize data for export to historical repo ---------------------------------
scc_sun_all_days_subset <- bind_all_data %>%
  group_by(Site, ResultDate, Test, Division, 
           Setting, SettingRollUp, MasterSetting, DashboardSetting,
           OrderPriority, AdjPriority, DashboardPriority,
           ReceiveResultTarget, CollectResultTarget) %>%
  summarize(TotalResulted = n(), TotalResultedTAT = sum(TATInclude), 
            TotalReceiveResultInTarget = sum(ReceiveResultInTarget[TATInclude == TRUE]), 
            TotalCollectResultInTarget = sum(CollectResultInTarget[TATInclude == TRUE]), 
            TotalAddOnOrder = sum(AddOnMaster == "AddOn"), 
            TotalMissingCollections = sum(MissingCollect)) %>%
  arrange(Site, ResultDate) %>%
  ungroup()

if (initial_run == TRUE) {
  scc_sun_repo <- scc_sun_all_days_subset
} else {
  # Format existing repository for binding
  # existing_repo$Site <- factor(existing_repo$Site, levels = site_order)
  # existing_repo$Test <- factor(existing_repo$Test, levels = cp_micro_lab_order)
  # existing_repo$MasterSetting <- factor(existing_repo$MasterSetting, levels = pt_setting_order)
  # existing_repo$DashboardSetting <- factor(existing_repo$DashboardSetting, levels = pt_setting_order2)
  # existing_repo$DashboardPriority <- factor(existing_repo$DashboardPriority, levels = dashboard_priority_order)
  
  # Convert ResultDate from date-time to date
  existing_repo <- existing_repo %>%
    mutate(ResultDate = date(ResultDate))
  # 
  # Bind new data with repository
  scc_sun_repo <- rbind(existing_repo, scc_sun_all_days_subset)
}

start_date <- format(min(scc_sun_repo$ResultDate), "%m-%d-%y")
end_date <- format(max(scc_sun_repo$ResultDate), "%m-%d-%y")

write_xlsx(scc_sun_repo, path = paste0(user_wd, "\\SCC Sunquest Script Repo", "\\Hist Repo Test ", start_date, " to ", end_date, " Created ", Sys.Date(), ".xlsx"))

