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

# Set working directory
# reference_file <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data\\Code Reference\\Analysis Reference 2020-01-22.xlsx"
user_wd <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data"
user_path <- paste0(user_wd, "\\*.*")
setwd(user_wd)

file_list_scc <- list.files(path = paste0(user_wd, "\\SCC CP Reports"), pattern = "^(Doc){1}.+(2020)\\-(01){1}\\-[0-9]{2}.xlsx")
file_list_sun <- list.files(path = paste0(user_wd, "\\SUN CP Reports"), pattern = "^(KPI_Daily_TAT_Report ){1}(2020)\\-(01){1}\\-[0-9]{2}.xls")

scc_list <- lapply(file_list_scc, function(x) read_excel(path = paste0(user_wd, "\\SCC CP Reports\\", x)))
sun_list <- lapply(file_list_sun, function(x) suppressWarnings(read_excel(path = paste0(user_wd, "\\SUN CP Reports\\", x))))

# for (i in 1:(length(scc_list)-1)) {
#   assign(paste0("scc_raw_data", i), scc_list[[i]])
# }
# 
# for (i in 1:(length(sun_list)-1)) {
#   assign(paste0("sun_raw_data", i), sun_list[[i]])
# }

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

site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")
pt_setting_order <- c("ED", "ICU", "IP Non-ICU", "Amb", "Other")
pt_setting_order2 <- c("ED & ICU", "IP Non-ICU", "Amb", "Other")
dashboard_pt_setting <- c("ED & ICU", "IP Non-ICU", "Amb")

dashboard_priority_order <- c("All", "Stat", "Routine")

# Preprocess data --------------------------------
preprocess_scc_sun <- function(raw_scc, raw_sun)  {
  # SCC DATA PROCESSING --------------------------
  # SCC lookup references ----------------------------------------------
  # Crosswalk labs included and remove out of scope labs
  raw_scc <- left_join(raw_scc, test_code[ , c("Test", "SCC_TestID", "Division")], by = c("TEST_ID" = "SCC_TestID"))
  raw_scc$TEST_ID <- as.factor(raw_scc$TEST_ID)
  raw_scc$Division <- as.factor(raw_scc$Division)
  raw_scc$TestIncl <- ifelse(is.na(raw_scc$Test), FALSE, TRUE)
  raw_scc <- raw_scc[raw_scc$TestIncl == TRUE, ]
  # Crosswalk units and identify ICUs
  raw_scc$WardandName <- paste(raw_scc$Ward, raw_scc$WARD_NAME)
  raw_scc <- left_join(raw_scc, scc_icu[ , c("Concatenate", "ICU")], by = c("WardandName" = "Concatenate"))
  raw_scc[is.na(raw_scc$ICU), "ICU"] <- FALSE
  # Crosswalk unit type
  raw_scc <- left_join(raw_scc, scc_setting, by = c("CLINIC_TYPE" = "Clinic_Type"))
  # Crosswalk site name
  raw_scc <- left_join(raw_scc, mshs_site, by = c("SITE" = "DataSite"))
  
  # SCC data formatting ----------------------------------------------
  raw_scc[c("Ward", "WARD_NAME", 
            "REQUESTING_DOC", 
            "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", "Test",
            "COLLECT_CENTER_ID", "SITE", "Site",
            "CLINIC_TYPE", "Setting", "SettingRollUp")] <- lapply(raw_scc[c("Ward", "WARD_NAME", 
                                                                            "REQUESTING_DOC", 
                                                                            "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", "Test",
                                                                            "COLLECT_CENTER_ID", "SITE", "Site",
                                                                            "CLINIC_TYPE", "Setting", "SettingRollUp")], as.factor)
  
  # Fix any timestamps that weren't imported correctly and then format as date/time
  raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")], function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, str_replace(x, "\\*.*\\*", ""), x))
  
  raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")],
                                                                                            as.POSIXlt, tz = "", format = "%Y-%m-%d %H:%M:%OS", options(digits.sec = 1))
  
  # Add a column for Resulted date for later use in repository
  raw_scc$ResultedDate <- as.Date(raw_scc$VERIFIED_DATE, format, "%m/%d/%Y")
  
  # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
  raw_scc$MasterSetting <- ifelse(raw_scc$CLINIC_TYPE == "E", "ED", 
                                  ifelse(raw_scc$CLINIC_TYPE == "O", "Amb", 
                                         ifelse(raw_scc$CLINIC_TYPE == "I" & raw_scc$ICU == TRUE, "ICU", 
                                                ifelse(raw_scc$CLINIC_TYPE == "I" & raw_scc$ICU != TRUE, "IP Non-ICU", "Other"))))
  raw_scc$DashboardSetting <- ifelse(raw_scc$MasterSetting == "ED" | raw_scc$MasterSetting == "ICU", "ED & ICU", raw_scc$MasterSetting)
  
  
  # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
  raw_scc$AdjPriority <- ifelse(raw_scc$MasterSetting == "ED" | raw_scc$MasterSetting == "ICU" | raw_scc$PRIORITY == "S", "Stat", "Routine")
  raw_scc$DashboardPriority <- ifelse(tat_targets$Priority[match(raw_scc$Test, tat_targets$Test)] == "All", "All", raw_scc$AdjPriority)
  
  # Calculate turnaround times
  raw_scc$CollectToReceive <- raw_scc$RECEIVE_DATE - raw_scc$COLLECTION_DATE
  raw_scc$ReceiveToResult <- raw_scc$VERIFIED_DATE - raw_scc$RECEIVE_DATE
  raw_scc$CollectToResult <- raw_scc$VERIFIED_DATE - raw_scc$COLLECTION_DATE
  raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min after specimen received
  raw_scc$AddOnMaster <- ifelse(difftime(raw_scc$ORDERING_DATE, raw_scc$RECEIVE_DATE, units = "mins") > 5, "AddOn", "Original")
  
  # Identify specimens with missing collections times as those with collection time defaulted to receive time
  raw_scc$MissingCollect <- ifelse(raw_scc$CollectToReceive == 0, TRUE, FALSE)
  
  # Determine target TAT based on test, priority, and patient setting
  raw_scc$Concate1 <- paste(raw_scc$Test, raw_scc$DashboardPriority)
  raw_scc$Concate2 <- paste(raw_scc$Test, raw_scc$DashboardPriority, raw_scc$MasterSetting)
  
  raw_scc$ReceiveResultTarget <- ifelse(!is.na(match(raw_scc$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_scc$Concate2, tat_targets$Concate)], 
                                        ifelse(!is.na(match(raw_scc$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_scc$Concate1, tat_targets$Concate)],
                                               tat_targets$ReceiveToResultTarget[match(raw_scc$Test, tat_targets$Concate)]))
  
  raw_scc$CollectResultTarget <- ifelse(!is.na(match(raw_scc$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_scc$Concate2, tat_targets$Concate)], 
                                        ifelse(!is.na(match(raw_scc$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_scc$Concate1, tat_targets$Concate)],
                                               tat_targets$CollectToResultTarget[match(raw_scc$Test, tat_targets$Concate)]))
  
  raw_scc$ReceiveResultInTarget <- ifelse(raw_scc$ReceiveToResult <= raw_scc$ReceiveResultTarget, TRUE, FALSE)
  raw_scc$CollectResultInTarget <- ifelse(raw_scc$CollectToResult <= raw_scc$CollectResultTarget, TRUE, FALSE)
  
  # Identify and remove duplicate tests
  raw_scc$Concate3 <- paste(raw_scc$LAST_NAME, raw_scc$FIRST_NAME, 
                            raw_scc$ORDER_ID, raw_scc$TEST_NAME,
                            raw_scc$COLLECTION_DATE, raw_scc$RECEIVE_DATE, raw_scc$VERIFIED_DATE)
  
  raw_scc <- raw_scc[!duplicated(raw_scc$Concate3), ]
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
  raw_scc$TATInclude <- ifelse(raw_scc$AddOnMaster != "Original" | raw_scc$MasterSetting == "Other" | raw_scc$CollectToResult < 0 | raw_scc$ReceiveToResult < 0 | is.na(raw_scc$CollectToResult) | is.na(raw_scc$ReceiveToResult), FALSE, TRUE)
  
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
  
  ## SUNQUEST DATA PROCESSING ----------
  # Sunquest lookup references ----------------------------------------------
  # Crosswalk labs included and remove out of scope labs
  raw_sun <- left_join(raw_sun, test_code[ , c("Test", "SUN_TestCode", "Division")], by = c("TestCode" = "SUN_TestCode"))
  raw_sun$TestCode <- as.factor(raw_sun$TestCode)
  raw_sun$Division <- as.factor(raw_sun$Division)
  raw_sun$TestIncl <- ifelse(is.na(raw_sun$Test), FALSE, TRUE)
  raw_sun <- raw_sun[raw_sun$TestIncl == TRUE, ]
  # Crosswalk units and identify ICUs
  raw_sun$LocandName <- paste(raw_sun$LocCode, raw_sun$LocName)
  raw_sun <- left_join(raw_sun, sun_icu[ , c("Concatenate", "ICU")], by = c("LocandName" = "Concatenate"))
  raw_sun[is.na(raw_sun$ICU), "ICU"] <- FALSE
  # Crosswalk unit type
  raw_sun <- left_join(raw_sun, sun_setting, by = c("LocType" = "LocType"))
  # Crosswalk site name
  raw_sun <- left_join(raw_sun, mshs_site, by = c("HospCode" = "DataSite"))
  
  # Sunquest data formatting --------------------------------------------
  raw_sun[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
            "LocType", "LocCode", "LocName", 
            "SpecimenPriority", "PhysName", "SHIFT",
            "ReceiveTech", "ResultTech", "PerformingLabCode",
            "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- lapply(raw_sun[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
                                                                                           "LocType", "LocCode", "LocName", 
                                                                                           "SpecimenPriority", "PhysName", "SHIFT",
                                                                                           "ReceiveTech", "ResultTech", "PerformingLabCode",
                                                                                           "Test", "LocandName", "Setting", "SettingRollUp", "Site")], as.factor)
  
  # Fix any timestamps that weren't imported correctly and then format as date/time
  raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, str_replace(x, "\\*.*\\*", ""), x))
  
  raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], 
                                                                                                as.POSIXlt, tz = "", format = "%m/%d/%Y %H:%M:%S")
  
  # Add a column for Resulted date for later use in repository
  raw_sun$ResultedDate <- as.Date(raw_sun$ResultDateTime, format, "%m/%d/%Y")
  
  
  
  # Sunquest data preprocessing --------------------------------------------------
  # Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
  raw_sun$MasterSetting <- ifelse(raw_sun$SettingRollUp == "ED", "ED", 
                                  ifelse(raw_sun$SettingRollUp == "Amb", "Amb", 
                                         ifelse(raw_sun$SettingRollUp == "IP" & raw_sun$ICU == TRUE, "ICU",
                                                ifelse(raw_sun$SettingRollUp == "IP" & raw_sun$ICU == FALSE, "IP Non-ICU", "Other"))))
  raw_sun$DashboardSetting <- ifelse(raw_sun$MasterSetting == "ED" | raw_sun$MasterSetting == "ICU", "ED & ICU", raw_sun$MasterSetting)
  
  # Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
  raw_sun$AdjPriority <- ifelse(raw_sun$MasterSetting != "ED" & raw_sun$MasterSetting != "ICU" & is.na(raw_sun$SpecimenPriority), "Routine",
                                ifelse(raw_sun$MasterSetting == "ED" | raw_sun$MasterSetting == "ICU" | raw_sun$SpecimenPriority == "S", "Stat", "Routine"))
  raw_sun$DashboardPriority <- ifelse(tat_targets$Priority[match(raw_sun$Test, tat_targets$Test)] == "All", "All", raw_sun$AdjPriority)
  
  # Calculate turnaround times
  raw_sun$CollectToReceive <- raw_sun$ReceiveDateTime - raw_sun$CollectDateTime
  raw_sun$ReceiveToResult <- raw_sun$ResultDateTime - raw_sun$ReceiveDateTime
  raw_sun$CollectToResult <- raw_sun$ResultDateTime - raw_sun$CollectDateTime
  raw_sun[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(raw_sun[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min after specimen received
  raw_sun$AddOnMaster <- ifelse(difftime(raw_sun$OrderDateTime, raw_sun$ReceiveDateTime, units = "mins") > 5, "AddOn", "Original")
  
  # Identify specimens with missing collections times as those with collection time defaulted to order time
  raw_sun$MissingCollect <- ifelse(raw_sun$CollectDateTime == raw_sun$OrderDateTime, TRUE, FALSE)
  
  # Determine target TAT based on test, priority, and patient setting
  raw_sun$Concate1 <- paste(raw_sun$Test, raw_sun$DashboardPriority)
  raw_sun$Concate2 <- paste(raw_sun$Test, raw_sun$DashboardPriority, raw_sun$MasterSetting)
  
  raw_sun$ReceiveResultTarget <- ifelse(!is.na(match(raw_sun$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_sun$Concate2, tat_targets$Concate)], 
                                        ifelse(!is.na(match(raw_sun$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(raw_sun$Concate1, tat_targets$Concate)],
                                               tat_targets$ReceiveToResultTarget[match(raw_sun$Test, tat_targets$Concate)]))
  
  raw_sun$CollectResultTarget <- ifelse(!is.na(match(raw_sun$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_sun$Concate2, tat_targets$Concate)], 
                                        ifelse(!is.na(match(raw_sun$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(raw_sun$Concate1, tat_targets$Concate)],
                                               tat_targets$CollectToResultTarget[match(raw_sun$Test, tat_targets$Concate)]))
  
  raw_sun$ReceiveResultInTarget <- ifelse(raw_sun$ReceiveToResult <= raw_sun$ReceiveResultTarget, TRUE, FALSE)
  raw_sun$CollectResultInTarget <- ifelse(raw_sun$CollectToResult <= raw_sun$CollectResultTarget, TRUE, FALSE)
  
  # Identify and remove duplicate tests
  raw_sun$Concate3 <- paste(raw_sun$PtNumber, 
                            raw_sun$HISOrderNumber, raw_sun$TSTName,
                            raw_sun$CollectDateTime, raw_sun$ReceiveDateTime, raw_sun$ResultDateTime)
  
  raw_sun <- raw_sun[!duplicated(raw_sun$Concate3), ]
  
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
  raw_sun$TATInclude <- ifelse(raw_sun$AddOnMaster != "Original" | raw_sun$MasterSetting == "Other" | raw_sun$CollectToResult < 0 | raw_sun$ReceiveToResult < 0 | is.na(raw_sun$CollectToResult) | is.na(raw_sun$ReceiveToResult), FALSE, TRUE)
  
  sun_master <- raw_sun[ ,c("LocCode", "LocName", "LocandName", 
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
  
  colnames(sun_master) <- c("LocCode", "LocName", "LocConcat", 
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
  
  
  scc_sun_master <- rbind(scc_master, sun_master)
  
  scc_sun_master[c("LocConcat", "RequestMD", "WorkShift", "AdjPriority", "AddOnMaster")] <-
    lapply(scc_sun_master[c("LocConcat", "RequestMD", "WorkShift", "AdjPriority", "AddOnMaster")], as.factor)
  
  scc_sun_master$Site <- factor(scc_sun_master$Site, levels = site_order)
  scc_sun_master$Test <- factor(scc_sun_master$Test, levels = cp_micro_lab_order)
  scc_sun_master$MasterSetting <- factor(scc_sun_master$MasterSetting, levels = pt_setting_order)
  scc_sun_master$DashboardSetting <- factor(scc_sun_master$DashboardSetting, levels = pt_setting_order2)
  scc_sun_master$DashboardPriority <- factor(scc_sun_master$DashboardPriority, levels = dashboard_priority_order)
  
  scc_sun_list <- list(raw_scc, raw_sun, scc_sun_master)
  
}

# Preprocess all SCC and Sunquest files
test <- mapply(preprocess_scc_sun, scc_list, sun_list)

# Bind together scc_sun_master file for each day
bind_test <- NULL

for (i in seq(from = 3, to = length(test), by = 3)) {
  bind_test <- rbind(bind_test, test[[i]])
}

# Remove duplicate entries across days
bind_test <- unique(bind_test)
