---
title: "MSHS Laboratory KPI Dashboard"
output: html_document
---
```{r global_options, echo=FALSE}
knitr::opts_chunk$set(echo=FALSE, warning=FALSE, message=FALSE)
```

```{r include = FALSE}
#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("xlsx")
#install.packages("readxl")
#install.packages("dplyr")
#install.packages("reshape2")
#install.packages("knitr")
#install.packages("kableExtra")
#install.packages("formattable")

#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(dplyr)
library(lubridate)
library(reshape2)
library(knitr)
library(kableExtra)
library(formattable)
```

```{r Determine date, warning = FALSE, message = FALSE, echo = FALSE}
#Clear existing history
rm(list = ls())
#-------------------------------holiday/weekend-------------------------------#

#Determine if yesterday was a holiday/weekend 

#get yesterday's DOW
# Yesterday_Day <- weekdays(as.Date(Sys.Date()-1))
Yest <- as.Date("11/24/2019", format = "%m/%d/%Y")
Yesterday_Day <- weekdays(Yest) #Rename as Yest_DOW
#Change the format for the date into timeDate format to be ready for the next function
Yesterday <- as.timeDate(Yest)

#get yesterday's DOW
Yesterday_Day <- dayOfWeek(Yesterday)

#Excludes Good Friday from the NYSE Holidays
NYSE_Holidays <- as.Date(holidayNYSE(year = c(2019, 2020)))
GoodFriday <- as.Date(GoodFriday())
MSHS_Holiday <- NYSE_Holidays[GoodFriday != NYSE_Holidays]

#This function determines whether yesterday was a holiday/weekend or no
Holiday_Det <- isHoliday(Yesterday, holidays = MSHS_Holiday)

#------------------------------Read Excel sheets------------------------------#
#The if-statement below helps in determining how many excel files are required

#-----------SCC Excel Files-----------#

if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){
  SCC_Sunday <- read_excel(choose.files(caption = "Select SCC Sunday Report"), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(choose.files(caption = "Select SCC Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(rbind(SCC_Holiday_Monday_or_Friday,SCC_Sunday),SCC_Saturday)
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sun")){
  SCC_Sunday <- read_excel(choose.files(caption = "Select SCC Sunday Report"), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(choose.files(caption = "Select SCC Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(SCC_Sunday,SCC_Saturday)
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Mon")|(Yesterday_Day !="Sun"))){
  SCC_Holiday_Weekday <- read_excel(choose.files(caption = "Select SCC Holiday Report"), sheet = 1, col_names = TRUE)
  SCC_Not_Weekday <- SCC_Holiday_Weekday
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
} else {
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
  SCC_Not_Weekday <- NULL
}

#-----------SunQuest Excel Files-----------#

if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){
  SQ_Holiday_Monday_or_Friday <- read_excel(choose.files(caption = "Select SunQuest Holiday Report"), sheet = 1, col_names = TRUE)
  SQ_Sunday <- read_excel(choose.files(caption = "Select SunQuest Sunday Report"), sheet = 1, col_names = TRUE)
  SQ_Saturday <- read_excel(choose.files(caption = "Select SunQuest Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(rbind(SQ_Holiday_Monday_or_Friday ,SQ_Sunday),SQ_Saturday)
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sun")){
  SQ_Sunday <- read_excel(choose.files(caption = "Select SunQuest Sunday Report"), sheet = 1, col_names = TRUE)
  SQ_Saturday <- read_excel(choose.files(caption = "Select SunQuest Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(SQ_Sunday,SQ_Saturday)
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Mon")|(Yesterday_Day !="Sunday"))){
  SQ_Holiday_Weekday <- read_excel(choose.files(caption = "Select SunQuest Holiday Report"), sheet = 1, col_names = TRUE)
  SQ_Not_Weekday <- SQ_Holiday_Weekday
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
} else {
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
  SQ_Not_Weekday <- NULL
}
```

```{r Reference Data, warning = FALSE, message = FALSE, echo = FALSE}
# CP and Micro --------------------------------
# Import analysis reference data starting with test codes for SCC and Sunquest
reference_file <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data\\Code Reference\\Analysis Reference 2019-12-02.xlsx"
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

scc_wday <- SCC_Weekday
sun_wday <- SQ_Weekday

cp_micro_lab_order <- c("Troponin", "Lactate WB", "BUN", "HGB", "PT", "Rapid Flu", "C. diff")

site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")
pt_setting_order <- c("ED", "ICU", "IP Non-ICU", "Amb", "Other")
pt_setting_order2 <- c("ED & ICU", "IP Non-ICU", "Amb", "Other")
dashboard_pt_setting <- c("ED & ICU", "IP Non-ICU", "Amb")

dashboard_priority_order <- c("All", "Stat", "Routine")
```

```{r Preprocessing Function, warning = FALSE, message = FALSE, echo = FALSE}
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
  raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")],
                                                                                             as.POSIXlt, tz = "", format = "%Y-%m-%d %H:%M:%OS", options(digits.sec = 1))
  
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
  
  raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(raw_sun[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], 
                                                                                                 as.POSIXlt, tz = "", format = "%m/%d/%Y %H:%M:%S")
  
  
  
  
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
```

```{r Data Preprocessing and Merging, message = FALSE, warning = FALSE, echo = FALSE}
scc_sun_wday_list <- preprocess_scc_sun(SCC_Weekday, SQ_Weekday)
scc_wday <- scc_sun_wday_list[[1]]
sun_wday <- scc_sun_wday_list[[2]]
scc_sun_wday_master <- scc_sun_wday_list[[3]]

if (is.null(SQ_Not_Weekday) & is.null(SQ_Not_Weekday)) {
  include_not_wday <- FALSE
  scc_sun_not_wday_list <- NULL
  scc_not_wday <- NULL
  sun_not_wday <- NULL
  scc_sun_not_wday_master <- NULL
} else {
  include_not_wday <- TRUE
  scc_sun_not_wday_list <- preprocess_scc_sun(SCC_Not_Weekday, SQ_Not_Weekday)
  scc_not_wday <- scc_sun_not_wday_list[[1]]
  sun_not_wday <- scc_sun_not_wday_list[[2]]
  scc_sun_not_wday_master <- scc_sun_not_wday_list[[3]]
}

resulted_date_wday <- unique(date(scc_sun_wday_master$ResultTime))
resulted_date_wday_header <- format(resulted_date_wday, format = "%a %m/%d/%y")

# Determine resulted date for weekday labs
wday_dates_df <- scc_sun_wday_master[(hour(scc_sun_wday_master$ResultTime) !=0 & minute(scc_sun_wday_master$ResultTime) !=0) & is.na(scc_sun_wday_master$ResultTime) == FALSE, "ResultTime"]
wday_dates_df$dates <- date(wday_dates_df$ResultTime)
wday_result_date <- unique(date(wday_dates_df$ResultTime))
wday_result_date <- format(wday_result_date, format = "%a %m/%d/%y")

if (is.null(scc_sun_not_wday_master) == FALSE) {
  not_wday_dates_df <- scc_sun_not_wday_master[(hour(scc_sun_not_wday_master$ResultTime) !=0 & minute(scc_sun_not_wday_master$ResultTime) !=0) & is.na(scc_sun_not_wday_master$ResultTime) == FALSE, "ResultTime"]
  not_wday_dates_df$dates <- date(not_wday_dates_df$ResultTime)
  not_wday_result_date <- unique(date(not_wday_dates_df$ResultTime))
  wkend_holiday_result_date <- ifelse(length(not_wday_result_date) == 1, format(not_wday_result_date, format = "%a %m/%d/%y"), paste0(format(not_wday_result_date[length(not_wday_result_date)], format = "%a %m/%d/%y"), "-", format(not_wday_result_date[1], format = "%a %m/%d/%y")))
} else {
}

```


```{r Custom Functions for subsetting and summarizing data for dashboards, warning = FALSE, message = FALSE, echo = FALSE}
# # Custom function to clean and format summarized data (used in "subsetting and summarizing function"lab_sub_summarize" function beflore)
# clean_summarized_data <- function(x) {
#   # Replace NaN TAT percentages with NA
#   x[is.finite(x$ReceiveResultPercent) == FALSE, "ReceiveResultPercent"] <- NA
#   x[is.finite(x$CollectResultPercent) == FALSE, "CollectResultPercent"] <- NA
#   # Look up target TAT for sites with 0 resulted labs using tat_targets reference
#   x$Concate1 <- paste(x$Test, x$DashboardPriority)
#   x$Concate2 <- paste(x$Test, x$DashboardPriority, x$DashboardSetting)
#   
#   x$ReceiveResultTarget[is.na(x$ReceiveResultTarget)] <- ifelse(!is.na(match(x$Concate2[is.na(x$ReceiveResultTarget)], tat_targets$Concate)), 
#                                                             tat_targets$ReceiveToResultTarget[match(x$Concate2[is.na(x$ReceiveResultTarget)], tat_targets$Concate)], ifelse(!is.na(match(x$Concate1[is.na(x$ReceiveResultTarget)], tat_targets$Concate)), 
#                                                             tat_targets$ReceiveToResultTarget[match(x$Concate1[is.na(x$ReceiveResultTarget)], tat_targets$Concate)], tat_targets$ReceiveToResultTarget[match(x$Test[is.na(x$ReceiveResultTarget)], tat_targets$Concate)]))
#   x$CollectResultTarget[is.na(x$CollectResultTarget)] <- ifelse(!is.na(match(x$Concate2[is.na(x$CollectResultTarget)], tat_targets$Concate)), 
#                                                             tat_targets$CollectToResultTarget[match(x$Concate2[is.na(x$CollectResultTarget)], tat_targets$Concate)], ifelse(!is.na(match(x$Concate1[is.na(x$CollectResultTarget)], tat_targets$Concate)), 
#                                                             tat_targets$CollectToResultTarget[match(x$Concate1[is.na(x$CollectResultTarget)], tat_targets$Concate)], tat_targets$CollectToResultTarget[match(x$Test[is.na(x$CollectResultTarget)], tat_targets$Concate)]))
#   x$Concate1 <- NULL
#   x$Concate2 <- NULL
#   # # Look up target TAT for sites with 0 resulted labs
#   # dummy_df <- unique(x[is.na(x$ReceiveResultTarget) != TRUE | is.na(x$CollectResultTarget) != TRUE, c(1,3:6)])
#   # x$ReceiveResultTarget <- dummy_df$ReceiveResultTarget[match(paste(x$Test, x$DashboardPriority,  x$DashboardSetting), paste(dummy_df$Test, dummy_df$DashboardPriority, dummy_df$DashboardSetting))]
#   # x$CollectResultTarget <- dummy_df$CollectResultTarget[match(paste(x$Test, x$DashboardPriority,  x$DashboardSetting), paste(dummy_df$Test, dummy_df$DashboardPriority, dummy_df$DashboardSetting))]
#   # Format target TAT for tables
#     x[c("ReceiveResultTarget", "CollectResultTarget")] <- lapply(x[c("ReceiveResultTarget", "CollectResultTarget")], function(y) ifelse(is.na(y), y, paste0("<=", y, " min")))
#   x[c("ReceiveResultPercent", "CollectResultPercent")] <- lapply(x[c("ReceiveResultPercent", "CollectResultPercent")], percent, digits = 0)
#   # Create new column with test and priority
#   x$TestAndPriority <- paste(x$Test, "-", x$DashboardPriority, "Labs")
#   # Remove unused combinations like Routine ED & ICU labs
#   x <- x[!(x$DashboardPriority == "Routine" & x$DashboardSetting == "ED & ICU"), ]
#   # Format TAT percentages based on status definitions and lab division
# 
#   x <- x %>%  mutate(ReceiveResultPercent = cell_spec(ReceiveResultPercent, "html", color = ifelse(is.na(ReceiveResultPercent), "lightgray", ifelse(((ReceiveResultPercent >= 0.95 & (LabDivision == "Chemistry" | LabDivision == "Hematology")) | (ReceiveResultPercent == 1.00 & LabDivision == "Microbiology RRL")), "green", ifelse(((ReceiveResultPercent >=0.8 &  ReceiveResultPercent < 0.95 & (LabDivision == "Chemistry" | LabDivision == "Hematology")) | (ReceiveResultPercent >= 0.90 & ReceiveResultPercent < 1.0 & LabDivision == "Microbiology RRL")), "orange", "red")))))
#   x <- x %>%  mutate(CollectResultPercent = cell_spec(CollectResultPercent, "html", color = ifelse(is.na(CollectResultPercent), "lightgray", ifelse(CollectResultPercent >= 0.95, "green", ifelse(CollectResultPercent >=0.8 &  CollectResultPercent < 0.95, "orange", "red")))))
#   x
# }
# 
# # Custon function for subsetting and summarizing data for each lab division's TAT dashboard
# lab_sub_summarize <- function(x, LabDivision) {
#   lab_df <- x[x$Division == LabDivision & x$TATInclude == TRUE, ]
#   lab_df$Test <- factor(lab_df$Test, unique(test_code$Test[test_code$Division == LabDivision]))
#   lab_df$DashboardSetting <- factor(lab_df$DashboardSetting, dashboard_pt_setting)
#   lab_df$DashboardPriority <- droplevels(lab_df$DashboardPriority)
#   lab_summary <- lab_df %>%
#     group_by(Test, Site, DashboardPriority, DashboardSetting, ReceiveResultTarget, CollectResultTarget, .drop = FALSE) %>%
#   summarize(ResultedVolume = n(), ReceiveResultInTarget = sum(ReceiveResultInTarget), CollectResultInTarget = sum(CollectResultInTarget),
#             ReceiveResultPercent = round(ReceiveResultInTarget/ResultedVolume, digits = 3), CollectResultPercent = round(CollectResultInTarget/ResultedVolume, digits = 3))
#   lab_summary <- clean_summarized_data(lab_summary)
#   lab_dashboard1 <- melt(lab_summary, id.var = c("Test", "Site", "DashboardPriority", "TestAndPriority", "DashboardSetting", "ReceiveResultTarget", "CollectResultTarget"), measure.vars = c("ReceiveResultPercent", "CollectResultPercent"))
#   lab_dashboard2 <- dcast(lab_dashboard1, Test + DashboardPriority + TestAndPriority + DashboardSetting +ReceiveResultTarget + CollectResultTarget ~ variable + Site, value.var = "value")
#   lab_dashboard2 <- lab_dashboard2[ , c(1:3, 5, 4, 7:12, 6, 4, 13:18)]
#     lab_sub_output <- list(lab_df, lab_summary, lab_dashboard1, lab_dashboard2)
# }

# Custome function to subset, summarize, clean, and format data for dashboards for all lab divisions
lab_sub_summarize <- function(x, LabDivision) {
  # Subset data to be included based on lab division and whether or not TAT meets inclusion criteria
  lab_df <- x[x$Division == LabDivision & x$TATInclude == TRUE, ]
  # Update factors
  lab_df$Test <- factor(lab_df$Test, unique(test_code$Test[test_code$Division == LabDivision]))
  lab_df$DashboardSetting <- factor(lab_df$DashboardSetting, dashboard_pt_setting)
  lab_df$DashboardPriority <- droplevels(lab_df$DashboardPriority)
  # Summarize data based on test, site, priority, setting, and TAT targets. Keep levels so each site is maintained for melting/casting later.
  lab_summary <- lab_df %>%
    group_by(Test, Site, DashboardPriority, DashboardSetting, ReceiveResultTarget, CollectResultTarget, .drop = FALSE) %>%
  summarize(ResultedVolume = n(), ReceiveResultInTarget = sum(ReceiveResultInTarget), CollectResultInTarget = sum(CollectResultInTarget),
            ReceiveResultPercent = round(ReceiveResultInTarget/ResultedVolume, digits = 3), CollectResultPercent = round(CollectResultInTarget/ResultedVolume, digits = 3))
  
  # Clean up summarized data
  # Replace NaN TAT percentages with NA
  lab_summary[is.finite(lab_summary$ReceiveResultPercent) == FALSE, "ReceiveResultPercent"] <- NA
  lab_summary[is.finite(lab_summary$CollectResultPercent) == FALSE, "CollectResultPercent"] <- NA
  # Look up target TAT for sites with 0 resulted labs using tat_targets reference
  lab_summary$Concate1 <- paste(lab_summary$Test, lab_summary$DashboardPriority)
  lab_summary$Concate2 <- paste(lab_summary$Test, lab_summary$DashboardPriority, lab_summary$DashboardSetting)
  
  lab_summary$ReceiveResultTarget[is.na(lab_summary$ReceiveResultTarget)] <- ifelse(!is.na(match(lab_summary$Concate2[is.na(lab_summary$ReceiveResultTarget)], tat_targets$Concate)), 
                                                            tat_targets$ReceiveToResultTarget[match(lab_summary$Concate2[is.na(lab_summary$ReceiveResultTarget)], tat_targets$Concate)], ifelse(!is.na(match(lab_summary$Concate1[is.na(lab_summary$ReceiveResultTarget)], tat_targets$Concate)), 
                                                            tat_targets$ReceiveToResultTarget[match(lab_summary$Concate1[is.na(lab_summary$ReceiveResultTarget)], tat_targets$Concate)], tat_targets$ReceiveToResultTarget[match(lab_summary$Test[is.na(lab_summary$ReceiveResultTarget)], tat_targets$Concate)]))
  lab_summary$CollectResultTarget[is.na(lab_summary$CollectResultTarget)] <- ifelse(!is.na(match(lab_summary$Concate2[is.na(lab_summary$CollectResultTarget)], tat_targets$Concate)), 
                                                            tat_targets$CollectToResultTarget[match(lab_summary$Concate2[is.na(lab_summary$CollectResultTarget)], tat_targets$Concate)], ifelse(!is.na(match(lab_summary$Concate1[is.na(lab_summary$CollectResultTarget)], tat_targets$Concate)), 
                                                            tat_targets$CollectToResultTarget[match(lab_summary$Concate1[is.na(lab_summary$CollectResultTarget)], tat_targets$Concate)], tat_targets$CollectToResultTarget[match(lab_summary$Test[is.na(lab_summary$CollectResultTarget)], tat_targets$Concate)]))
  lab_summary$Concate1 <- NULL
  lab_summary$Concate2 <- NULL
  
  # Format target TAT for tables from numbers to "<=X min"
  lab_summary[c("ReceiveResultTarget", "CollectResultTarget")] <- lapply(lab_summary[c("ReceiveResultTarget", "CollectResultTarget")], function(y) ifelse(is.na(y), y, paste0("<=", y, " min")))
  lab_summary[c("ReceiveResultPercent", "CollectResultPercent")] <- lapply(lab_summary[c("ReceiveResultPercent", "CollectResultPercent")], percent, digits = 0)
  # Create new column with test and priority to be used in tables later
  lab_summary$TestAndPriority <- paste(lab_summary$Test, "-", lab_summary$DashboardPriority, "Labs")
  # Remove unused combinations, specifically Routine ED & ICU labs
  lab_summary <- lab_summary[!(lab_summary$DashboardPriority == "Routine" & lab_summary$DashboardSetting == "ED & ICU"), ]
  # Apply conditional color formatting to TAT percentages based on status definitions for each lab division
  lab_summary <- lab_summary %>%  mutate(ReceiveResultPercent = cell_spec(ReceiveResultPercent, "html", color = ifelse(is.na(ReceiveResultPercent), "lightgray", ifelse(((ReceiveResultPercent >= 0.95 & (LabDivision == "Chemistry" | LabDivision == "Hematology")) | (ReceiveResultPercent == 1.00 & LabDivision == "Microbiology RRL")), "green", ifelse(((ReceiveResultPercent >=0.8 &  ReceiveResultPercent < 0.95 & (LabDivision == "Chemistry" | LabDivision == "Hematology")) | (ReceiveResultPercent >= 0.90 & ReceiveResultPercent < 1.0 & LabDivision == "Microbiology RRL")), "orange", "red")))))
  lab_summary <- lab_summary %>%  mutate(CollectResultPercent = cell_spec(CollectResultPercent, "html", color = ifelse(is.na(CollectResultPercent), "lightgray", ifelse(((CollectResultPercent >= 0.95 & (LabDivision == "Chemistry" | LabDivision == "Hematology")) | (CollectResultPercent == 1.00 & LabDivision == "Microbiology RRL")), "green", ifelse(((CollectResultPercent >=0.8 &  CollectResultPercent < 0.95 & (LabDivision == "Chemistry" | LabDivision == "Hematology")) | (CollectResultPercent >= 0.90 & CollectResultPercent < 1.0 & LabDivision == "Microbiology RRL")), "orange", "red")))))

  # Melt summarized data into a "long" dataframe
  lab_dashboard1 <- melt(lab_summary, id.var = c("Test", "Site", "DashboardPriority", "TestAndPriority", "DashboardSetting", "ReceiveResultTarget", "CollectResultTarget"), measure.vars = c("ReceiveResultPercent", "CollectResultPercent"))
  # Cast dataframe into wide format for use in tables later
  lab_dashboard2 <- dcast(lab_dashboard1, Test + DashboardPriority + TestAndPriority + DashboardSetting +ReceiveResultTarget + CollectResultTarget ~ variable + Site, value.var = "value")
  # Rearrange casted dataframe columns based on desired table aesthetics
  lab_dashboard2 <- lab_dashboard2[ , c(1:3, 5, 4, 7:12, 6, 4, 13:18)]
  # Save outputs in a list
  lab_sub_output <- list(lab_df, lab_summary, lab_dashboard1, lab_dashboard2)
  lab_sub_output
}





# Custom function for styling kable tables
chem_hem_micro_kable <- function(data) {
  kable(data, format = "html", escape = FALSE, align = "c", 
        col.names = c("Test & Priority", "Target", "Setting", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL", "Target", "Setting", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")) %>%
  kable_styling(bootstrap_options = "hover", position = "center", font_size = 11) %>%
  column_spec(column = c(1, 9, 17), border_right = "thin solid lightgray") %>%
  add_header_above(c(" ", "Receive to Result Within Target" = (ncol(data)-1)/2, "Collect to Result Within Target" = (ncol(data)-1)/2), background = c("white", "#00AEEF", "#221f72"), color = "white", line = FALSE, font_size = 13) %>%
  column_spec(column = 2:9, background = "#00AEEF", color = "white", include_thead = TRUE) %>%
  column_spec(column = 10:17, background = "#221f72", color = "white", include_thead = TRUE) %>%
  column_spec(column = 2:17, background = "inherit", color = "inherit") %>%
  column_spec(column = 1, width_min = "125px", include_thead = TRUE) %>%
  column_spec(column = c(3, 11), width_min = "100px", include_thead = TRUE) %>%
  row_spec(row = 0, font_size = 13) %>%
  collapse_rows(columns = c(1, 2, 10))
}


```

```{r Subset and format data for each lab divisions weekday dashboard, warning = FALSE, message = FALSE, echo = FALSE}

# Chemistry
chem_sub_output <-lab_sub_summarize(scc_sun_wday_master, LabDivision = "Chemistry")
chem_subset <- chem_sub_output[[1]]
chemistry_summary <- chem_sub_output[[2]]
chemistry_dashboard1 <- chem_sub_output[[3]]
chemistry_dashboard2 <- chem_sub_output[[4]]

# Chemistry: Manually remove unused combinations
chemistry_dashboard2 <- chemistry_dashboard2[!(((chemistry_dashboard2$Test == "Troponin" | chemistry_dashboard2$Test == "Lactate WB") & (chemistry_dashboard2$DashboardPriority != "All" | chemistry_dashboard2$DashboardSetting == "Amb")) | (chemistry_dashboard2$Test == "BUN" & chemistry_dashboard2$DashboardPriority == "All")), ]
row.names(chemistry_dashboard2) <- 1:nrow(chemistry_dashboard2)

chem_table <- chemistry_dashboard2[ , 3:ncol(chemistry_dashboard2)]

# Hematology
hem_sub_output <-lab_sub_summarize(scc_sun_wday_master, LabDivision = "Hematology")
hem_subset <- hem_sub_output[[1]]
hematology_summary <- hem_sub_output[[2]]
hematology_dashboard1 <- hem_sub_output[[3]]
hematology_dashboard2 <- hem_sub_output[[4]]

hem_table<- hematology_dashboard2[ , 3:ncol(hematology_dashboard2)]

# Microbiology RRL
micro_sub_output <-lab_sub_summarize(scc_sun_wday_master, LabDivision = "Microbiology RRL")
micro_subset <- micro_sub_output[[1]]
micro_summary <- micro_sub_output[[2]]
micro_dashboard1 <- micro_sub_output[[3]]
micro_dashboard2 <- micro_sub_output[[4]]

# Microbiology RRL: Manually remove unused combinations
micro_dashboard2 <- micro_dashboard2[!(micro_dashboard2$Test == "C. diff" & micro_dashboard2$DashboardSetting == "Amb"), ]
row.names(micro_dashboard2) <- 1:nrow(micro_dashboard2)

micro_table <- micro_dashboard2[ , 3:ncol(micro_dashboard2)]
```

#### *Chemistry KPI (Labs Resulted on `r wday_result_date`)*
<h5>Status Definitions: <span style = "color:red">Red:</span> <80%, 
<span style = "color:orange">Yellow:</span> >=80% & <95%,
<span style = "color:green">Green:</span> >=95%</h5>
```{r Chemistry dashboard table, warning = FALSE, message = FALSE, echo = FALSE}
chem_hem_micro_kable(chem_table)
```

#### *Hematology KPI (Labs Resulted on `r wday_result_date`)*
<h5>Status Definitions: <span style = "color:red">Red:</span> <80%, 
<span style = "color:orange">Yellow:</span> >=80% & <95%,
<span style = "color:green">Green:</span> >=95%</h5>
```{r Hematology dashboard table, warning = FALSE, message = FALSE, echo = FALSE}
chem_hem_micro_kable(hem_table)
```

```{r Microbiology RRL volume and TAT table creation, warning = FALSE, message = FALSE, echo = FALSE}
# chem_hem_micro_kable(micro_table)

micro_volume_dashboard1 <- melt(micro_summary, id.var = c("Test", "Site", "DashboardPriority", "TestAndPriority", "DashboardSetting", "ReceiveResultTarget", "CollectResultTarget"), measure.vars = "ResultedVolume")

# Create a table with micro volume with similar formatting to TAT tables to later combine with TAT
micro_volume_dashboard2 <- dcast(micro_volume_dashboard1, Test + DashboardPriority + TestAndPriority + DashboardSetting +ReceiveResultTarget + CollectResultTarget ~ variable + Site, value.var = "value") 

micro_volume_dashboard2[c("ReceiveResultTarget", "CollectResultTarget")] <- "Resulted Volume"
original_length <- ncol(micro_volume_dashboard2)

micro_volume_dashboard2 <- micro_volume_dashboard2[ ,c(1:ncol(micro_volume_dashboard2), (ncol(micro_volume_dashboard2)-(6-1)):ncol(micro_volume_dashboard2))]

micro_volume_dashboard2 <- micro_volume_dashboard2[ , c(1:3, 5, 4, 7:12, 6, 4, 13:18)]
colnames(micro_volume_dashboard2) <- colnames(micro_dashboard2)

micro_volume_dashboard2[ , (original_length):ncol(micro_volume_dashboard2)] <- " "

# Combine TAT table and volume table
micro_tat_vol_comb <- rbind(micro_dashboard2, micro_volume_dashboard2)

micro_tat_vol_comb <- micro_tat_vol_comb[order(micro_tat_vol_comb$Test, micro_tat_vol_comb$ReceiveResultTarget), ]
row.names(micro_tat_vol_comb) <- 1:nrow(micro_tat_vol_comb)

micro_tat_vol_table <- micro_tat_vol_comb[ , 3:ncol(micro_tat_vol_comb)]
```

#### *Microbiology RRL KPI (Labs Resulted on `r wday_result_date`)*
<h5>Status Definitions: <span style = "color:red">Red:</span> <90%, 
<span style = "color:orange">Yellow:</span> >=90% & <100%,
<span style = "color:green">Green:</span> =100%</h5>
```{r Create Micro RRL TAT and Volume Table, warning = FALSE, message = FALSE, echo = FALSE}
chem_hem_micro_kable(micro_tat_vol_table)
```

#### *Missing Collection Times and Add On Order Volume (Labs Resulted on `r wday_result_date`)*
<h5>Status Definitions: <span style = "color:red">Red:</span> >15%, 
<span style = "color:orange">Yellow:</span> <=15% & >5%,
<span style = "color:green">Green:</span> <=5%</h5>
```{r Missing Collection Times Tables, warning = FALSE, message = FALSE, echo = FALSE}
# # Determine percentage of labs with missing collection times -------------------------------
missing_collect <- scc_sun_wday_master[scc_sun_wday_master$TATInclude == TRUE, ] %>%
  group_by(Site, .drop = FALSE) %>%
  summarize(ResultedVolume = n(), MissingCollection = sum(MissingCollect, na.rm = TRUE), Percent = MissingCollection/ResultedVolume)

# Format percentage of labs with missing collection times as percentage
missing_collect$Percent <- percent(missing_collect$Percent, digits = 0)

# Apply conditional formatting to percentage of labs with missing collection times
missing_collect <- missing_collect %>%
  mutate(Percent = cell_spec(Percent, "html", color = ifelse(is.na(Percent), "grey",
                                                         ifelse(Percent <= 0.05, "green",
                                                                ifelse(Percent >0.05 &  Percent <= 0.15, "orange", "red")))))

missing_collect_table <- dcast(missing_collect, "Percentage of Specimens" ~ Site, value.var = "Percent")

missing_collect_table %>%
  kable(format = "html", escape = FALSE, align = "c",  
        col.names = c("Site", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")) %>%
  kable_styling(bootstrap = "hover", position = "float_left", font_size = 11, full_width = FALSE) %>%
  add_header_above(c(" " = 1, "Percentage of Labs with Missing Collect Times" = ncol(missing_collect_table)-1), background = c("white", "#00AEEF"), color = "white", line = FALSE, font_size = 13) %>%
  column_spec(column = c(1, ncol(missing_collect_table)), border_right = "thin solid lightgray") %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "#00AEEF", color = "white", include_thead = TRUE) %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "inherit", color = "inherit", width_max = 0.15) %>%
  row_spec(row = 0, font_size =13)


add_on_volume <- scc_sun_wday_master %>%
  group_by(Test, Site, .drop = FALSE) %>%
  summarize(AddOnVolume = sum(AddOnMaster == "AddOn", na.rm = TRUE))

add_on_table <- dcast(add_on_volume, Test ~ Site, value.var = "AddOnVolume")

add_on_table %>%
  kable(format = "html", escape = FALSE, align = "c", 
        col.names = c("Test", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL"), color = "gray") %>%
  kable_styling(bootstrap = "hover", position = "right", font_size = 11, full_width = FALSE) %>%
  add_header_above(c(" " = 1, "Volume of Add On Labs" = ncol(missing_collect_table)-1), background = c("white", "#00AEEF"), color = "white", line = FALSE, font_size = 13) %>%
  column_spec(column = c(1, ncol(missing_collect_table)), border_right = "thin solid lightgray") %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "#00AEEF", color = "white", include_thead = TRUE) %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "inherit", color = "inherit") %>%
  row_spec(row = 0, font_size =13)


```

```{r Add On Order Volume, warning = FALSE, message = FALSE, echo = FALSE}
# # Determine number of add-on orders stratified by test and site ------------------------------
# add_on_volume <- scc_sun_wday_master %>%
#   group_by(Test, Site, .drop = FALSE) %>%
#   summarize(AddOnVolume = sum(AddOnMaster == "AddOn", na.rm = TRUE))
# 
# add_on_table <- dcast(add_on_volume, Test ~ Site, value.var = "AddOnVolume")
# 
# add_on_table %>%
#   kable(format = "html", escape = FALSE, align = "c", 
#         col.names = c("Test", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL"), color = "gray") %>%
#   kable_styling(bootstrap = "hover", position = "center", font_size = 11, full_width = FALSE) %>%
#   add_header_above(c(" " = 1, "Volume of Add On Labs" = ncol(missing_collect_table)-1), background = c("white", "#00AEEF"), color = "white", line = FALSE, font_size = 13) %>%
#   column_spec(column = c(1, ncol(missing_collect_table)), border_right = "thin solid lightgray") %>%
#   column_spec(column = c(2:ncol(missing_collect_table)), background = "#00AEEF", color = "white", include_thead = TRUE) %>%
#   column_spec(column = c(2:ncol(missing_collect_table)), background = "inherit", color = "inherit") %>%
#   row_spec(row = 0, font_size =13)
```


```{r Conditional Block for Weekends, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}

# Chemistry
chem_sub_output_not_wday <-lab_sub_summarize(scc_sun_not_wday_master, LabDivision = "Chemistry")
chem_subset_not_wday <- chem_sub_output_not_wday[[1]]
chemistry_summary_not_wday <- chem_sub_output_not_wday[[2]]
chemistry_dashboard1_not_wday <- chem_sub_output_not_wday[[3]]
chemistry_dashboard2_not_wday <- chem_sub_output_not_wday[[4]]

# Chemistry: Manually remove unused combinations
chemistry_dashboard2_not_wday <- chemistry_dashboard2_not_wday[!(((chemistry_dashboard2_not_wday$Test == "Troponin" | chemistry_dashboard2_not_wday$Test == "Lactate WB") & (chemistry_dashboard2_not_wday$DashboardPriority != "All" | chemistry_dashboard2_not_wday$DashboardSetting == "Amb")) | (chemistry_dashboard2_not_wday$Test == "BUN" & chemistry_dashboard2_not_wday$DashboardPriority == "All")), ]
row.names(chemistry_dashboard2_not_wday) <- 1:nrow(chemistry_dashboard2_not_wday)

chem_table_not_wday <- chemistry_dashboard2_not_wday[ , 3:ncol(chemistry_dashboard2_not_wday)]

# Hematology
hem_sub_output_not_wday <-lab_sub_summarize(scc_sun_not_wday_master, LabDivision = "Hematology")
hem_subset_not_wday <- hem_sub_output_not_wday[[1]]
hematology_summary_not_wday <- hem_sub_output_not_wday[[2]]
hematology_dashboard1_not_wday <- hem_sub_output_not_wday[[3]]
hematology_dashboard2_not_wday <- hem_sub_output_not_wday[[4]]

hem_table_not_wday <- hematology_dashboard2_not_wday[ , 3:ncol(hematology_dashboard2_not_wday)]

# Microbiology RRL
micro_sub_output_not_wday <-lab_sub_summarize(scc_sun_not_wday_master, LabDivision = "Microbiology RRL")
micro_subset_not_wday <- micro_sub_output_not_wday[[1]]
micro_summary_not_wday <- micro_sub_output_not_wday[[2]]
micro_dashboard1_not_wday <- micro_sub_output_not_wday[[3]]
micro_dashboard2_not_wday <- micro_sub_output_not_wday[[4]]

# Microbiology RRL: Manually remove unused combinations
micro_dashboard2_not_wday <- micro_dashboard2_not_wday[!(micro_dashboard2_not_wday$Test == "C. diff" & micro_dashboard2_not_wday$DashboardSetting == "Amb"), ]
row.names(micro_dashboard2_not_wday) <- 1:nrow(micro_dashboard2_not_wday)

micro_table_not_wday <- micro_dashboard2_not_wday[ , 3:ncol(micro_dashboard2_not_wday)]
```


```{r Chemistry dashboard table for weekends and holidays, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}
asis_output(paste('<h4> *Chemistry KPI (Labs Resulted on', wkend_holiday_result_date, ')*</h4>'))
asis_output('<h5>Status Definitions: <span style = "color:red">Red:</span> <80%, <span style = "color:orange">Yellow:</span> >=80% & <95%, <span style = "color:green">Green:</span> >=95%</h5>')
chem_hem_micro_kable(chem_table_not_wday)
```


```{r Hematology dashboard table for weekends/holidays, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}
asis_output(paste('<h4> *Hematology KPI (Labs Resulted on', wkend_holiday_result_date, ')*</h4>'))
asis_output('<h5>Status Definitions: <span style = "color:red">Red:</span> <80%, <span style = "color:orange">Yellow:</span> >=80% & <95%, <span style = "color:green">Green:</span> >=95%</h5>')
chem_hem_micro_kable(hem_table_not_wday)
```

```{r Microbiology RRL volume and TAT table creation for weekends/holidays, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}

micro_volume_dashboard1_not_wday <- melt(micro_summary_not_wday, id.var = c("Test", "Site", "DashboardPriority", "TestAndPriority", "DashboardSetting", "ReceiveResultTarget", "CollectResultTarget"), measure.vars = "ResultedVolume")

# Create a table with micro volume with similar formatting to TAT tables to later combine with TAT
micro_volume_dashboard2_not_wday <- dcast(micro_volume_dashboard1_not_wday, Test + DashboardPriority + TestAndPriority + DashboardSetting +ReceiveResultTarget + CollectResultTarget ~ variable + Site, value.var = "value") 

micro_volume_dashboard2_not_wday[c("ReceiveResultTarget", "CollectResultTarget")] <- "Resulted Volume"
original_length_not_wday <- ncol(micro_volume_dashboard2_not_wday)

micro_volume_dashboard2_not_wday <- micro_volume_dashboard2_not_wday[ ,c(1:ncol(micro_volume_dashboard2_not_wday), (ncol(micro_volume_dashboard2_not_wday)-(6-1)):ncol(micro_volume_dashboard2_not_wday))]

micro_volume_dashboard2_not_wday <- micro_volume_dashboard2_not_wday[ , c(1:3, 5, 4, 7:12, 6, 4, 13:18)]
colnames(micro_volume_dashboard2_not_wday) <- colnames(micro_dashboard2_not_wday)

micro_volume_dashboard2_not_wday[ , (original_length_not_wday):ncol(micro_volume_dashboard2_not_wday)] <- " "

# Combine TAT table and volume table
micro_tat_vol_comb_not_wday <- rbind(micro_dashboard2_not_wday, micro_volume_dashboard2_not_wday)

micro_tat_vol_comb_not_wday <- micro_tat_vol_comb_not_wday[order(micro_tat_vol_comb_not_wday$Test, micro_tat_vol_comb_not_wday$ReceiveResultTarget), ]
row.names(micro_tat_vol_comb_not_wday) <- 1:nrow(micro_tat_vol_comb_not_wday)

micro_tat_vol_table_not_wday <- micro_tat_vol_comb_not_wday[ , 3:ncol(micro_tat_vol_comb_not_wday)]
```

```{r Microbiology RRL dashboard table for weekends/holidays, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}
asis_output(paste('<h4> *Microbiology RRL KPI (Labs Resulted on ', wkend_holiday_result_date, ')*</h4>'))
asis_output('<h5>Status Definitions: <span style = "color:red">Red:</span> <90%, <span style = "color:orange">Yellow:</span> >=90% & <100%, <span style = "color:green">Green:</span> =100%</h5>')
chem_hem_micro_kable(micro_tat_vol_table_not_wday)
```

```{r Missing Collection Times Tables for weekends/holidays, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}
asis_output(paste('<h4>*Missing Collection Times and Add On Order Volume (Labs Resulted on ', wkend_holiday_result_date, ')*</h4>'))
asis_output('<h5>Status Definitions: <span style = "color:red">Red:</span> >15%, <span style = "color:orange">Yellow:</span> <=15% & >5%, <span style = "color:green">Green:</span> <=5%</h5>')

# Determine percentage of labs with missing collection times -------------------------------
missing_collect_not_wday <- scc_sun_not_wday_master[scc_sun_not_wday_master$TATInclude == TRUE, ] %>%
  group_by(Site, .drop = FALSE) %>%
  summarize(ResultedVolume = n(), MissingCollection = sum(MissingCollect, na.rm = TRUE), Percent = MissingCollection/ResultedVolume)

# Format percentage of labs with missing collection times as percentage
missing_collect_not_wday$Percent <- percent(missing_collect_not_wday$Percent, digits = 0)

# Apply conditional formatting to percentage of labs with missing collection times
missing_collect_not_wday <- missing_collect_not_wday %>%
  mutate(Percent = cell_spec(Percent, "html", color = ifelse(is.na(Percent), "grey",
                                                         ifelse(Percent <= 0.05, "green",
                                                                ifelse(Percent >0.05 &  Percent <= 0.15, "orange", "red")))))

missing_collect_table_not_wday <- dcast(missing_collect_not_wday, "Percentage of Specimens" ~ Site, value.var = "Percent")

missing_collect_table_not_wday %>%
  kable(format = "html", escape = FALSE, align = "c",  
        col.names = c("Site", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")) %>%
  kable_styling(bootstrap = "hover", position = "float_left", font_size = 11, full_width = FALSE) %>%
  add_header_above(c(" " = 1, "Percentage of Labs with Missing Collect Times" = ncol(missing_collect_table)-1), background = c("white", "#00AEEF"), color = "white", line = FALSE, font_size = 13) %>%
  column_spec(column = c(1, ncol(missing_collect_table)), border_right = "thin solid lightgray") %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "#00AEEF", color = "white", include_thead = TRUE) %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "inherit", color = "inherit", width_max = 0.15) %>%
  row_spec(row = 0, font_size =13)

add_on_volume_not_wday <- scc_sun_not_wday_master %>%
  group_by(Test, Site, .drop = FALSE) %>%
  summarize(AddOnVolume = sum(AddOnMaster == "AddOn", na.rm = TRUE))

add_on_table_not_wday <- dcast(add_on_volume_not_wday, Test ~ Site, value.var = "AddOnVolume")

add_on_table_not_wday %>%
  kable(format = "html", escape = FALSE, align = "c",
        col.names = c("Test", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL"), color = "gray") %>%
  kable_styling(bootstrap = "hover", position = "right", font_size = 11, full_width = FALSE) %>%
  add_header_above(c(" " = 1, "Volume of Add On Labs" = ncol(missing_collect_table)-1), background = c("white", "#00AEEF"), color = "white", line = FALSE, font_size = 13) %>%
  column_spec(column = c(1, ncol(missing_collect_table)), border_right = "thin solid lightgray") %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "#00AEEF", color = "white", include_thead = TRUE) %>%
  column_spec(column = c(2:ncol(missing_collect_table)), background = "inherit", color = "inherit") %>%
  row_spec(row = 0, font_size =13)
```

```{r Add On Order Volume for weekends/holidays, eval = include_not_wday, warning = FALSE, message = FALSE, echo = FALSE}
# Determine number of add-on orders stratified by test and site ------------------------------
# add_on_volume_not_wday <- scc_sun_not_wday_master %>%
#   group_by(Test, Site, .drop = FALSE) %>%
#   summarize(AddOnVolume = sum(AddOnMaster == "AddOn", na.rm = TRUE))
# 
# add_on_table_not_wday <- dcast(add_on_volume_not_wday, Test ~ Site, value.var = "AddOnVolume")
# 
# print("Add On Volume for Weekends/Holidays")
# add_on_table_not_wday %>%
#   kable(format = "html", escape = FALSE, align = "c",
#         col.names = c("Test", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL"), color = "gray") %>%
#   kable_styling(bootstrap = "hover", position = "center", font_size = 11, full_width = FALSE) %>%
#   add_header_above(c(" " = 1, "Volume of Add On Labs" = ncol(missing_collect_table)-1), background = c("white", "#00AEEF"), color = "white", line = FALSE, font_size = 13) %>%
#   column_spec(column = c(1, ncol(missing_collect_table)), border_right = "thin solid lightgray") %>%
#   column_spec(column = c(2:ncol(missing_collect_table)), background = "#00AEEF", color = "white", include_thead = TRUE) %>%
#   column_spec(column = c(2:ncol(missing_collect_table)), background = "inherit", color = "inherit") %>%
#   row_spec(row = 0, font_size =13)
```

