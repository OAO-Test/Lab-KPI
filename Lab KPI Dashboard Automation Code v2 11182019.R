
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat and Kate Nevin

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("readxl")
#install.packages("dplyr")
#install.packages("reshape2")
#install.packages("bizdays")
#install.packages("rmarkdown")
#install.packages("tinytex")

#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(dplyr)
library(lubridate)
library(reshape2)
library(bizdays)
library(dplyr)
library(reshape2)
library(rmarkdown)
library(tinytex)
#-------------------------------holiday/weekend-------------------------------#

#Determine if yesterday was a holiday/weekend 

#get yesterday's DOW
# Yesterday_Day <- weekdays(as.Date(Sys.Date()-1))
Yest <- as.Date("11/25/2019", format = "%m/%d/%Y")
Yesterday_Day <- weekdays(Yest) #Rename as Yest_DOW
#Change the format for the date into timeDate format to be ready for the next function
Yesterday <- as.timeDate(Yest)

#get yesterday's DOW
Yesterday_Day <- dayOfWeek(Yesterday)

#Excludes Good Friday from the NYSE Holidays
NYSE_Holidays <- as.Date(holidayNYSE())
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
  SSC_Not_Weekday <- NULL
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

# Read in and analyze Powerpath data
#-----------PowerPath Excel Files-----------#
#For the powerpath files the read excel is starting from the second row
#Also I made sure to remove the last line
if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){
  PP_Holiday_Monday_or_Friday <- read_excel(choose.files(caption = "Select PowerPath Holiday Report"), skip = 1, 1)
  PP_Holiday_Monday_or_Friday  <- PP_Holiday_Monday[-nrow(PP_Holiday_Monday_or_Friday),]
  PP_Sunday <- read_excel(choose.files(caption = "Select PowerPath Sunday Report"), skip = 1, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read_excel(choose.files(caption = "Select PowerPath Saturday Report"), skip = 1, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- data.frame(rbind(rbind(PP_Holiday_Monday_or_Friday ,PP_Sunday),PP_Saturday),stringsAsFactors = FALSE)
  PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),],stringsAsFactors = FALSE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sun")){
  PP_Sunday <- read_excel(choose.files(caption = "Select PowerPath Sunday Report"), skip = 1, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read_excel(choose.files(caption = "Select PowerPath Saturday Report"), skip = 1, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- data.frame(rbind(PP_Sunday,PP_Saturday),stringsAsFactors = FALSE)
  PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Mon")|(Yesterday_Day !="Sun"))){
  PP_Holiday_Weekday <- read_excel(choose.files(caption = "Select PowerPath Holiday Report"), skip = 1, 1)
  PP_Holiday_Weekday <- PP_Holiday_Weekday[-nrow(PP_Holiday_Weekday),]
  PP_Not_Weekday <- data.frame(PP_Holiday_Weekday, stringsAsFactors = FALSE)
  PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
} else {
  PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
  PP_Not_Weekday <- NULL
}

#-----------Cytology Backlog Excel Files-----------#
#For the backlog files the read excel is starting from the second row
#Also I made sure to remove the last line

Cytology_Backlog <- read_excel(choose.files(caption = "Select Cytology Backlog Report"), skip = 1, 1)
Cytology_Backlog <- data.frame(Cytology_Backlog[-nrow(Cytology_Backlog),], stringsAsFactors = FALSE)

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

# cp_lab_order <- c("Troponin", "Lactate WB", "BUN", "HGB", "PT")
# chem_order <- c("Troponin", "Lactate WB", "BUN")
# hem_order <- c("HGB", "PT")
# micro_order <- c("Rapid Flu", "C. diff")
site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")
pt_setting_order <- c("ED", "ICU", "IP Non-ICU", "Amb", "Other")
pt_setting_order2 <- c("ED & ICU", "IP Non-ICU", "Amb", "Other")

dashboard_priority_order <- c("AllSpecimen", "Stat", "Routine")

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
  raw_scc$DashboardPriority <- ifelse(tat_targets$Priority[match(raw_scc$Test, tat_targets$Test)] == "All", "AllSpecimen", raw_scc$AdjPriority)
  
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
  raw_sun$DashboardPriority <- ifelse(tat_targets$Priority[match(raw_sun$Test, tat_targets$Test)] == "All", "AllSpecimen", raw_sun$AdjPriority)
  
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

scc_sun_wday_list <- preprocess_scc_sun(SCC_Weekday, SQ_Weekday)
scc_wday <- scc_sun_wday_list[[1]]
sun_wday <- scc_sun_wday_list[[2]]
scc_sun_wday_master <- scc_sun_wday_list[[3]]




#------------------------------Data Pre-Processing------------------------------#

#Using Rev Center to determine patient setting
Patient_Setting <- data.frame(read_excel(choose.files(caption = "Select Patient Setting Dataframe"), sheet = "final"), stringsAsFactors = FALSE)

#vlookup the Rev_Center and its corresponding patient setting for the PowerPath Data
PP_Weekday_PS <- merge(x=PP_Weekday, y=Patient_Setting, all.x = TRUE ) 
PP_Not_Weekday_PS <- merge(x=PP_Not_Weekday, y=Patient_Setting, all.x = TRUE )

#Set up a default calendar for TAT calculations
create.calendar("MSHS_working_days", MSHS_Holiday, weekdays=c("saturday","sunday"))
bizdays.options$set(default.calendar="MSHS_working_days")

#Cytology
#Keep the cyto gyn and cyto non-gyn

Cytology_Weekday <- PP_Weekday_PS[which(PP_Weekday_PS$spec_group=="CYTO NONGYN" | PP_Weekday_PS$spec_group=="CYTO GYN"),]
Cytology_Not_Weekday <- PP_Not_Weekday_PS[which(PP_Not_Weekday_PS$spec_group=="CYTO NONGYN" | PP_Not_Weekday_PS$spec_group=="CYTO GYN"),]

#If there are any N/A in the dates, remove that sample
Cytology_Weekday <- Cytology_Weekday[!(is.na(Cytology_Weekday$Case_created_date)|is.na(Cytology_Weekday$Collection_Date) | is.na(Cytology_Weekday$Received_Date) | is.na(Cytology_Weekday$signed_out_date)),]
Cytology_Not_Weekday <- Cytology_Not_Weekday[!(is.na(Cytology_Not_Weekday$Case_created_date)|is.na(Cytology_Not_Weekday$Collection_Date) | is.na(Cytology_Not_Weekday$Received_Date) | is.na(Cytology_Not_Weekday$signed_out_date)),]

#Change all Dates into POSIXct format to start the calculations

Cytology_Weekday$Case_created_date <- as.POSIXct(Cytology_Weekday$Case_created_date,format='%m/%d/%y %I:%M %p')
Cytology_Weekday$Collection_Date <- as.POSIXct(Cytology_Weekday$Collection_Date,format='%m/%d/%y %I:%M %p')
Cytology_Weekday$Received_Date <- as.POSIXct(Cytology_Weekday$Received_Date,format='%m/%d/%y %I:%M %p')
Cytology_Weekday$signed_out_date <- as.POSIXct(Cytology_Weekday$signed_out_date,format='%m/%d/%y %I:%M %p')

if (is.null(PP_Not_Weekday)){
  Cytology_Not_Weekday <- Cytology_Not_Weekday
} else {
  Cytology_Not_Weekday$Case_created_date <- as.POSIXct(Cytology_Not_Weekday$Case_created_date,format='%m/%d/%y %I:%M %p')
  Cytology_Not_Weekday$Collection_Date <- as.POSIXct(Cytology_Not_Weekday$Collection_Date,format='%m/%d/%y %I:%M %p')
  Cytology_Not_Weekday$Received_Date <- as.POSIXct(Cytology_Not_Weekday$Received_Date,format='%m/%d/%y %I:%M %p')
  Cytology_Not_Weekday$signed_out_date <- as.POSIXct(Cytology_Not_Weekday$signed_out_date,format='%m/%d/%y %I:%M %p')
}


#add columns for calculations: collection to signed out and received to signed out
#collection to signed out
#All days in the calendar
Cytology_Weekday$Collection_to_signed_out <- as.numeric(difftime(Cytology_Weekday$signed_out_date, Cytology_Weekday$Collection_Date, units = "days"))

if (is.null(PP_Not_Weekday)){
  Cytology_Not_Weekday <- Cytology_Not_Weekday
} else {
  Cytology_Not_Weekday$Collection_to_signed_out <- as.numeric(difftime(Cytology_Not_Weekday$signed_out_date, Cytology_Not_Weekday$Collection_Date, units = "days"))
}

#recieve to signed out
#without weekends and holidays
Cytology_Weekday$Received_to_signed_out <- bizdays(Cytology_Weekday$Received_Date, Cytology_Weekday$signed_out_date)

if (is.null(PP_Not_Weekday)){
  Cytology_Not_Weekday <- Cytology_Not_Weekday
} else {
  Cytology_Not_Weekday$Received_to_signed_out <- bizdays(Cytology_Not_Weekday$Received_Date, Cytology_Not_Weekday$signed_out_date)
}

#find any negative calculation and remove them
Cytology_Weekday <- Cytology_Weekday[!(Cytology_Weekday$Collection_to_signed_out<=0 | Cytology_Weekday$Received_to_signed_out<=0),]

if (is.null(PP_Not_Weekday)){
  Cytology_Not_Weekday <- Cytology_Not_Weekday
} else {
  Cytology_Not_Weekday <- Cytology_Not_Weekday[!(Cytology_Not_Weekday$Collection_to_signed_out<=0 | Cytology_Not_Weekday$Received_to_signed_out<=0),]
}

#Calculate Average for Collection to signed out and number of cases signed

Cytology_Cases_Signed <- summarise(group_by(Cytology_Weekday,spec_group, Patient.Setting), No_Cases_Signed = n())

Cytology_Patient_Metric <- summarise(group_by(Cytology_Weekday,spec_group, Facility,Patient.Setting), Avg_Collection_to_Signed_out=format(round(mean(Collection_to_signed_out),2)))

Cytology_Patient_Metric <- dcast(Cytology_Patient_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Avg_Collection_to_Signed_out" )


if (is.null(PP_Not_Weekday)){
  Cytology_Patient_Metric_Not_Weekday <- NULL
  Cytology_Cases_Signed_Not_Weekday <- NULL
} else {
  Cytology_Cases_Signed_Not_Weekday <- summarise(group_by(Cytology_Not_Weekday,spec_group, Patient.Setting), No_Cases_Signed = n())
  
  Cytology_Patient_Metric_Not_Weekday <- summarise(group_by(Cytology_Not_Weekday,spec_group, Facility,Patient.Setting), Avg_Collection_to_Signed_out=format(round(mean(Collection_to_signed_out),2)))
  
  Cytology_Patient_Metric_Not_Weekday <- dcast(Cytology_Patient_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Avg_Collection_to_Signed_out" )
}

#Calculate % Receive to result TAT within target
GYN_Lab_Metric <- summarise(group_by(Cytology_Weekday[Cytology_Weekday$spec_group=="CYTO GYN",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 10)/sum(Received_to_signed_out >= 0),2)))
GYN_Lab_Metric <- dcast(GYN_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )

NONGYN_Lab_Metric <- summarise(group_by(Cytology_Weekday[Cytology_Weekday$spec_group=="CYTO NONGYN",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 2)/sum(Received_to_signed_out >= 0),2)))
NONGYN_Lab_Metric <- dcast(NONGYN_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )

if (is.null(PP_Not_Weekday)){
  GYN_Lab_Metric_Not_Weekday <- NULL
  NONGYN_Lab_Metric_Not_Weekday <-NULL
} else {
  GYN_Lab_Metric_Not_Weekday <- summarise(group_by(Cytology_Not_Weekday[Cytology_Not_Weekday$spec_group=="CYTO GYN",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 10)/sum(Received_to_signed_out >= 0),2)))
  GYN_Lab_Metric_Not_Weekday <- dcast(GYN_Lab_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
  
  NONGYN_Lab_Metric_Not_Weekday <- summarise(group_by(Cytology_Not_Weekday[Cytology_Not_Weekday$spec_group=="CYTO NONGYN",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 2)/sum(Received_to_signed_out >= 0),2)))
  NONGYN_Lab_Metric_Not_Weekday <- dcast(NONGYN_Lab_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
}


#here I will merge number of cases signed, received to result TAT, and acollect to result TAT calcs into one table

#Cytology Weekday table
Cytology_Table_Weekday <- left_join(full_join(Cytology_Cases_Signed, Cytology_Patient_Metric), bind_rows(GYN_Lab_Metric,NONGYN_Lab_Metric), by = c("spec_group", "Patient.Setting"))

#Cytology Not-Weekday table
if (is.null(PP_Not_Weekday)){
  Cytology_Table_Not_Weekday <- NULL
} else{
  Cytology_Table_Not_Weekday <- left_join(full_join(Cytology_Cases_Signed_Not_Weekday, Cytology_Patient_Metric_Not_Weekday), bind_rows(GYN_Lab_Metric_Not_Weekday,NONGYN_Lab_Metric_Not_Weekday), by = c("spec_group", "Patient.Setting"))
}


#Surgical Pathology

#Upload the exclusion vs inclusion criteria associated with the GI codes
GI_Codes <- data.frame(read_excel(choose.files(caption = "Select GI Codes dataframe"), 1), stringsAsFactors = FALSE)

#Merge the exclusion/inclusion cloumn into the modified powerpath Dataset for weekdays and not weekdays

PP_Weekday_Excl <- merge(x = PP_Weekday_PS, y= GI_Codes, all.x = TRUE)
Surgical_Pathology_Weekday <- PP_Weekday_Excl[which(((PP_Weekday_Excl$spec_group=="GI") &(PP_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Weekday_Excl$spec_group=="Breast"),]

PP_Not_Weekday_Excl <- merge(x = PP_Not_Weekday_PS, y= GI_Codes, all.x = TRUE)
Surgical_Pathology_Not_Weekday <- PP_Not_Weekday_Excl[which(((PP_Not_Weekday_Excl$spec_group=="GI") &(PP_Not_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Not_Weekday_Excl$spec_group=="Breast"),]

#If there are any N/A in the dates, remove that sample
Surgical_Pathology_Weekday <- Surgical_Pathology_Weekday[!(is.na(Surgical_Pathology_Weekday$Case_created_date)|is.na(Surgical_Pathology_Weekday$Collection_Date) | is.na(Surgical_Pathology_Weekday$Received_Date) | is.na(Surgical_Pathology_Weekday$signed_out_date)),]
Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday[!(is.na(Surgical_Pathology_Not_Weekday$Case_created_date)|is.na(Surgical_Pathology_Not_Weekday$Collection_Date) | is.na(Surgical_Pathology_Not_Weekday$Received_Date) | is.na(Surgical_Pathology_Not_Weekday$signed_out_date)),]

#Change all Dates into POSIXct format to start the calculations

Surgical_Pathology_Weekday$Case_created_date <- as.POSIXct(Surgical_Pathology_Weekday$Case_created_date,format='%m/%d/%y %I:%M %p')
Surgical_Pathology_Weekday$Collection_Date <- as.POSIXct(Surgical_Pathology_Weekday$Collection_Date,format='%m/%d/%y %I:%M %p')
Surgical_Pathology_Weekday$Received_Date <- as.POSIXct(Surgical_Pathology_Weekday$Received_Date,format='%m/%d/%y %I:%M %p')
Surgical_Pathology_Weekday$signed_out_date <- as.POSIXct(Surgical_Pathology_Weekday$signed_out_date,format='%m/%d/%y %I:%M %p')

if (is.null(PP_Not_Weekday)){
  Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday
} else {
  Surgical_Pathology_Not_Weekday$Case_created_date <- as.POSIXct(Surgical_Pathology_Not_Weekday$Case_created_date,format='%m/%d/%y %I:%M %p')
  Surgical_Pathology_Not_Weekday$Collection_Date <- as.POSIXct(Surgical_Pathology_Not_Weekday$Collection_Date,format='%m/%d/%y %I:%M %p')
  Surgical_Pathology_Not_Weekday$Received_Date <- as.POSIXct(Surgical_Pathology_Not_Weekday$Received_Date,format='%m/%d/%y %I:%M %p')
  Surgical_Pathology_Not_Weekday$signed_out_date <- as.POSIXct(Surgical_Pathology_Not_Weekday$signed_out_date,format='%m/%d/%y %I:%M %p')
}

#add columns for calculations: collection to signed out and received to signed out
#Collection to signed out
#All days in the calendar
Surgical_Pathology_Weekday$Collection_to_signed_out <- as.numeric(difftime(Surgical_Pathology_Weekday$signed_out_date, Surgical_Pathology_Weekday$Collection_Date, units = "days"))

if (is.null(PP_Not_Weekday)){
  Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday
} else {
  Surgical_Pathology_Not_Weekday$Collection_to_signed_out <- as.numeric(difftime(Surgical_Pathology_Not_Weekday$signed_out_date, Surgical_Pathology_Not_Weekday$Collection_Date, units = "days"))
}


#recieve to signed out
#without weekends and holidays
Surgical_Pathology_Weekday$Received_to_signed_out <- bizdays(Surgical_Pathology_Weekday$Received_Date, Surgical_Pathology_Weekday$signed_out_date)

if (is.null(PP_Not_Weekday)){
  Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday
} else {
  Surgical_Pathology_Not_Weekday$Received_to_signed_out <- bizdays(Surgical_Pathology_Not_Weekday$Received_Date, Surgical_Pathology_Not_Weekday$signed_out_date)
}

#find any negative calculation and remove them
Surgical_Pathology_Weekday <- Surgical_Pathology_Weekday[!(Surgical_Pathology_Weekday$Collection_to_signed_out<=0 | Surgical_Pathology_Weekday$Received_to_signed_out<=0),]
if (is.null(PP_Not_Weekday)){
  Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday
} else {
  Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday[!(Surgical_Pathology_Not_Weekday$Collection_to_signed_out<=0 | Surgical_Pathology_Not_Weekday$Received_to_signed_out<=0),]
}

#Calculate Average for Collection to signed out

Surgical_Pathology_Cases_Signed <- summarise(group_by(Surgical_Pathology_Weekday,spec_group, Patient.Setting), No_Cases_Signed = n())

Surgical_Pathology_Patient_Metric <- summarise(group_by(Surgical_Pathology_Weekday,spec_group, Facility,Patient.Setting), Avg_Collection_to_Signed_out=format(round(mean(Collection_to_signed_out),2)))

Surgical_Pathology_Patient_Metric <- dcast(Surgical_Pathology_Patient_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Avg_Collection_to_Signed_out" )

if (is.null(PP_Not_Weekday)){
  Surgical_Pathology_Patient_Metric_Not_Weekday <- NULL
  Surgical_Pathology_Cases_Signed_Not_Weekday <- NULL
} else {
  Surgical_Pathology_Cases_Signed_Not_Weekday <- summarise(group_by(Surgical_Pathology_Not_Weekday,spec_group, Patient.Setting), No_Cases_Signed = n())
  
  Surgical_Pathology_Patient_Metric_Not_Weekday <- summarise(group_by(Surgical_Pathology_Not_Weekday,spec_group, Facility,Patient.Setting), Avg_Collection_to_Signed_out=format(round(mean(Collection_to_signed_out),2)))
  
  Surgical_Pathology_Patient_Metric_Not_Weekday <- dcast(Surgical_Pathology_Patient_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Avg_Collection_to_Signed_out" )
}

#Calculate % Receive to result TAT within target
Breast_Lab_Metric <- summarise(group_by(Surgical_Pathology_Weekday[Surgical_Pathology_Weekday$spec_group=="Breast",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 5)/sum(Received_to_signed_out >= 0),2)))
Breast_Lab_Metric <- dcast(Breast_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )

GI_Lab_Metric <- summarise(group_by(Surgical_Pathology_Weekday[Surgical_Pathology_Weekday$spec_group=="GI",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 3)/sum(Received_to_signed_out >= 0),2)))
GI_Lab_Metric <- dcast(GI_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )

if (is.null(PP_Not_Weekday)){
  Breast_Lab_Metric_Not_Weekday <- NULL
  GI_Lab_Metric_Not_Weekday <-NULL
} else {
  Breast_Lab_Metric_Not_Weekday <- summarise(group_by(Surgical_Pathology_Not_Weekday[Surgical_Pathology_Not_Weekday$spec_group=="Breast",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 5)/sum(Received_to_signed_out >= 0),2)))
  Breast_Lab_Metric_Not_Weekday <- dcast(Breast_Lab_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
  
  GI_Lab_Metric_Not_Weekday <- summarise(group_by(Surgical_Pathology_Not_Weekday[Surgical_Pathology_Not_Weekday$spec_group=="GI",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 3)/sum(Received_to_signed_out >= 0),2)))
  GI_Lab_Metric_Not_Weekday <- dcast(GI_Lab_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
}

#here I will merge number of cases signed, received to result TAT, and acollect to result TAT calcs into one table

#Pathology Weekday table
Pathology_Table_Weekday <- left_join(full_join(Surgical_Pathology_Cases_Signed, Surgical_Pathology_Patient_Metric), bind_rows(Breast_Lab_Metric,GI_Lab_Metric), by = c("spec_group", "Patient.Setting"))

#Pathology Not-Weekday table
if (is.null(PP_Not_Weekday)){
  Pathology_Table_Not_Weekday <- NULL
} else{
  Pathology_Table_Not_Weekday <- left_join(full_join(Surgical_Pathology_Cases_Signed_Not_Weekday, Surgical_Pathology_Patient_Metric_Not_Weekday), bind_rows(Breast_Lab_Metric_Not_Weekday,GI_Lab_Metric_Not_Weekday), by = c("spec_group", "Patient.Setting"))
}

# Manipulate and reshape SCC and Sunquest data -----------------------------------------------
# Start summarizing data
chemistry <- scc_sun_wday_master[scc_sun_wday_master$Division == "Chemistry" & scc_sun_wday_master$TATInclude == TRUE, ] %>%
  group_by(Test, Site, DashboardPriority, DashboardSetting) %>%
  summarize(ResultedVolume = n(), ReceiveResultInTarget = sum(ReceiveResultInTarget), CollectResultInTarget = sum(CollectResultInTarget),
            ReceiveResultPercent = round(ReceiveResultInTarget/ResultedVolume, digits = 3)*100, CollectResultPercent = round(CollectResultInTarget/ResultedVolume, digits = 3)*100)

chemistry_dashboard <- melt(chemistry, id.var = c("Test", "Site", "DashboardPriority", "DashboardSetting"), measure.vars = c("ReceiveResultPercent", "CollectResultPercent"))
chemistry_summary <- dcast(chemistry_dashboard, Test + DashboardPriority + DashboardSetting ~ variable + Site, value.var = "value")                        
                    
hematology <- scc_sun_wday_master[scc_sun_wday_master$Division == "Hematology" & scc_sun_wday_master$TATInclude == TRUE, ] %>%
  group_by(Test, Site, DashboardPriority, DashboardSetting) %>%
  summarize(ResultedVolume = n(), ReceiveResultInTarget = sum(ReceiveResultInTarget), CollectResultInTarget = sum(CollectResultInTarget),
            ReceiveResultPercent = round(ReceiveResultInTarget/ResultedVolume, digits = 3)*100, CollectResultPercent = round(CollectResultInTarget/ResultedVolume, digits = 3)*100)

hematology_table <- melt(hematology, id.var = c("Test", "Site", "DashboardPriority", "DashboardSetting"), measure.vars = c("ResultedVolume", "ReceiveResultInTarget", "CollectResultInTarget", "ReceiveResultPercent", "CollectResultPercent"))

hematology_table2 <- dcast(hematology_table, Test + DashboardPriority + DashboardSetting ~ variable + Site, value.var = "value")

hematology_dashboard <- melt(hematology, id.var = c("Test", "Site", "DashboardPriority", "DashboardSetting"), measure.vars = c("ReceiveResultPercent", "CollectResultPercent"))
hematology_summary <- dcast(hematology_dashboard, Test + DashboardPriority + DashboardSetting ~ variable + Site, value.var = "value")

micro <- scc_sun_wday_master[scc_sun_wday_master$Division == "Microbiology RRL" & scc_sun_wday_master$TATInclude == TRUE, ] %>%
  group_by(Test, Site, DashboardPriority, DashboardSetting) %>%
  summarize(ResultedVolume = n(), ReceiveResultInTarget = sum(ReceiveResultInTarget), CollectResultInTarget = sum(CollectResultInTarget),
            ReceiveResultPercent = round(ReceiveResultInTarget/ResultedVolume, digits = 3)*100, CollectResultPercent = round(CollectResultInTarget/ResultedVolume, digits = 3)*100)

micro_table <- melt(micro, id.var = c("Test", "Site", "DashboardPriority", "DashboardSetting"), value.var = c("ResultedVolume", "ReceiveResultInTarget", "CollectResultInTarget", "ReceiveResultPercent", "CollectResultPercent"))

micro_table2 <- dcast(micro_table, Test + DashboardPriority + DashboardSetting ~ variable + Site, value.var = "value")

micro_dashboard <- melt(micro, id.var = c("Test", "Site", "DashboardPriority", "DashboardSetting"), measure.vars = c("ReceiveResultPercent", "CollectResultPercent"))
micro_dashboard <- micro_dashboard[!(micro_dashboard$Test == "C. diff" & micro_dashboard$DashboardSetting == "Amb"), ]

micro_summary <- dcast(micro_dashboard, Test + DashboardPriority + DashboardSetting ~ variable + Site, value.var = "value")
micro_summary[micro_summary$Test == "Rapid Flu" & micro_summary$DashboardSetting == "Amb", ]

micro_volume <- melt(micro, id.var = c("Test", "Site", "DashboardPriority", "DashboardSetting"), measure.vars = c("ResultedVolume"))
micro_volume <- dcast(micro_volume, Test + DashboardPriority + DashboardSetting ~ variable + Site)

# # Determine percentage of labs with missing collection times -------------------------------
missing_collect <- scc_sun_wday_master %>%
  group_by(Site) %>%
  summarize(ResultedVolume = n(), MissingCollection = sum(MissingCollect, na.rm = TRUE), Percent = round(MissingCollection/ResultedVolume*100, digits = 1))

missing_collect_table <- dcast(missing_collect, "Percent Specimens Missing Collection" ~ Site, value.var = "Percent")

# Determine number of add-on orders stratified by test and site ------------------------------
add_on_volume <- scc_sun_wday_master %>%
  group_by(Test, Site) %>%
  summarize(AddOnVolume = sum(AddOnMaster == "AddOn", na.rm = TRUE))

add_on_table <- dcast(add_on_volume, Test ~ Site, value.var = "AddOnVolume")
