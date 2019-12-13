
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("xlsx")
#install.packages("readxl")
#install.packages("dplyr")
#install.packages("reshape2")

#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(dplyr)
library(lubridate)
library(reshape2)

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
}

# #-----------PowerPath Excel Files-----------#
# #For the powerpath files the read excel is starting from the second row
# #Also I made sure to remove the last line
# if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){
#   PP_Holiday_Monday_or_Friday <- read_excel(choose.files(caption = "Select PowerPath Holiday Report"), skip = 1, 1)
#   PP_Holiday_Monday_or_Friday  <- PP_Holiday_Monday[-nrow(PP_Holiday_Monday_or_Friday),]
#   PP_Sunday <- read_excel(choose.files(caption = "Select PowerPath Sunday Report"), skip = 1, 1)
#   PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
#   PP_Saturday <- read_excel(choose.files(caption = "Select PowerPath Saturday Report"), skip = 1, 1)
#   PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
#   #Merge the weekend data with the holiday data in one data frame
#   PP_Not_Weekday <- data.frame(rbind(rbind(PP_Holiday_Monday_or_Friday ,PP_Sunday),PP_Saturday),stringsAsFactors = FALSE)
#   PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
#   PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),],stringsAsFactors = FALSE)
# } else if ((Holiday_Det) & (Yesterday_Day =="Sun")){
#   PP_Sunday <- read_excel(choose.files(caption = "Select PowerPath Sunday Report"), skip = 1, 1)
#   PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
#   PP_Saturday <- read_excel(choose.files(caption = "Select PowerPath Saturday Report"), skip = 1, 1)
#   PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
#   #Merge the weekend data with the holiday data in one data frame
#   PP_Not_Weekday <- data.frame(rbind(PP_Sunday,PP_Saturday),stringsAsFactors = FALSE)
#   PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
#   PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
# } else if ((Holiday_Det) & ((Yesterday_Day !="Mon")|(Yesterday_Day !="Sun"))){
#   PP_Holiday_Weekday <- read_excel(choose.files(caption = "Select PowerPath Holiday Report"), skip = 1, 1)
#   PP_Holiday_Weekday <- PP_Holiday_Weekday[-nrow(PP_Holiday_Weekday),]
#   PP_Not_Weekday <- data.frame(PP_Holiday_Weekday, stringsAsFactors = FALSE)
#   PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
#   PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
# } else {
#   PP_Weekday <- read_excel(choose.files(caption = "Select PowerPath Weekday Report"), skip = 1, 1)
#   PP_Weekday <- data.farme(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
# }
# 
# #-----------Cytology Backlog Excel Files-----------#
# #For the backlog files the read excel is starting from the second row
# #Also I made sure to remove the last line
# 
# Cytology_Backlog <- read_excel(choose.files(caption = "Select Cytology Backlog Report"), skip = 1, 1)
# Cytology_Backlog <- data.frame(Cytology_Backlog[-nrow(Cytology_Backlog),], stringsAsFactors = FALSE)
# 
# 
# #------------------------------Data Pre-Processing------------------------------#
# 
# 
# #Using Rev Center to determine patient setting
# Patient_Setting <- data.frame(read_excel(choose.files(caption = "Select Cytology Backlog Report"), sheet = "final"), stringsAsFactors = FALSE)
# 
# #vlookup the Rev_Center and its corresponding patient setting for the PowerPath Data
# PP_Weekday_PS <- merge(x=PP_Weekday, y=Patient_Setting, all.x = TRUE ) 
# PP_Not_Weekday_PS <- merge(x=PP_Not_Weekday, y=Patient_Setting, all.x = TRUE )
# 
# #Cytology
# #Keep the cyto gyn and cyto non-gyn
# 
# Cytology_Weekday <- PP_Weekday_PS[which(PP_Weekday_PS$spec_group=="CYTO NONGYN" | PP_Weekday_PS$spec_group=="CYTO GYN"),]
# Cytology_NoT_Weekday <- PP_Not_Weekday_PS[which(PP_Not_Weekday_PS$spec_group=="CYTO NONGYN" | PP_Not_Weekday_PS$spec_group=="CYTO GYN"),]
# 
# #Surgical Pathology
# 
# #Upload the exclusion vs inclusion criteria associated with the GI codes
# GI_Codes <- data.frame(read_excel(choose.files(caption = "Select Cytology Backlog Report"), 1), stringsAsFactors = FALSE)
# 
# #Merge the exclusion/inclusion cloumn into the modified powerpath Dataset for weekdays and not weekdays
# 
# PP_Weekday_Excl <- merge(x = PP_Weekday_PS, y= GI_Codes, all.x = TRUE)
# Surgical_Pathology_Weekday <- PP_Weekday_Excl[which(((PP_Weekday_Excl$spec_group=="GI") &(PP_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Weekday_Excl$spec_group=="Breast"),]
# 
# PP_Not_Weekday_Excl <- merge(x = PP_Not_Weekday_PS, y= GI_Codes, all.x = TRUE)
# Surgical_Pathology_Not_Weekday <- PP_Not_Weekday_Excl[which(((PP_Not_Weekday_Excl$spec_group=="GI") &(PP_Not_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Not_Weekday_Excl$spec_group=="Breast"),]

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

cp_micro_lab_order <- c("Troponin", "Lactate WB", "BUN", "HGB", "PT")
site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")
pt_setting_order <- c("ED", "ICU", "IP Non-ICU", "Amb", "Other")
pt_setting_order2 <- c("ED & ICU", "IP Non-ICU", "Amb", "Other")

# SCC lookup references ----------------------------------------------
# Crosswalk labs included and remove out of scope labs
scc_wday <- left_join(scc_wday, test_code[ , c("Test", "SCC_TestID")], by = c("TEST_ID" = "SCC_TestID"))
scc_wday$TEST_ID <- as.factor(scc_wday$TEST_ID)
scc_wday$TestIncl <- ifelse(is.na(scc_wday$Test), FALSE, TRUE)
scc_wday <- scc_wday[scc_wday$TestIncl == TRUE, ]
# Crosswalk units and identify ICUs
scc_wday$WardandName <- paste(scc_wday$Ward, scc_wday$WARD_NAME)
scc_wday <- left_join(scc_wday, scc_icu[ , c("Concatenate", "ICU")], by = c("WardandName" = "Concatenate"))
scc_wday[is.na(scc_wday$ICU), "ICU"] <- FALSE
# Crosswalk unit type
scc_wday <- left_join(scc_wday, scc_setting, by = c("CLINIC_TYPE" = "Clinic_Type"))
# Crosswalk site name
scc_wday <- left_join(scc_wday, mshs_site, by = c("SITE" = "DataSite"))

# SCC data formatting ----------------------------------------------
scc_wday[c("Ward", "WARD_NAME", 
           "REQUESTING_DOC", 
           "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", "Test",
           "COLLECT_CENTER_ID", "SITE", "Site",
           "CLINIC_TYPE", "Setting", "SettingRollUp")] <- lapply(scc_wday[c("Ward", "WARD_NAME", 
                                                                             "REQUESTING_DOC", 
                                                                             "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", "Test",
                                                                             "COLLECT_CENTER_ID", "SITE", "Site",
                                                                            "CLINIC_TYPE", "Setting", "SettingRollUp")], as.factor)

scc_wday[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(scc_wday[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")],
                                                                                           as.POSIXlt, tz = "", format = "%Y-%m-%d %H:%M:%OS", options(digits.sec = 1))

# SCC data preprocessing -------------------------------------------
# Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
scc_wday$MasterSetting <- ifelse(scc_wday$CLINIC_TYPE == "E", "ED", 
                                 ifelse(scc_wday$CLINIC_TYPE == "O", "Amb", 
                                         ifelse(scc_wday$CLINIC_TYPE == "I" & scc_wday$ICU == TRUE, "ICU", 
                                                ifelse(scc_wday$CLINIC_TYPE == "I" & scc_wday$ICU != TRUE, "IP Non-ICU", "Other"))))

# Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
scc_wday$AdjPriority <- ifelse(scc_wday$MasterSetting == "ED" | scc_wday$MasterSetting == "ICU" | scc_wday$PRIORITY == "S", "Stat", "Routine")
scc_wday$DashboardPriority <- ifelse(tat_targets$Priority[match(scc_wday$Test, tat_targets$Test)] == "All", "AllSpecimen", scc_wday$AdjPriority)

# Calculate turnaround times
scc_wday$CollectToReceive <- scc_wday$RECEIVE_DATE - scc_wday$COLLECTION_DATE
scc_wday$ReceiveToResult <- scc_wday$VERIFIED_DATE - scc_wday$RECEIVE_DATE
scc_wday$CollectToResult <- scc_wday$VERIFIED_DATE - scc_wday$COLLECTION_DATE
scc_wday[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(scc_wday[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")

# Identify add on orders as orders placed more than 5 min after specimen received
scc_wday$AddOnMaster <- ifelse(difftime(scc_wday$ORDERING_DATE, scc_wday$RECEIVE_DATE, units = "mins") > 5, "AddOn", "Original")

# Identify specimens with missing collections times as those with collection time defaulted to receive time
scc_wday$MissingCollect <- ifelse(scc_wday$CollectToReceive == 0, TRUE, FALSE)

# Determine target TAT based on test, priority, and patient setting
scc_wday$Concate1 <- paste(scc_wday$Test, scc_wday$DashboardPriority)
scc_wday$Concate2 <- paste(scc_wday$Test, scc_wday$DashboardPriority, scc_wday$MasterSetting)

scc_wday$ReceiveResultTarget <- ifelse(!is.na(match(scc_wday$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(scc_wday$Concate2, tat_targets$Concate)], 
                                ifelse(!is.na(match(scc_wday$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(scc_wday$Concate1, tat_targets$Concate)],
                                       tat_targets$ReceiveToResultTarget[match(scc_wday$Test, tat_targets$Concate)]))

scc_wday$CollectResultTarget <- ifelse(!is.na(match(scc_wday$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(scc_wday$Concate2, tat_targets$Concate)], 
                                ifelse(!is.na(match(scc_wday$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(scc_wday$Concate1, tat_targets$Concate)],
                                       tat_targets$CollectToResultTarget[match(scc_wday$Test, tat_targets$Concate)]))

scc_wday$ReceiveResultInTarget <- ifelse(scc_wday$ReceiveToResult <= scc_wday$ReceiveResultTarget, TRUE, FALSE)
scc_wday$CollectResultInTarget <- ifelse(scc_wday$CollectToResult <= scc_wday$CollectResultTarget, TRUE, FALSE)

# Identify and remove duplicate tests
scc_wday$Concate3 <- paste(scc_wday$LAST_NAME, scc_wday$FIRST_NAME, 
                           scc_wday$ORDER_ID, scc_wday$TEST_NAME,
                            scc_wday$COLLECTION_DATE, scc_wday$RECEIVE_DATE, scc_wday$VERIFIED_DATE)

scc_wday <- scc_wday[!duplicated(scc_wday$Concate3), ]

# Identify which labs to include in TAT analysis
# Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
scc_wday$TATInclude <- ifelse(scc_wday$AddOnMaster == "AddOn" | scc_wday$MasterSetting == "Other" | scc_wday$CollectToResult < 0 | scc_wday$ReceiveToResult < 0 | is.na(scc_wday$CollectToResult) | is.na(scc_wday$ReceiveToResult), FALSE, TRUE)


# Sunquest lookup references -------------------------------------------
# Crosswalk labs included and remove out of scope labs
sun_wday <- left_join(sun_wday, test_code[ , c("Test", "SUN_TestCode")], by = c("TestCode" = "SUN_TestCode"))
sun_wday$TestCode <- as.factor(sun_wday$TestCode)
sun_wday$TestIncl <- ifelse(is.na(sun_wday$Test), FALSE, TRUE)
sun_wday <- sun_wday[sun_wday$TestIncl == TRUE, ]
# Crosswalk units and identify ICUs
sun_wday$LocandName <- paste(sun_wday$LocCode, sun_wday$LocName)
sun_wday <- left_join(sun_wday, sun_icu[ , c("Concatenate", "ICU")], by = c("LocandName" = "Concatenate"))
sun_wday[is.na(sun_wday$ICU), "ICU"] <- FALSE
# Crosswalk unit type
sun_wday <- left_join(sun_wday, sun_setting, by = c("LocType" = "LocType"))
# Crosswalk site name
sun_wday <- left_join(sun_wday, mshs_site, by = c("HospCode" = "DataSite"))


# Sunquest data formatting --------------------------------------------
sun_wday[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
           "LocType", "LocCode", "LocName", 
           "SpecimenPriority", "PhysName", "SHIFT",
           "ReceiveTech", "ResultTech", "PerformingLabCode",
           "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- lapply(sun_wday[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
                                                                                           "LocType", "LocCode", "LocName", 
                                                                                           "SpecimenPriority", "PhysName", "SHIFT",
                                                                                           "ReceiveTech", "ResultTech", "PerformingLabCode",
                                                                                           "Test", "LocandName", "Setting", "SettingRollUp", "Site")], as.factor)

sun_wday[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")] <- lapply(sun_wday[c("OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime")], 
                                                                                               as.POSIXlt, tz = "", format = "%m/%d/%Y %H:%M:%S")




# Sunquest data preprocessing --------------------------------------------------
# Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
sun_wday$MasterSetting <- ifelse(sun_wday$SettingRollUp == "ED", "ED", 
                                 ifelse(sun_wday$SettingRollUp == "Amb", "Amb", 
                                        ifelse(sun_wday$SettingRollUp == "IP" & sun_wday$ICU == TRUE, "ICU",
                                               ifelse(sun_wday$SettingRollUp == "IP" & sun_wday$ICU == FALSE, "IP Non-ICU", "Other"))))


# Update priority to reflect ED/ICU as stat and create Master Priority for labs where all specimens are treated as stat
sun_wday$AdjPriority <- ifelse(sun_wday$MasterSetting != "ED" & sun_wday$MasterSetting != "ICU" & is.na(sun_wday$SpecimenPriority), "Routine",
                               ifelse(sun_wday$MasterSetting == "ED" | sun_wday$MasterSetting == "ICU" | sun_wday$SpecimenPriority == "S", "Stat", "Routine"))
sun_wday$DashboardPriority <- ifelse(tat_targets$Priority[match(sun_wday$Test, tat_targets$Test)] == "All", "AllSpecimen", sun_wday$AdjPriority)

# Calculate turnaround times
sun_wday$CollectToReceive <- sun_wday$ReceiveDateTime - sun_wday$CollectDateTime
sun_wday$ReceiveToResult <- sun_wday$ResultDateTime - sun_wday$ReceiveDateTime
sun_wday$CollectToResult <- sun_wday$ResultDateTime - sun_wday$CollectDateTime
sun_wday[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- lapply(sun_wday[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], as.numeric, units = "mins")

# Identify add on orders as orders placed more than 5 min after specimen received
sun_wday$AddOnMaster <- ifelse(difftime(sun_wday$OrderDateTime, sun_wday$ReceiveDateTime, units = "mins") > 5, "AddOn", "Original")

# Identify specimens with missing collections times as those with collection time defaulted to order time
sun_wday$MissingCollect <- ifelse(sun_wday$CollectDateTime == sun_wday$OrderDateTime, TRUE, FALSE)

# Determine target TAT based on test, priority, and patient setting
sun_wday$Concate1 <- paste(sun_wday$Test, sun_wday$DashboardPriority)
sun_wday$Concate2 <- paste(sun_wday$Test, sun_wday$DashboardPriority, sun_wday$MasterSetting)

sun_wday$ReceiveResultTarget <- ifelse(!is.na(match(sun_wday$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(sun_wday$Concate2, tat_targets$Concate)], 
                                       ifelse(!is.na(match(sun_wday$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(sun_wday$Concate1, tat_targets$Concate)],
                                              tat_targets$ReceiveToResultTarget[match(sun_wday$Test, tat_targets$Concate)]))

sun_wday$CollectResultTarget <- ifelse(!is.na(match(sun_wday$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(sun_wday$Concate2, tat_targets$Concate)], 
                                       ifelse(!is.na(match(sun_wday$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(sun_wday$Concate1, tat_targets$Concate)],
                                              tat_targets$CollectToResultTarget[match(sun_wday$Test, tat_targets$Concate)]))

sun_wday$ReceiveResultInTarget <- ifelse(sun_wday$ReceiveToResult <= sun_wday$ReceiveResultTarget, TRUE, FALSE)
sun_wday$CollectResultInTarget <- ifelse(sun_wday$CollectToResult <= sun_wday$CollectResultTarget, TRUE, FALSE)

# Identify and remove duplicate tests
sun_wday$Concate3 <- paste(sun_wday$PtNumber, 
                           sun_wday$HISOrderNumber, sun_wday$TSTName,
                           sun_wday$CollectDateTime, sun_wday$ReceiveDateTime, sun_wday$ResultDateTime)

sun_wday <- sun_wday[!duplicated(sun_wday$Concate3), ]



# Identify which labs to include in TAT analysis
# Exclude add on orders, orders from "other" settings, orders with collect or receive times after result, or orders with missing collect, receive, or result timestamps
sun_wday$TATInclude <- ifelse(sun_wday$AddOnMaster == "AddOn" | sun_wday$MasterSetting == "Other" | sun_wday$CollectToResult < 0 | sun_wday$ReceiveToResult < 0 | is.na(sun_wday$CollectToResult) | is.na(sun_wday$ReceiveToResult), FALSE, TRUE)


# Standardize and combine SCC and Sunquest data ----------------------------------------------
# Create new dataframe with columns of interest and rename with common names across SCC and Sunquest to bind later
scc_wday_master <- scc_wday[ , c("Ward", "WARD_NAME", "WardandName", 
                                 "ORDER_ID", "REQUESTING_DOC NAME", "MPI", "WORK SHIFT", 
                                 "TEST_NAME", "Test", "PRIORITY", 
                                 "Site", "ICU", "CLINIC_TYPE", 
                                 "Setting", "SettingRollUp", "MasterSetting", 
                                 "AdjPriority", "DashboardPriority", 
                                 "ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE", 
                                 "CollectToReceive", "ReceiveToResult", "CollectToResult", 
                                 "AddOnMaster", "MissingCollect", 
                                 "ReceiveResultTarget", "CollectResultTarget", 
                                 "ReceiveResultInTarget", "CollectResultInTarget", 
                                 "TATInclude")]

sun_wday_master <- sun_wday[ ,c("LocCode", "LocName", "LocandName", 
                           "HISOrderNumber", "PhysName", "PtNumber", "SHIFT",            
                           "TSTName", "Test", "SpecimenPriority", 
                           "Site", "ICU", "LocType",               
                           "Setting", "SettingRollUp", "MasterSetting", 
                           "AdjPriority", "DashboardPriority", 
                           "OrderDateTime", "CollectDateTime", "ReceiveDateTime", "ResultDateTime", 
                           "CollecttoReceive", "ReceivetoResult", "CollecttoResult", 
                           "AddOnMaster", "MissingCollect", 
                           "ReceiveResultTarget", "CollectResultTarget", 
                           "ReceiveResultInTarget", "CollectResultInTarget", 
                           "TATInclude")]

colnames(scc_wday_master) <- c("LocCode", "LocName", "LocConcat", 
                               "OrderID", "RequestMD", "MSMRN", "WorkShift", 
                               "TestName", "Test", "OrderPriority", 
                               "Site", "ICU", "LocType", 
                               "Setting", "SettingRollUp", "MasterSetting", 
                               "AdjPriority", "DashboardPriority", 
                               "OrderTime", "CollectTime", "ReceiveTime", "ResultTime", 
                               "CollectToReceiveTAT", "ReceiveToResultTAT", "CollectToResultTAT", 
                               "AddOnMaster", "MissingCollect", 
                               "ReceiveResultTarget", "CollectResultTarget", 
                               "ReceiveResultInTarget", "CollectResultInTarget", 
                               "TATInclude")

colnames(sun_wday_master) <- c("LocCode", "LocName", "LocConcat", 
                               "OrderID", "RequestMD", "MSMRN", "WorkShift", 
                               "TestName", "Test", "OrderPriority", 
                               "Site", "ICU", "LocType", 
                               "Setting", "SettingRollUp", "MasterSetting", 
                               "AdjPriority", "DashboardPriority", 
                               "OrderTime", "CollectTime", "ReceiveTime", "ResultTime", 
                               "CollectToReceiveTAT", "ReceiveToResultTAT", "CollectToResultTAT", 
                               "AddOnMaster", "MissingCollect", 
                               "ReceiveResultTarget", "CollectResultTarget", 
                               "ReceiveResultInTarget", "CollectResultInTarget", 
                               "TATInclude")

cp_micro_wday_master <- rbind(scc_wday_master, sun_wday_master)

cp_micro_wday_master[c("LocConcat", "RequestMD", "WorkShift", "MasterSetting", "AdjPriority", "DashboardPriority", "AddOnMaster")] <-
  lapply(cp_micro_wday_master[c("LocConcat", "RequestMD", "WorkShift", "MasterSetting", "AdjPriority", "DashboardPriority", "AddOnMaster")], as.factor)

cp_micro_wday_master$Site <- factor(cp_micro_wday_master$Site, levels = site_order)
cp_micro_wday_master$Test <- factor(cp_micro_wday_master$Test, levels = cp_micro_lab_order)
cp_micro_wday_master$MasterSetting <- factor(cp_micro_wday_master$MasterSetting, levels = pt_setting_order)


# Manipulate and reshape SCC and Sunquest data -----------------------------------------------
# Start summarizing data
chem_hem <- cp_micro_wday_master[(cp_micro_wday_master$Test == "Troponin" | cp_micro_wday_master$Test == "Lactate WB" | cp_micro_wday_master$Test == "BUN" | cp_micro_wday_master$Test == "HGB" | cp_micro_wday_master$Test == "PT") & cp_micro_wday_master$TATInclude == TRUE, ] %>%
  group_by(Test, Site, DashboardPriority, MasterSetting) %>%
  summarize(ResultedVolume = n(), CollectResultVolInTarget = sum(CollectResultInTarget), ReceiveResultInTarget = sum(ReceiveResultInTarget),
            CollectToResultPercent = round(CollectResultVolInTarget/ResultedVolume, digits = 1)*100, ReceiveResultPercent = round(ReceiveResultInTarget/ResultedVolume, digits = 1)*100)

chem_hem_2 <- melt(chem_hem, na.rm = FALSE, id.vars = c("Test", "Site", "DashboardPriority", "MasterSetting"), 
                   measure.vars = c("CollectToResultPercent", "ReceiveResultPercent"))

chem_hem_2 <- dcast(chem_hem_2, Test + DashboardPriority + MasterSetting ~ rev(variable) + Site, value.var = "value")

micro <- scc_wday_master[(scc_wday_master$Test == "Rapid Flu" | scc_wday_master$Test == "C. diff") & scc_wday_master$TATInclude == TRUE, ] %>%
  group_by(Test, Site, DashboardPriority, MasterSetting) %>%
  summarize(ResultedVolume = n(), CollectResultVolInTarget = sum(CollectResultInTarget), ReceiveResultInTarget = sum(ReceiveResultInTarget),
            CollectToResultPercent = CollectResultVolInTarget/ResultedVolume*100, ReceiveResultPercent = ReceiveResultInTarget/ResultedVolume*100)
