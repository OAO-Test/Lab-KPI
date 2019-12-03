
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("xlsx")
#install.packages("readxl")
#install.packages("dplyr")

#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(dplyr)
library(lubridate)

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
  SCC_Holiday_Monday_or_Friday <- read_excel(choose.files(caption = "Select SCC Holiday Report"), sheet = 1, col_names = TRUE)
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

# #-----------SunQuest Excel Files-----------#
# 
# if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){
#   SQ_Holiday_Monday_or_Friday <- read_excel(choose.files(caption = "Select SunQuest Holiday Report"), sheet = 1, col_names = TRUE)
#   SQ_Sunday <- read_excel(choose.files(caption = "Select SunQuest Sunday Report"), sheet = 1, col_names = TRUE)
#   SQ_Saturday <- read_excel(choose.files(caption = "Select SunQuest Saturday Report"), sheet = 1, col_names = TRUE)
#   #Merge the weekend data with the holiday data in one data frame
#   SQ_Not_Weekday <- rbind(rbind(SQ_Holiday_Monday_or_Friday ,SQ_Sunday),SQ_Saturday)
#   SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
# } else if ((Holiday_Det) & (Yesterday_Day =="Sun")){
#   SQ_Sunday <- read_excel(choose.files(caption = "Select SunQuest Sunday Report"), sheet = 1, col_names = TRUE)
#   SQ_Saturday <- read_excel(choose.files(caption = "Select SunQuest Saturday Report"), sheet = 1, col_names = TRUE)
#   #Merge the weekend data with the holiday data in one data frame
#   SQ_Not_Weekday <- rbind(SQ_Sunday,SQ_Saturday)
#   SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
# } else if ((Holiday_Det) & ((Yesterday_Day !="Mon")|(Yesterday_Day !="Sunday"))){
#   SQ_Holiday_Weekday <- read_excel(choose.files(caption = "Select SunQuest Holiday Report"), sheet = 1, col_names = TRUE)
#   SQ_Not_Weekday <- SQ_Holiday_Weekday
#   SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
# } else {
#   SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
# }

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
test_code <- read_excel("Analysis REference 2019-12-02.xlsx", sheet = "TestNames")
tat_targets <- read_excel("Analysis REference 2019-12-02.xlsx", sheet = "Turnaround Targets")
scc_icu <- read_excel("Analysis Reference 2019-12-02.xlsx", sheet = "SCC_ICU")
scc_setting <- read_excel("Analysis Reference 2019-12-02.xlsx", sheet = "SCC_ClinicType")

scc_wday <- SCC_Weekday

# Lookup references----------------------------------------------
scc_wday <- left_join(scc_wday, test_code[ , c("Test", "SCC_TestID")], by = c("TEST_ID" = "SCC_TestID"))
scc_wday$TEST_ID <- as.factor(scc_wday$TEST_ID)
scc_wday$TestIncl <- ifelse(is.na(scc_wday$Test), FALSE, TRUE)
scc_wday <- scc_wday[scc_wday$TestIncl == TRUE, ]
scc_wday$WardandName <- paste(scc_wday$Ward, scc_wday$WARD_NAME)
scc_wday <- left_join(scc_wday, scc_icu[ , c("Concatenate", "ICU")], by = c("WardandName" = "Concatenate"))
scc_wday[is.na(scc_wday$ICU), "ICU"] <- FALSE
scc_wday <- left_join(scc_wday, scc_setting, by = c("CLINIC_TYPE" = "Clinic_Type"))

# Format data fields----------------------------------------------
scc_wday[c("Ward", "WARD_NAME", 
           "REQUESTING_DOC", 
           "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", "Test",
           "COLLECT_CENTER_ID", "SITE", 
           "CLINIC_TYPE", "Setting", "Roll Up")] <- lapply(scc_wday[c("Ward", "WARD_NAME", 
                                                                             "REQUESTING_DOC", 
                                                                             "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", "Test",
                                                                             "COLLECT_CENTER_ID", "SITE", "CLINIC_TYPE", "Setting", "Roll Up")], as.factor)

scc_wday[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")] <- lapply(scc_wday[c("ORDERING_DATE", "COLLECTION_DATE", "RECEIVE_DATE", "VERIFIED_DATE")],
                                                                                           as.POSIXlt, tz = "", format = "%Y-%m-%d %H:%M:%OS", options(digits.sec = 1))

# Update patient setting to reflect ICU/Non-ICU and update priorities for ED and ICU labs
scc_wday$MasterSetting <- ifelse(scc_wday$CLINIC_TYPE == "E", "ED", 
                                 ifelse(scc_wday$CLINIC_TYPE == "O", "Amb", 
                                         ifelse(scc_wday$CLINIC_TYPE == "I" & scc_wday$ICU == TRUE, "ICU", 
                                                ifelse(scc_wday$CLINIC_TYPE == "I" & scc_wday$ICU != TRUE, "IP Non-ICU", "Other"))))
scc_wday$MasterPriority <- ifelse(scc_wday$MasterSetting == "ED" | scc_wday$MasterSetting == "ICU" | scc_wday$PRIORITY == "S", "Stat", "Routine")

# Calculate turnaround times
scc_wday$CollectToResult <- scc_wday$VERIFIED_DATE - scc_wday$COLLECTION_DATE
scc_wday$CollectToReceive <- scc_wday$RECEIVE_DATE - scc_wday$COLLECTION_DATE
scc_wday$ReceiveToResult <- scc_wday$VERIFIED_DATE - scc_wday$RECEIVE_DATE
scc_wday[c("CollectToResult", "CollectToReceive", "ReceiveToResult")] <- lapply(scc_wday[c("CollectToResult", "CollectToReceive", "ReceiveToResult")], as.numeric, units = "mins")
scc_wday$AddOn <- ifelse(difftime(scc_wday$ORDERING_DATE, scc_wday$RECEIVE_DATE, units = "mins") > 5, "AddOn", "Original")

scc_wday$ReceiveResultTarget <- ifelse(tat_targets$Priority[match(scc_wday$Test, tat_targets$Test)] == "All" & 
                                         tat_targets$`Pt Setting`[match(scc_wday$Test, tat_targets$Test)] == "All", tat_targets$ReceiveToResultTarget[match(scc_wday$Test, tat_targets$Test)], "Other")

scc_wday$Concate1 <- paste(scc_wday$Test, scc_wday$MasterPriority)
scc_wday$Concate2 <- paste(scc_wday$Test, scc_wday$MasterPriority, scc_wday$MasterSetting)



tat_targets$Concate <- ifelse(tat_targets$Priority == "All" & tat_targets$`Pt Setting` == "All", tat_targets$Test,
                              ifelse(tat_targets$Priority != "All" & tat_targets$`Pt Setting` == "All", paste(tat_targets$Test, tat_targets$Priority),
                                     paste(tat_targets$Test, tat_targets$Priority, tat_targets$`Pt Setting`)))

scc_wday$ReceiveResultTarget <- ifelse(!is.na(match(scc_wday$Concate2, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(scc_wday$Concate2, tat_targets$Concate)], 
                                ifelse(!is.na(match(scc_wday$Concate1, tat_targets$Concate)), tat_targets$ReceiveToResultTarget[match(scc_wday$Concate1, tat_targets$Concate)],
                                       tat_targets$ReceiveToResultTarget[match(scc_wday$Test, tat_targets$Concate)]))

scc_wday$CollectResultTarget <- ifelse(!is.na(match(scc_wday$Concate2, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(scc_wday$Concate2, tat_targets$Concate)], 
                                ifelse(!is.na(match(scc_wday$Concate1, tat_targets$Concate)), tat_targets$CollectToResultTarget[match(scc_wday$Concate1, tat_targets$Concate)],
                                       tat_targets$CollectToResultTarget[match(scc_wday$Test, tat_targets$Concate)]))

df <- scc_wday[ , c("Test", "MasterPriority", "MasterSetting", "ReceiveResultTarget", "CollectResultTarget")]
x <- unique(df)
