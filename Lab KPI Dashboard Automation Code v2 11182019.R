
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
library(rJava)
#### Set-Up #### 
options(java.parameters = "-Xmx1000m")
################
library(xlsx)
library(readxl)
library(dplyr)

#-------------------------------holiday/weekend-------------------------------#

#Determine if yesterday was a holiday/weekend 

#get yesterday's DOW
# Yesterday_Day <- weekdays(as.Date(Sys.Date()-1))
Yest <- as.Date("11/25/2019", format = "%m/%d/%Y")
Yesterday_Day <- weekdays(Yest) #Rename as Yest_DOW
#Change the format for the date into timeDate format to be ready for the next function
Yesterday <- as.timeDate(Yest)

#Excludes Good Friday from the NYSE Holidays
NYSE_Holidays <- as.Date(holidayNYSE())
GoodFriday <- as.Date(GoodFriday())
MSHS_Holiday <- NYSE_Holidays[GoodFriday != NYSE_Holidays]

#This function determines whether yesterday was a holiday/weekend or no
Holiday_Det <- isHoliday(Yesterday, holidays = MSHS_Holiday)

#------------------------------Read Excel sheets------------------------------#
#The if-statement below helps in determining how many excel files are required

#-----------SCC Excel Files-----------#

if((Holiday_Det) & (Yesterday_Day =="Monday")){
  SCC_Holiday_Monday <- read_excel(choose.files(caption = "Select SCC Holiday Report"), sheet = 1, col_names = TRUE)
  SCC_Sunday <- read_excel(choose.files(caption = "Select SCC Sunday Report"), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(choose.files(caption = "Select SCC Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(rbind(SCC_Holiday_Monday,SCC_Sunday),SCC_Saturday)
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sunday")){
  SCC_Sunday <- read_excel(choose.files(caption = "Select SCC Sunday Report"), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(choose.files(caption = "Select SCC Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(SCC_Sunday,SCC_Saturday)
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Monday")|(Yesterday_Day !="Sunday"))){
  SCC_Holiday_Weekday <- read_excel(choose.files(caption = "Select SCC Holiday Report"), sheet = 1, col_names = TRUE)
  SCC_Not_Weekday <- SCC_Holiday_Weekday
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
} else {
  SCC_Weekday <- read_excel(choose.files(caption = "Select SCC Weekday Report"), sheet = 1, col_names = TRUE)
}

#-----------SunQuest Excel Files-----------#

if((Holiday_Det) & (Yesterday_Day =="Monday")){
  SQ_Holiday_Monday <- read_excel(choose.files(caption = "Select SunQuest Holiday Report"), sheet = 1, col_names = TRUE)
  SQ_Sunday <- read_excel(choose.files(caption = "Select SunQuest Sunday Report"), sheet = 1, col_names = TRUE)
  SQ_Saturday <- read_excel(choose.files(caption = "Select SunQuest Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(rbind(SQ_Holiday_Monday,SQ_Sunday),SQ_Saturday)
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sunday")){
  SQ_Sunday <- read_excel(choose.files(caption = "Select SunQuest Sunday Report"), sheet = 1, col_names = TRUE)
  SQ_Saturday <- read_excel(choose.files(caption = "Select SunQuest Saturday Report"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(SQ_Sunday,SQ_Saturday)
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Monday")|(Yesterday_Day !="Sunday"))){
  SQ_Holiday_Weekday <- read_excel(choose.files(caption = "Select SunQuest Holiday Report"), sheet = 1, col_names = TRUE)
  SQ_Not_Weekday <- SQ_Holiday_Weekday
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
} else {
  SQ_Weekday <- read_excel(choose.files(caption = "Select SunQuest Weekday Report"), sheet = 1, col_names = TRUE)
}

#-----------PowerPath Excel Files-----------#
#For the powerpath files the read excel is starting from the second row
#Also I made sure to remove the last line
if((Holiday_Det) & (Yesterday_Day =="Monday")){
  PP_Holiday_Monday <- read.xlsx(choose.files(caption = "Select PowerPath Holiday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Holiday_Monday <- PP_Holiday_Monday[-nrow(PP_Holiday_Monday),]
  PP_Sunday <- read.xlsx(choose.files(caption = "Select PowerPath Sunday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read.xlsx(choose.files(caption = "Select PowerPath Saturday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- rbind(rbind(PP_Holiday_Monday,PP_Sunday),PP_Saturday)
  PP_Weekday <- read.xlsx(choose.files(caption = "Select PowerPath Weekday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
} else if ((Holiday_Det) & (Yesterday_Day =="Sunday")){
  PP_Sunday <- read.xlsx(choose.files(caption = "Select PowerPath Sunday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read.xlsx(choose.files(caption = "Select PowerPath Saturday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- rbind(PP_Sunday,PP_Saturday)
  PP_Weekday <- read.xlsx(choose.files(caption = "Select PowerPath Weekday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
} else if ((Holiday_Det) & ((Yesterday_Day !="Monday")|(Yesterday_Day !="Sunday"))){
  PP_Holiday_Weekday <- read.xlsx(choose.files(caption = "Select PowerPath Holiday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Holiday_Weekday <- PP_Holiday_Weekday[-nrow(PP_Holiday_Weekday),]
  PP_Not_Weekday <- PP_Holiday_Weekday
  PP_Weekday <- read.xlsx(choose.files(caption = "Select PowerPath Weekday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
} else {
  PP_Weekday <- read.xlsx(choose.files(caption = "Select PowerPath Weekday Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
}

#-----------Cytology Backlog Excel Files-----------#
#For the backlog files the read excel is starting from the second row
#Also I made sure to remove the last line

Cytology_Backlog <- read.xlsx(choose.files(caption = "Select Cytology Backlog Report"), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
Cytology_Backlog <- Cytology_Backlog[-nrow(Cytology_Backlog),]


#------------------------------Data Pre-Processing------------------------------#
# Import analysis reference data starting with test codes for SCC and Sunquest
test_code <- read_excel("Analysis REference 2019-12-02.xlsx", sheet = "TestNames")
scc_icu <- read_excel("Analysis Reference 2019-12-02.xlsx", sheet = "SCC_ICU")

scc_wday <- SCC_Weekday
sq_wday <- SQ_Weekday

# Format data fields
scc_wday[c("Ward", "WARD_NAME", 
           "REQUESTING_DOC", 
           "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", 
           "COLLECT_CENTER_ID", "SITE", "CLINIC_TYPE")] <- lapply(scc_wday[c("Ward", "WARD_NAME", 
                                                                             "REQUESTING_DOC", 
                                                                             "GROUP_TEST_ID", "TEST_ID", "TEST_NAME", "PRIORITY", 
                                                                             "COLLECT_CENTER_ID", "SITE", "CLINIC_TYPE")], as.factor)

scc_wday$ORDERING_DATE <- as.POSIXct(scc_wday$ORDERING_DATE, tz = "", format = "%Y-%m-%d %H:%M:%S.%f")

scc_wday$Ward <- as.factor(scc_wday$Ward)
scc_wday$WARD_NAME <- as.factor(scc_wday$WARD_NAME)

scc_wday <- left_join(scc_wday, test_code[ , c("Test", "SCC_TestID")], by = c("TEST_ID" = "SCC_TestID"))
scc_wday$TestIncl <- ifelse(is.na(scc_wday$Test), FALSE, TRUE)
scc_wday <- scc_wday[scc_wday$TestIncl == TRUE, ]
scc_wday$WardandName <- paste(scc_wday$Ward, scc_wday$WARD_NAME)
scc_wday <- left_join(scc_wday, scc_icu[ , c("Concatenate", "ICU?")], by = c("WardandName" = "Concatenate"))
