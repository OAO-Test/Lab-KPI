
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("xlsx")
#install.packages("readxl")

#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)

#-------------------------------holiday/weekend-------------------------------#

#Determine if yesterday was a holiday/weekend 

#Change the format for the date into timeDate format to be ready for the next function
Yesterday <- as.timeDate(format(Sys.Date()-1,"%m/%d/%Y"))

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
  PP_Weekday <- data.farme(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
}

#-----------Cytology Backlog Excel Files-----------#
#For the backlog files the read excel is starting from the second row
#Also I made sure to remove the last line

Cytology_Backlog <- read_excel(choose.files(caption = "Select Cytology Backlog Report"), skip = 1, 1)
Cytology_Backlog <- data.frame(Cytology_Backlog[-nrow(Cytology_Backlog),], stringsAsFactors = FALSE)


#------------------------------Data Pre-Processing------------------------------#

#Cytology
#Keep the cyto gyn and cyto non-gyn

Cytology_Weekday <- PP_Weekday[which(PP_Weekday$spec_group=="CYTO NONGYN" | PP_Weekday$spec_group=="CYTO GYN"),]
Cytology_NoT_Weekday <- PP_Not_Weekday[which(PP_Not_Weekday$spec_group=="CYTO NONGYN" | PP_Not_Weekday$spec_group=="CYTO GYN"),]

#Surgical Pathology
#Keep the Breast and GI

GI_Codes <- data.frame(read_excel(choose.files(caption = "Select Cytology Backlog Report"), 1), stringsAsFactors = FALSE)

#Merge the exclusion/inclusion cloumn into PP Dataset
PP_Weekday_Excl <- merge(x = PP_Weekday, y= GI_Codes, all.x = TRUE)
Surgical_Pathology_Weekday <- PP_Weekday_Excl[which(((PP_Weekday_Excl$spec_group=="GI") &(PP_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Weekday_Excl$spec_group=="Breast"),]

PP_Not_Weekday_Excl <- merge(x = PP_Not_Weekday, y= GI_Codes, all.x = TRUE)
Surgical_Pathology_Not_Weekday <- PP_Not_Weekday_Excl[which(((PP_Not_Weekday_Excl$spec_group=="GI") &(PP_Not_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Not_Weekday_Excl$spec_group=="Breast"),]























