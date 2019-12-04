
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("readxl")
#install.packages("rmarkdown")

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


#------------------------------Data Pre-Processing------------------------------#

#Using Rev Center to determine patient setting
Patient_Setting <- data.frame(read_excel(choose.files(caption = "Select Cytology Backlog Report"), sheet = "final"), stringsAsFactors = FALSE)

#vlookup the Rev_Center and its corresponding patient setting for the PowerPath Data
PP_Weekday_PS <- merge(x=PP_Weekday, y=Patient_Setting, all.x = TRUE ) 
PP_Not_Weekday_PS <- merge(x=PP_Not_Weekday, y=Patient_Setting, all.x = TRUE )

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

Cytology_Not_Weekday$Case_created_date <- as.POSIXct(Cytology_Not_Weekday$Case_created_date,format='%m/%d/%y %I:%M %p')
Cytology_Not_Weekday$Collection_Date <- as.POSIXct(Cytology_Not_Weekday$Collection_Date,format='%m/%d/%y %I:%M %p')
Cytology_Not_Weekday$Received_Date <- as.POSIXct(Cytology_Not_Weekday$Received_Date,format='%m/%d/%y %I:%M %p')
Cytology_Not_Weekday$signed_out_date <- as.POSIXct(Cytology_Not_Weekday$signed_out_date,format='%m/%d/%y %I:%M %p')

#add columns for calculations: collection to signed out and received to signed out
#order to signed out
Cytology_Weekday$Collection_to_signed_out <- as.numeric(difftime(Cytology_Weekday$signed_out_date, Cytology_Weekday$Collection_Date, units = "days"))
Cytology_Not_Weekday$Collection_to_signed_out <- as.numeric(difftime(Cytology_Not_Weekday$signed_out_date, Cytology_Not_Weekday$Collection_Date, units = "days"))

#recieve to signed out
Cytology_Weekday$Received_to_signed_out <- as.numeric(difftime(Cytology_Weekday$signed_out_date, Cytology_Weekday$Received_Date, units = "days"))
Cytology_Not_Weekday$Received_to_signed_out <- as.numeric(difftime(Cytology_Not_Weekday$signed_out_date, Cytology_Not_Weekday$Received_Date, units = "days"))

#find any negative calculation and remove them
Cytology_Weekday <- Cytology_Weekday[!(Cytology_Weekday$Collection_to_signed_out<=0 | Cytology_Weekday$Received_to_signed_out<=0),]
Cytology_Not_Weekday <- Cytology_Not_Weekday[!(Cytology_Not_Weekday$Collection_to_signed_out<=0 | Cytology_Not_Weekday$Received_to_signed_out<=0),]


#Surgical Pathology

#Upload the exclusion vs inclusion criteria associated with the GI codes
GI_Codes <- data.frame(read_excel(choose.files(caption = "Select Cytology Backlog Report"), 1), stringsAsFactors = FALSE)

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

Surgical_Pathology_Not_Weekday$Case_created_date <- as.POSIXct(Surgical_Pathology_Not_Weekday$Case_created_date,format='%m/%d/%y %I:%M %p')
Surgical_Pathology_Not_Weekday$Collection_Date <- as.POSIXct(Surgical_Pathology_Not_Weekday$Collection_Date,format='%m/%d/%y %I:%M %p')
Surgical_Pathology_Not_Weekday$Received_Date <- as.POSIXct(Surgical_Pathology_Not_Weekday$Received_Date,format='%m/%d/%y %I:%M %p')
Surgical_Pathology_Not_Weekday$signed_out_date <- as.POSIXct(Surgical_Pathology_Not_Weekday$signed_out_date,format='%m/%d/%y %I:%M %p')

#add columns for calculations: collection to signed out and received to signed out
#order to signed out
Surgical_Pathology_Weekday$Collection_to_signed_out <- as.numeric(difftime(Surgical_Pathology_Weekday$signed_out_date, Surgical_Pathology_Weekday$Collection_Date, units = "days"))
Surgical_Pathology_Not_Weekday$Collection_to_signed_out <- as.numeric(difftime(Surgical_Pathology_Not_Weekday$signed_out_date, Surgical_Pathology_Not_Weekday$Collection_Date, units = "days"))

#recieve to signed out
Surgical_Pathology_Weekday$Received_to_signed_out <- as.numeric(difftime(Surgical_Pathology_Weekday$signed_out_date, Surgical_Pathology_Weekday$Received_Date, units = "days"))
Surgical_Pathology_Not_Weekday$Received_to_signed_out <- as.numeric(difftime(Surgical_Pathology_Not_Weekday$signed_out_date, Surgical_Pathology_Not_Weekday$Received_Date, units = "days"))

#find any negative calculation and remove them
Surgical_Pathology_Weekday <- Surgical_Pathology_Weekday[!(Surgical_Pathology_Weekday$Collection_to_signed_out<=0 | Surgical_Pathology_Weekday$Received_to_signed_out<=0),]
Surgical_Pathology_Not_Weekday <- Surgical_Pathology_Not_Weekday[!(Surgical_Pathology_Not_Weekday$Collection_to_signed_out<=0 | Surgical_Pathology_Not_Weekday$Received_to_signed_out<=0),]




















