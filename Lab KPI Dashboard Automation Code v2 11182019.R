
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("readxl")
#install.packages("bizdays")
#install.packages("rmarkdown")
#install.packages("tinytex")
#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(bizdays)
library(dplyr)
library(reshape2)
library(rmarkdown)
library(tinytex)
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
  Cytology_Cases_Signed <- NULL
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
BREAST_Lab_Metric <- summarise(group_by(Surgical_Pathology_Weekday[Surgical_Pathology_Weekday$spec_group=="BREAST",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 5)/sum(Received_to_signed_out >= 0),2)))
BREAST_Lab_Metric <- dcast(BREAST_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )

GI_Lab_Metric <- summarise(group_by(Surgical_Pathology_Weekday[Surgical_Pathology_Weekday$spec_group=="GI",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 3)/sum(Received_to_signed_out >= 0),2)))
GI_Lab_Metric <- dcast(GI_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )

if (is.null(PP_Not_Weekday)){
  BREAST_Lab_Metric_Not_Weekday <- NULL
  GI_Lab_Metric_Not_Weekday <-NULL
} else {
  BREAST_Lab_Metric_Not_Weekday <- summarise(group_by(Surgical_Pathology_Not_Weekday[Surgical_Pathology_Not_Weekday$spec_group=="BREAST",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 5)/sum(Received_to_signed_out >= 0),2)))
  BREAST_Lab_Metric_Not_Weekday <- dcast(BREAST_Lab_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
  
  GI_Lab_Metric_Not_Weekday <- summarise(group_by(Surgical_Pathology_Not_Weekday[Surgical_Pathology_Not_Weekday$spec_group=="GI",],spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= 3)/sum(Received_to_signed_out >= 0),2)))
  GI_Lab_Metric_Not_Weekday <- dcast(GI_Lab_Metric_Not_Weekday, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
}
















