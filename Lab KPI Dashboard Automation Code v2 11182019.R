
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
library(rJava)
#### Set-Up #### 
options(java.parameters = "-Xmx1000m")
################
library(xlsx)
library(readxl)

#-------------------------------holiday/weekend-------------------------------#

#Determine if yesterday was a holiday/weekend 

#get yesterday's DOW
Yesterday_Day <- weekdays(as.Date(Sys.Date()-1))
#Change the format for the date into timeDate format to be ready for the next function
Yesterday <- as.timeDate(format(Sys.Date()-1,"%m/%d/%Y"))

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
  SCC_Holiday_Monday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SCC_Sunday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(rbind(SCC_Holiday_Monday,SCC_Sunday),SCC_Saturday)
  SCC_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sunday")){
  SCC_Sunday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(SCC_Sunday,SCC_Saturday)
  SCC_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Monday")|(Yesterday_Day !="Sunday"))){
  SCC_Holiday_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SCC_Not_Weekday <- SCC_Holiday_Weekday
  SCC_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
} else {
  SCC_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
}

#-----------SunQuest Excel Files-----------#

if((Holiday_Det) & (Yesterday_Day =="Monday")){
  SQ_Holiday_Monday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SQ_Sunday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SQ_Saturday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(rbind(SQ_Holiday_Monday,SQ_Sunday),SQ_Saturday)
  SQ_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sunday")){
  SQ_Sunday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SQ_Saturday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(SQ_Sunday,SQ_Saturday)
  SQ_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Monday")|(Yesterday_Day !="Sunday"))){
  SQ_Holiday_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
  SQ_Not_Weekday <- SQ_Holiday_Weekday
  SQ_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
} else {
  SQ_Weekday <- read_excel(file.choose(), sheet = 1, col_names = TRUE)
}

#-----------PowerPath Excel Files-----------#
#For the powerpath files the read excel is starting from the second row
#Also I made sure to remove the last line
if((Holiday_Det) & (Yesterday_Day =="Monday")){
  PP_Holiday_Monday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Holiday_Monday <- PP_Holiday_Monday[-nrow(PP_Holiday_Monday),]
  PP_Sunday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- rbind(rbind(PP_Holiday_Monday,PP_Sunday),PP_Saturday)
  PP_Weekday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
} else if ((Holiday_Det) & (Yesterday_Day =="Sunday")){
  PP_Sunday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- rbind(PP_Sunday,PP_Saturday)
  PP_Weekday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
} else if ((Holiday_Det) & ((Yesterday_Day !="Monday")|(Yesterday_Day !="Sunday"))){
  PP_Holiday_Weekday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Holiday_Weekday <- PP_Holiday_Weekday[-nrow(PP_Holiday_Weekday),]
  PP_Weekday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
} else {
  PP_Weekday <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
  PP_Weekday <- PP_Weekday[-nrow(PP_Weekday),]
}

#-----------Cytology Backlog Excel Files-----------#
#For the backlog files the read excel is starting from the second row
#Also I made sure to remove the last line

Cytology_Backlog <- read.xlsx(file.choose(), startRow = 2, header = TRUE, stringsAsFactors=FALSE, 1)
Cytology_Backlog <- Cytology_Backlog[-nrow((Cytology_Backlog),)]

#------------------------------Data Pre-Processing------------------------------#

#-----------Pathology-----------#


#-----------Cytology-----------#


#------------------------------Turn Around Times Calculations------------------------------#

#-----------Pathology-----------#
#----Weekdays----#



#----Not Weekdays----#


#-----------Cytology-----------#
#----Weekdays----#



#----Not Weekdays----#


#------------------------------Backlog Calculations------------------------------#


