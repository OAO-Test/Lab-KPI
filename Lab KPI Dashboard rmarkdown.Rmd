---
title: "MSHS Laboratory Daily Status Report"
output: html_document
---

```{r setup, include=FALSE}
#install.packages("knitr")
knitr::opts_chunk$set(echo = TRUE)
```


```{r, echo=FALSE, warning=FALSE, message=FALSE}
#The purpose of this code: is to build an automated version of the cytology/pathology dashboard in R
#Coder: Asala Erekat

#-------------------------------Install packages-------------------------------#

#Install packages only the first time you run the code
#install.packages("timeDate")
#install.packages("readxl")
#install.packages("bizdays")
#install.packages("dplyr")
#install.packages("reshape2")
#install.packages("rmarkdown")
#install.packages("tinytex")
#install.packages("kableExtra")
#install.packages("formattable")

#Clear the Enviroment before running the code
rm(list = ls())
#-------------------------------Required packages-------------------------------#

#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(bizdays)
library(dplyr)
library(reshape2)
library(rmarkdown)
library(tinytex)
library(kableExtra)
library(formattable)
library(knitr)
#-------------------------------holiday/weekend-------------------------------#
Today <- as.timeDate(format(Sys.Date(),"%m/%d/%Y"))

#Determine if yesterday was a holiday/weekend 

#Change the format for the date into timeDate format to be ready for the next function
Yesterday <- as.timeDate(format(Sys.Date()-4,"%m/%d/%Y"))

#get yesterday's DOW
Yesterday_Day <- dayOfWeek(Yesterday)

#Excludes Good Friday from the NYSE Holidays
NYSE_Holidays <- as.Date(holidayNYSE(year = 1990:2100))
GoodFriday <- as.Date(GoodFriday())
MSHS_Holiday <- NYSE_Holidays[GoodFriday != NYSE_Holidays]

#This function determines whether yesterday was a holiday/weekend or no
Holiday_Det <- isHoliday(Yesterday, holidays = MSHS_Holiday)

#Set up a default calendar for collect to received TAT calculations
create.calendar("MSHS_working_days", MSHS_Holiday, weekdays=c("saturday","sunday"))
bizdays.options$set(default.calendar="MSHS_working_days")

#------------------------------Read Excel sheets------------------------------#
#The if-statement below helps in determining how many excel files are required

#-----------PowerPath Excel Files-----------#
#For the powerpath files the read excel is starting from the second row
#Also I made sure to remove the last line
if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){
  PP_Holiday_Monday_or_Friday <- read_excel(choose.files(caption = "Select PowerPath Holiday Report"), skip = 1, 1)
  PP_Holiday_Monday_or_Friday  <- PP_Holiday_Monday_or_Friday[-nrow(PP_Holiday_Monday_or_Friday),]
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
Cytology_Backlog_Raw <- read_excel(choose.files(caption = "Select Cytology Backlog Report"),skip =1 , 1)
Cytology_Backlog_Raw <- data.frame(Cytology_Backlog_Raw[-nrow(Cytology_Backlog_Raw),], stringsAsFactors = FALSE)

#-----------Patient Setting Excel File-----------#
#Using Rev Center to determine patient setting
Patient_Setting <- data.frame(read_excel(choose.files(caption = "Select Patient Setting Dataframe"), sheet = "final"), stringsAsFactors = FALSE)

#-----------Anatomic Pathology Targets Excel File-----------#
TAT_Targets <- data.frame(read_excel(choose.files(caption = "AP TAT Targets"), 1), stringsAsFactors = FALSE)

#-----------GI Codes Excel File-----------#
#Upload the exclusion vs inclusion criteria associated with the GI codes
GI_Codes <- data.frame(read_excel(choose.files(caption = "Select GI Codes dataframe"), 1), stringsAsFactors = FALSE)

#-----------Create table template-----------#
columns <- c("spec_group","Patient.Setting","No_Cases_Signed","MSH.x", "BIMC.x","MSQ.x", "MSS.x","NYEE.x","PACC.x","R.x","SL.x", "KH.x","BIMC.y", "MSH.y", "MSQ.y", "MSS.y","NYEE.y","PACC.y","R.y","SL.y", "KH.y")
#Cyto
Table_Template_Cyto <- data.frame(matrix(ncol=21, nrow = 4)) 
colnames(Table_Template_Cyto) <- columns
Table_Template_Cyto[1] <- c('CYTO GYN','CYTO GYN','CYTO NONGYN', 'CYTO NONGYN')
Table_Template_Cyto[2] <- c('IP', 'Amb')

#Patho
Table_Template_Patho <- data.frame(matrix(ncol=21, nrow = 4)) 
colnames(Table_Template_Patho) <- columns
Table_Template_Patho[1] <- c('Breast','Breast','GI', 'GI')
Table_Template_Patho[2] <- c('IP', 'Amb')

#------------------------------Extract the Cytology GYN and NON-GYN Data Only------------------------------#
#Cytology
#Keep the cyto gyn and cyto non-gyn

Cytology_Weekday_Raw <- PP_Weekday[which(PP_Weekday$spec_group=="CYTO NONGYN" | PP_Weekday$spec_group=="CYTO GYN"),]
Cytology_Not_Weekday_Raw <- PP_Not_Weekday[which(PP_Not_Weekday$spec_group=="CYTO NONGYN" | PP_Not_Weekday$spec_group=="CYTO GYN"),]

#------------------------------Extract the All Breast and GI Specimens Data Only------------------------------#

#Merge the exclusion/inclusion cloumn into the modified powerpath Dataset for weekdays and not weekdays

PP_Weekday_Excl <- merge(x = PP_Weekday, y= GI_Codes, all.x = TRUE)
Surgical_Pathology_Weekday <- PP_Weekday_Excl[which(((PP_Weekday_Excl$spec_group=="GI") &(PP_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Weekday_Excl$spec_group=="Breast"),]

PP_Not_Weekday_Excl <- merge(x = PP_Not_Weekday, y= GI_Codes, all.x = TRUE)
Surgical_Pathology_Not_Weekday <- PP_Not_Weekday_Excl[which(((PP_Not_Weekday_Excl$spec_group=="GI") &(PP_Not_Weekday_Excl$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | PP_Not_Weekday_Excl$spec_group=="Breast"),]

#-----------Reporting Dates-----------#
#1. Weekday Date
Report_Date_Weekday <- unique(as.timeDate(format(PP_Weekday$signed_out_date, "%m/%d/%Y")))
Report_Date_Weekday_Day <- unique(dayOfWeek(Report_Date_Weekday))

#2. Weekend/Holiday Dates
if (is.null(PP_Not_Weekday)){
  Report_Date_Weekend <- NULL
  Report_Date_Weekend_Day <- NULL
}else {
  Report_Date_Weekend <- unique(as.timeDate(format(PP_Not_Weekday$signed_out_date, "%m/%d/%Y")))
  Report_Date_Weekend_Day <- unique(dayOfWeek(Report_Date_Weekend))
}

#------------------------------Data Pre-Processing------------------------------#
############Create a function for Data pre-processing############

pre_processing <- function(Raw_Data){
  if (is.null(Raw_Data) || nrow(Raw_Data)==0){
    Raw_Data_PS <- NULL
    #Raw_Data_PS_Target <- NULL
    Raw_Data_New <- NULL
    Raw_Data_New_Cases_Signed <- NULL
    Raw_Data_New_Patient_Metric <- NULL
    Raw_Data_New_Lab_Metric <- NULL
    Processed_Data_Table <- NULL
  } else {
    #vlookup the Rev_Center and its corresponding patient setting for the PowerPath Data
    Raw_Data_PS <- merge(x=Raw_Data, y=Patient_Setting, all.x = TRUE ) 

    #vlookup targets based on spec_group and patient setting
    Raw_Data_New <- merge(x=Raw_Data_PS, y=TAT_Targets, all.x = TRUE, by = c("spec_group","Patient.Setting"))
    
    #If there are any N/A in the dates, remove the sample
#    Raw_Data_New <- Raw_Data_PS_Target[!(is.na(Raw_Data_PS_Target$Case_created_date)|is.na(Raw_Data_PS_Target$Collection_Date) | is.na(Raw_Data_PS_Target$Received_Date) | is.na(Raw_Data_PSv_Target$signed_out_date)),]
    
    #Change all Dates into POSIXct format to start the calculations
    
    Raw_Data_New[c("Case_created_date","Collection_Date","Received_Date","signed_out_date")] <- lapply(Raw_Data_New[c("Case_created_date","Collection_Date","Received_Date","signed_out_date")], as.POSIXct, format='%m/%d/%y %I:%M %p')
    
    #add columns for calculations: collection to signed out and received to signed out
    #collection to signed out
    #All days in the calendar
    Raw_Data_New$Collection_to_signed_out <- as.numeric(difftime(Raw_Data_New$signed_out_date, Raw_Data_New$Collection_Date, units = "days"))
    
    #recieve to signed out
    #without weekends and holidays
    Raw_Data_New$Received_to_signed_out <- bizdays(Raw_Data_New$Received_Date, Raw_Data_New$signed_out_date)
    
    #find any negative calculation and remove them
#    Raw_Data_New <- Raw_Data_New[!(Raw_Data_New$Collection_to_signed_out<=0 | Raw_Data_New$Received_to_signed_out<=0),]
    
    #Calculate Average for Collection to signed out and number of cases signed
    
    Raw_Data_New_Cases_Signed <- summarise(group_by(Raw_Data_New,spec_group, Patient.Setting), No_Cases_Signed = n())
    
    Raw_Data_New_Patient_Metric <- summarise(group_by(Raw_Data_New,spec_group, Facility,Patient.Setting), Avg_Collection_to_Signed_out=format(round(mean(Collection_to_signed_out, na.rm = TRUE),2)))
    
    Raw_Data_New_Patient_Metric <- dcast(Raw_Data_New_Patient_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Avg_Collection_to_Signed_out" )
    
    #Calculate % Receive to result TAT within target
    Raw_Data_New_Lab_Metric <-summarise(group_by(Raw_Data_New,spec_group, Facility,Patient.Setting), Received_to_Signed_out_within_target = format(round(sum(Received_to_signed_out <= Received.to.signed.out.target..Days., na.rm = TRUE)/sum(Received_to_signed_out >= 0, na.rm = TRUE),2)))
    
    Raw_Data_New_Lab_Metric <- dcast(Raw_Data_New_Lab_Metric, spec_group + Patient.Setting ~ Facility, value.var = "Received_to_Signed_out_within_target" )
    
    #here I will merge number of cases signed, received to result TAT, and acollect to result TAT calcs into one table
    
    #Cytology Weekday table
    Processed_Data_Table <- left_join(full_join(Raw_Data_New_Cases_Signed, Raw_Data_New_Lab_Metric), Raw_Data_New_Patient_Metric, by = c("spec_group", "Patient.Setting"))
    Processed_Data_Table <- Processed_Data_Table[!(Processed_Data_Table$Patient.Setting == "Other"),]
  }
  return(Processed_Data_Table)
}
    
Cytology_Table_Weekday <- pre_processing(Cytology_Weekday_Raw)
Cytology_Table_Not_Weekday <- pre_processing(Cytology_Not_Weekday_Raw)
Surgical_Pathology_Table_Weekday <- pre_processing(Surgical_Pathology_Weekday)
Surgical_Pathology_Table_Not_Weekday <- pre_processing(Surgical_Pathology_Not_Weekday)

############Create a function for Table standardization for cyto and patho############
#To add all the missing rows and columns
columns_order <- c("spec_group","Patient.Setting","No_Cases_Signed","MSH.x", "MSS.x","MSQ.x","BIMC.x","PACC.x","KH.x","R.x","SL.x", "NYEE.x","MSH.y", "MSS.y","MSQ.y","BIMC.y","PACC.y","KH.y","R.y","SL.y", "NYEE.y")
#Cytology table
Table_Merging_Cyto <- function(Cytology_Table){
  if (is.null(Cytology_Table)){
    Cytology_Table_New <- NULL
    Cytology_Table_New2 <- NULL
  } else {
    #first step is merging the table template with the cytology table and this will include all of the missing columns.

    Cytology_Table_New <- merge(x = Table_Template_Cyto, y= Cytology_Table, all.y=TRUE)
    
    #second step is merging the table with all of the columns with only the first two columns of the template to include all the missing rows
    Cytology_Table_New2 <- merge(x = Table_Template_Cyto[c(1,2)], y= Cytology_Table_New, all.x = TRUE, by = c("spec_group", "Patient.Setting"))

    rows_order_Cyto <- factor(rownames(Cytology_Table_New2),levels = c(2,1,4,3))
    Cytology_Table_New2 <- Cytology_Table_New2[order(rows_order_Cyto), columns_order]
  }
  return(Cytology_Table_New2)
}

Cytology_Table_Weekday_New2 <- Table_Merging_Cyto(Cytology_Table_Weekday)
Cytology_Table_Not_Weekday_New2 <- Table_Merging_Cyto(Cytology_Table_Not_Weekday)

#Pathology table
Table_Merging_Patho <- function(Surgical_Pathology_Table){
  if (is.null(Surgical_Pathology_Table)){
    Surgical_Pathology_Table_New <- NULL
    Surgical_Pathology_Table_New2 <- NULL
  } else {
    #first step is merging the table template with the cytology table and this will include all of the missing columns.

    Surgical_Pathology_Table_New <- merge(x = Table_Template_Patho, y= Surgical_Pathology_Table, all.y=TRUE)
    
    #second step is merging the table with all of the columns with only the first two columns of the template to include all the missing rows
    Surgical_Pathology_Table_New2 <- merge(x = Table_Template_Patho[c(1,2)], y= Surgical_Pathology_Table_New, all.x = TRUE, by = c("spec_group", "Patient.Setting"))
    rows_order_Patho <- factor(rownames(Surgical_Pathology_Table_New2),levels = c(2,1,4,3))
    Surgical_Pathology_Table_New2 <- Surgical_Pathology_Table_New2[order(rows_order_Patho), columns_order]
  }
  return(Surgical_Pathology_Table_New2)
}

Surgical_Pathology_Table_Weekday_New2 <- Table_Merging_Patho(Surgical_Pathology_Table_Weekday)
Surgical_Pathology_Table_Not_Weekday_New2 <- Table_Merging_Patho(Surgical_Pathology_Table_Not_Weekday)


############Create a function for conditional formatting############
Conditional_Formatting <- function(Table_New2){
  if (is.null(Table_New2)){
    Table_New3 <- NULL
  } else {
    
    Table_New2[,4:21] <- lapply(Table_New2[,4:21], as.numeric)
    Table_New2[,4:12] <- lapply(Table_New2[,4:12],percent, d=0)
    
    #steps for conditional formatting:
    
    Table_New3 <- melt(Table_New2, id =c("spec_group","Patient.Setting","No_Cases_Signed","MSH.y","MSS.y","MSQ.y","BIMC.y","PACC.y","KH.y","R.y","SL.y", "NYEE.y"))
    
    
    Table_New3 <- Table_New3 %>%
      mutate(value = ifelse(value > 0.9, cell_spec(value, "html", color  = "green"),
                       ifelse(value < 0.8, cell_spec(value, "html", color = "red"),
                              ifelse(is.na(value), cell_spec(value, "html", color = "grey"), cell_spec(value, "html", color = "orange")))))
    
    
    Table_New3 <- dcast(Table_New3,spec_group + Patient.Setting + No_Cases_Signed + BIMC.y + MSH.y + MSQ.y + MSS.y + NYEE.y + PACC.y + R.y + SL.y + KH.y ~ variable )
    Table_New3 <- merge(x=Table_New3, y=TAT_Targets[1:3])
    columns_order <- c("spec_group","Received.to.signed.out.target..Days.","Patient.Setting","No_Cases_Signed","MSH.x", "MSS.x","MSQ.x","BIMC.x","PACC.x","KH.x","R.x","SL.x", "NYEE.x","MSH.y", "MSS.y","MSQ.y","BIMC.y","PACC.y","KH.y","R.y","SL.y", "NYEE.y")
    rows_order <- factor(rownames(Table_New3),levels = c(2,1,4,3))
    Table_New3 <- Table_New3[order(rows_order), columns_order]
    Table_New3$Received.to.signed.out.target..Days.<- paste("<= ", Table_New3$Received.to.signed.out.target..Days., "days")
    row.names(Table_New3) <- NULL
  }
  return(Table_New3)
}

Cytology_Table_Weekday_New3 <- Conditional_Formatting(Cytology_Table_Weekday_New2)
Cytology_Table_Not_Weekday_New3 <- Conditional_Formatting(Cytology_Table_Not_Weekday_New2)
Surgical_Pathology_Table_Weekday_New3 <- Conditional_Formatting(Surgical_Pathology_Table_Weekday_New2)
Surgical_Pathology_Table_Not_Weekday_New3 <- Conditional_Formatting(Surgical_Pathology_Table_Not_Weekday_New2)

```



```{r, echo=FALSE, warning=FALSE, message=FALSE}
#Cytology backlog Calculation

#vlookup the Rev_Center and its corresponding patient setting for the PowerPath Data
Cytology_Backlog_PS <- merge(x=Cytology_Backlog_Raw, y=Patient_Setting, all.x = TRUE ) 

#vlookup targets based on spec_group and patient setting
Cytology_Backlog_PS_Target <- merge(x=Cytology_Backlog_PS, y=TAT_Targets, all.x = TRUE, by = c("spec_group","Patient.Setting"))

#Keep the cyto gyn and cyto non-gyn
Cytology_Backlog <- Cytology_Backlog_PS_Target[which(Cytology_Backlog_PS_Target$spec_group=="CYTO NONGYN" | Cytology_Backlog_PS_Target$spec_group=="CYTO GYN"),]

#Change all Dates into POSIXct format to start the calculations
Cytology_Backlog[c("Case_created_date","Collection_Date","Received_Date","signed_out_date")] <- lapply(Cytology_Backlog[c("Case_created_date","Collection_Date","Received_Date","signed_out_date")], as.POSIXct, format='%m/%d/%y %I:%M %p')

#Backlog Calculations: Date now - case created date
#without weekends and holidays
Cytology_Backlog$Backlog <- bizdays(Cytology_Backlog$Case_created_date, Today)

#Case Volume
Cytology_Backlog_Volume <- summarise(group_by(Cytology_Backlog,spec_group), Cyto_Backlog = format(round(sum(Backlog > Received.to.signed.out.target..Days., na.rm = TRUE),2)))

#Days of work
Cytology_case_Volume_DOW <- as.numeric(Cytology_Backlog_Volume$Cyto_Backlog[1])/80

#accessioned volume
#1. Find the date that we need to report --> the date of the last weekday
Accessioned_Date <- as.Date(Cytology_Weekday_Raw$signed_out_date[1], "%m/%d/%Y")
#2. count the accessioned volume that was accessioned on that date from the cytology report
Cytology_Weekday_Raw$Accessioned_Date_Only <- as.Date(Cytology_Weekday_Raw$Case_created_date, "%m/%d/%Y")
Cytology_Accessioned_Vol1 <- summarise(group_by(Cytology_Weekday_Raw,spec_group), Cyto_Acc_Vol1 = as.numeric(sum(Accessioned_Date == Accessioned_Date_Only, na.rm = TRUE)))
#3. count the accessioned volume that was accessioned on that date from the backlog report
Cytology_Backlog_Raw$Accessioned_Date_Only <- as.Date(Cytology_Backlog_Raw$Case_created_date, "%m/%d/%Y")
Cytology_Accessioned_Vol2 <- summarise(group_by(Cytology_Backlog_Raw,spec_group), Cyto_Acc_Vol2 = as.numeric(sum(Accessioned_Date == Accessioned_Date_Only, na.rm = TRUE)))
#4. sum the two counts
Cytology_Accessioned_Vol3 <- merge(x= Cytology_Accessioned_Vol1, y= Cytology_Accessioned_Vol2)
Cytology_Accessioned_Vol3$Total_Accessioned_Volume <- Cytology_Accessioned_Vol3$Cyto_Acc_Vol1 + Cytology_Accessioned_Vol3$Cyto_Acc_Vol2
Cytology_Accessioned_Vol3$Cyto_Acc_Vol1 <- NULL
Cytology_Accessioned_Vol3$Cyto_Acc_Vol2 <- NULL
Backlog_Accessioned_Table <- merge(Cytology_Accessioned_Vol3, Cytology_Backlog_Volume)

```


```{r, echo=FALSE}
Table_Formatting <- function (Table_New3){
  if (is.null(Table_New3)){
    Table_New3 <- NULL
  } else { 
    Table_New3%>%
    select(everything()) %>%
    kable(escape = F, align = "c",caption = paste("Report Creation Date:", Today), col.names = c("Case Type","Target","Setting","No. Cases Signed","MSH", "MSSM","MSQ","MSBI","PACC","MSB","MSW","MSSL", "NYEE","MSH", "MSSM","MSQ","MSBI","PACC","MSB","MSW","MSSL", "NYEE")) %>%
    kable_styling(c("hover","condensed","responsive"), full_width = FALSE, position = "center", row_label_position = "c", font_size = 11) %>%
    column_spec(5:13, background = "#00B9F2",include_thead = TRUE)%>%
    row_spec(1:4, background = "white")%>%
    column_spec(14:22, background = "#221F72", include_thead = TRUE)%>%
    row_spec(1:4, background = "white")%>%
    row_spec(0, color = "white") %>%
    collapse_rows(columns = c(1,2))%>%
    column_spec(1:4, color = "black", include_thead = TRUE)%>%
    add_header_above(c(" "= 4, "Receive to Result TAT within Target"= 9, "Average Collection to Result TAT (Calendar Days)"=9), background = c("white", "#00B9F2", "#221F72"), color= "white")%>%
    add_header_above(c("Status Definitions: Green >= 90%, Yellow >= 80% & < 90%, Red < 80%"=22))
  }
}
```


#### Surgical Pathology Efficiency Indicator (Labs Resulted on `r Report_Date_Weekday_Day` `r Report_Date_Weekday`)
```{r Surgical Pathology Efficiency Indicator Weekday, echo=FALSE, warning=FALSE, message=FALSE}
Table_Formatting(Surgical_Pathology_Table_Weekday_New3)
```
#### Cytology Efficiency Indicator (Labs Resulted on `r Report_Date_Weekday_Day` `r Report_Date_Weekday`)
```{r Cytology Efficiency Indicators Weekday, echo=FALSE, warning=FALSE, message=FALSE}
Table_Formatting(Cytology_Table_Weekday_New3)
```

#### Cytology Accessioned and Backlog Volume
```{r Cytology Backlog and Accessioned, echo=FALSE, warning=FALSE, message=FALSE}
Backlog_Accessioned_Table %>%
kable(escape = F, align = "c",caption = paste("Report Creation Date:", Today), col.names = c("Case Type","Cases Accessioned","Backlog Volume")) %>%
    kable_styling(c("hover","condensed","responsive"), full_width = FALSE, position = "left", row_label_position = "c", font_size = 11) %>%
    column_spec(2, background = "#00B9F2",include_thead = TRUE)%>%
    row_spec(1:2, background = "white")%>%
    column_spec(3, background = "#221F72", include_thead = TRUE)%>%
    row_spec(1:2, background = "white")%>%
    row_spec(0, color = "white") %>%
    column_spec(1, color = "black", include_thead = TRUE)
```

```{r Surgical Pathology Efficiency Indicators Weekend/Holiday, echo=FALSE, warning=FALSE, message=FALSE, eval = (!is.null(Surgical_Pathology_Table_Not_Weekday_New3))}
asis_output(paste('<h4> Surgical Pathology Efficiency Indicator (Labs Resulted Between the Period', Report_Date_Weekend_Day[length(Report_Date_Weekend_Day)], Report_Date_Weekend[length(Report_Date_Weekend)], '<h4> to', Report_Date_Weekend_Day[1], Report_Date_Weekend[1],')'))
Table_Formatting(Surgical_Pathology_Table_Not_Weekday_New3)
```

```{r Cytology Efficiency Indicators Weekend/Holiday, echo=FALSE, warning=FALSE, message=FALSE, eval = (!is.null(Cytology_Table_Not_Weekday_New3))}
asis_output(paste('<h4> Cytology Efficiency Indicator (Labs Resulted Between the Period', Report_Date_Weekend_Day[length(Report_Date_Weekend_Day)], Report_Date_Weekend[length(Report_Date_Weekend)], '<h4> to', Report_Date_Weekend_Day[1], Report_Date_Weekend[1],')'))
Table_Formatting(Cytology_Table_Not_Weekday_New3)
```









