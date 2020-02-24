#Required packages: run these everytime you run the code
library(timeDate)
library(readxl)
library(bizdays)
library(dplyr)
library(lubridate)
library(reshape2)
library(knitr)
library(kableExtra)
library(formattable)
library(rmarkdown)

#-------------------------------holiday/weekend-------------------------------#
# Get today and yesterday's date
Today <- as.timeDate(format(Sys.Date(),"%m/%d/%Y"))
#Today <- as.timeDate(as.Date("02/10/2020", format = "%m/%d/%Y"))

#Determine if yesterday was a holiday/weekend 
#get yesterday's DOW
Yesterday <- as.timeDate(format(Sys.Date()-1,"%m/%d/%Y"))
#Yesterday <- as.timeDate(as.Date("02/09/2020", format = "%m/%d/%Y"))

#Get yesterday's DOW
Yesterday_Day <- dayOfWeek(Yesterday)

#Remove Good Friday from MSHS Holidays
NYSE_Holidays <- as.Date(holidayNYSE(year = 1990:2100))
GoodFriday <- as.Date(GoodFriday())
MSHS_Holiday <- NYSE_Holidays[GoodFriday != NYSE_Holidays]

#Determine whether yesterday was a holiday/weekend
Holiday_Det <- isHoliday(Yesterday, holidays = MSHS_Holiday)

#Set up a default calendar for collect to received TAT calculations for Pathology and Cytology
create.calendar("MSHS_working_days", MSHS_Holiday, weekdays=c("saturday","sunday"))
bizdays.options$set(default.calendar="MSHS_working_days")


# Set working directory
user_wd <- "J:\\deans\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data"
setwd(user_wd)

file_list_SP <- list.files(path = paste0(user_wd, "\\AP & Cytology Signed Cases Reports"), pattern = "^(KPI REPORT\\ - \\RAW DATA V3).+(2020)\\-(01|02){1}\\-[0-9]{2}")

SP_list <- lapply(file_list_SP, function(x) read_excel(path = paste0(user_wd, "\\AP & Cytology Signed Cases Reports\\", x),skip = 1))

SP_Dataframe_combined <- bind_rows(SP_list)
SP_Dataframe_combined <- SP_Dataframe_combined[!SP_Dataframe_combined$Facility ==  "Page -1 of 1",]

# Import analysis reference data starting with test codes for SCC and Sunquest
# reference_file <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data\\Code Reference\\Analysis Reference 2020-01-22.xlsx"
reference_file <- choose.files(caption = "Select analysis reference file")

# To Be Updated in Analysis reference file
#-----------Patient Setting Excel File-----------#
#Using Rev Center to determine patient setting
Patient_Setting <- data.frame(read_excel(reference_file, sheet = "AP_Patient Setting"), stringsAsFactors = FALSE)

#-----------Anatomic Pathology Targets Excel File-----------#
TAT_Targets <- data.frame(read_excel(reference_file, sheet = "AP_TAT Targets"), stringsAsFactors = FALSE)

#-----------GI Codes Excel File-----------#
#Upload the exclusion vs inclusion criteria associated with the GI codes
GI_Codes <- data.frame(read_excel(reference_file, sheet = "GI_Codes"), stringsAsFactors = FALSE)

# create an extra column for old Fcaility and change the facility from MSSM to MSH 
SP_Dataframe_combined$Facility_Old <- SP_Dataframe_combined$Facility
SP_Dataframe_combined$Facility[SP_Dataframe_combined$Facility_Old == "MSS"] <- "MSH" 

#------------------------------Extract the All Breast, GI Specimens, cyto gyn and cyto nongyn Data Only------------------------------#

#Merge the exclusion/inclusion cloumn into the combined PP data

SP_Dataframe_combined_Exc <- merge(x = SP_Dataframe_combined, y= GI_Codes, all.x = TRUE)

SP_Dataframe_combined_Exc <- SP_Dataframe_combined_Exc[which(((SP_Dataframe_combined_Exc$spec_group=="GI") &(SP_Dataframe_combined_Exc$GI.Codes.Must.Include.in.Analysis..All.GI.Biopsies.=="Include")) | SP_Dataframe_combined_Exc$spec_group=="Breast" | SP_Dataframe_combined_Exc$spec_group=="CYTO NONGYN" | SP_Dataframe_combined_Exc$spec_group=="CYTO GYN"),]

#------------------------------Data Pre-Processing------------------------------#
############Create a function for Data pre-processing############

pre_processing_historical <- function(Raw_Data){
  if (is.null(Raw_Data) || nrow(Raw_Data)==0){
    Raw_Data_PS <- NULL
    Raw_Data_New <- NULL
    Summarized_Table <- NULL
  } else {
    #vlookup the Rev_Center and its corresponding patient setting for the PowerPath Data
    Raw_Data_PS <- merge(x=Raw_Data, y=Patient_Setting, all.x = TRUE ) 
    
    #vlookup targets based on spec_group and patient setting
    Raw_Data_New <- merge(x=Raw_Data_PS, y=TAT_Targets, all.x = TRUE, by = c("spec_group","Patient.Setting"))
    
    #Change all Dates into POSIXct format to start the calculations
    
    Raw_Data_New[c("Case_created_date","Collection_Date","Received_Date","signed_out_date")] <- lapply(Raw_Data_New[c("Case_created_date","Collection_Date","Received_Date","signed_out_date")], as.POSIXct,tz="", format='%m/%d/%y %I:%M %p')
    
    #add columns for calculations: collection to signed out and received to signed out
    #collection to signed out
    #All days in the calendar
    Raw_Data_New$Collection_to_signed_out <- as.numeric(difftime(Raw_Data_New$signed_out_date, Raw_Data_New$Collection_Date, units = "days"))
    
    #recieve to signed out
    #without weekends and holidays
    Raw_Data_New$Received_to_signed_out <- bizdays(Raw_Data_New$Received_Date, Raw_Data_New$signed_out_date)
    
    Summarized_Table <- summarise(group_by(Raw_Data_New,Spec_code, spec_group, Facility,Patient.Setting, Rev_ctr ,as.Date(signed_out_date),weekdays(as.Date(signed_out_date)),Received.to.signed.out.target..Days.,Collected.to.signed.out.target..Days.),  No_Cases_Signed = n(), Lab_Metric_TAT_Avg = round(mean(Received_to_signed_out, na.rm = TRUE),0), Lab_Metric_TAT_med = round(median(Received_to_signed_out, na.rm = TRUE),0), Lab_Metric_TAT_sd =round(sd(Received_to_signed_out, na.rm = TRUE),1), Lab_Metric_within_target = format(round(sum(Received_to_signed_out <= Received.to.signed.out.target..Days., na.rm = TRUE)/sum(Received_to_signed_out >= 0, na.rm = TRUE),2)), Patient_Metric_TAT_avg=format(ceiling(mean(Collection_to_signed_out, na.rm = TRUE))), Patient_Metric_TAT_med = round(median(Collection_to_signed_out, na.rm = TRUE),0), Patient_Metric_TAT_sd =round(sd(Collection_to_signed_out, na.rm = TRUE),1))
    
  }
  return(Summarized_Table)
}

Historical_Data_Summarized <- pre_processing_historical(SP_Dataframe_combined_Exc)

write.csv(Historical_Data_Summarized, "Historical_Repo_Surgical_Pathology.csv")




