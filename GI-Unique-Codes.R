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

# Set working directory
user_wd <- "J:\\deans\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data"
setwd(user_wd)

file_list_SP <- list.files(path = paste0(user_wd, "\\AP & Cytology Signed Cases Reports"), pattern = "^(KPI REPORT\\ - \\RAW DATA V(3|2){1}).+(2020|2019)\\-(01|02|03|09|10|11|12){1}\\-[0-9]{2}")

SP_list <- lapply(file_list_SP, function(x) read_excel(path = paste0(user_wd, "\\AP & Cytology Signed Cases Reports\\", x),skip = 1))

SP_Dataframe_combined <- bind_rows(SP_list)
SP_Dataframe_combined <- SP_Dataframe_combined[!SP_Dataframe_combined$Facility ==  "Page -1 of 1",]

#subset the data to get only GI

SP_Dataframe_GI_combined <- SP_Dataframe_combined[SP_Dataframe_combined$spec_group == "GI",]

#get the unique GI codes and save in CSV file
GI_Codes_Unique <- unique(SP_Dataframe_GI_combined$Spec_code)
write.csv(GI_Codes_Unique, "GI_Codes_Unique_Oct2019-Mar2020.csv")



