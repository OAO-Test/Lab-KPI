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
library(writexl)
library(gsubfn)
library(stringr)

#Clear existing history
rm(list = ls())
#-------------------------------holiday/weekend-------------------------------#
# Get today and yesterday's date

today <- Sys.Date()

# Select file/folder path for easier file selection and navigation

if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab Kpi/Data")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab Kpi/Data")
}

test_pp_data <- read_excel(
  path = paste0(user_directory,
                "\\AP & Cytology Signed Cases Reports\\",
                "KPI REPORT - RAW DATA V4_V2 2021-04-08.xls"),
           skip = 1)

test_pp_data <- test_pp_data[-nrow(test_pp_data),]

epic_data <- read_excel(
  path = paste0(user_directory,
                "\\EPIC Cytology\\",
                "MSHS Pathology Orders EPIC 2021-04-08.xlsx"))

cyto_sign_out_df <- test_pp_data %>%
  filter(spec_sort_order == "A" &
           spec_group %in% c("CYTO NONGYN", "CYTO GYN")) %>%
  mutate(SignOutDept = ifelse(str_detect(signed_out_Pathologist, "Interface"),
                              "Micro", "Cyto"))

epic_data_final <- epic_data %>%
  filter(LAB_STATUS %in% c("Final result", "Edited Result - FINAL"))

epic_data_spec <- epic_data_final %>%
  mutate(Case_no = SPECIMEN_ID) %>%
  select(Case_no)

epic_data_spec <- unique(epic_data_spec)

pp_data <- test_pp_data %>%
  mutate(Facility_Old = Facility,
         Facility = ifelse(Facility_Old == "MSS", "MSH",
                           ifelse(Facility_Old == "STL", "SL",
                                  Facility_Old)))

cyto_raw <- pp_data %>%
  filter(spec_sort_order == "A" &
           spec_group %in% c("CYTO NONGYN", "CYTO GYN"))

cyto_final <- merge(x = cyto_raw, y = epic_data_spec)

cyto_raw_new <- pp_data %>%
  filter(spec_sort_order == "A" &
           spec_group %in% c("CYTO NONGYN", "CYTO GYN")) %>%
  mutate(SignOutDept = ifelse(str_detect(signed_out_Pathologist, "Interface"),
                              "Microbiology", "Cytology"),
         FinalizedCase = Case_no %in% epic_data_spec$Case_no)


sp_raw <- test_pp_data %>%
  mutate(Facility_Old = Facility,
         Facility = ifelse(Facility_Old == "MSS", "MSH",
                           ifelse(Facility_Old == "STL", "SL",
                                  Facility_Old)),
         spec_group = ifelse(spec_group == "BREAST", "Breast", spec_group))











