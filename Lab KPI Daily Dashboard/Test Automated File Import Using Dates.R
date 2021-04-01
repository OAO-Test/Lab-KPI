# Code to test automatically identifying reports based on dates
# Eliminates need for user to select files individually

#Required packages: run these every time you run the code
library(timeDate)
library(readxl)
library(bizdays)
library(dplyr)
library(lubridate)
library(reshape2)
library(knitr)
# library(gdtools)
library(kableExtra)
library(formattable)
library(rmarkdown)
library(stringr)
library(writexl)


#Clear existing history
rm(list = ls())
#-------------------------------holiday/weekend-------------------------------#
# Get today and yesterday's date
# today <- Sys.Date()
today <- as.Date("3/15/21", format = "%m/%d/%y")

#Determine if yesterday was a holiday/weekend
#get yesterday's DOW
yesterday <- today  - 1

#Get yesterday's DOW
yesterday_day <- wday(yesterday, label = TRUE, abbr = TRUE)

#Remove Good Friday from MSHS Holidays
nyse_holidays <- as.Date(holidayNYSE(year = 1990:2100))
good_friday <- as.Date(GoodFriday())
mshs_holiday <- nyse_holidays[good_friday != nyse_holidays]

#Determine whether yesterday was a holiday/weekend
#holiday_det <- isHoliday(yesterday, holidays = mshs_holiday)
holiday_det <- isHoliday(as.timeDate(yesterday), holidays = mshs_holiday)

#Set up a calendar for collect to received TAT calculations for Path & Cyto
create.calendar("MSHS_working_days", mshs_holiday,
                weekdays = c("saturday", "sunday"))
bizdays.options$set(default.calendar = "MSHS_working_days")


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


# Determine dates of most recent weekday and weekend/holiday if applicable
if (((holiday_det) & (yesterday_day == "Mon")) |
    ((yesterday_day == "Sun") & (isHoliday(as.timeDate(yesterday - 2),
                                           holidays = mshs_holiday)))) {
  # Scenario 1: Mon Holiday or Friday Holiday
  # Save scenario
  scenario <- 1
  include_not_wday <- TRUE
  # Determine weekday and weekend/holiday dates
  wkday_date <- today - 4
  wkend_holiday_date <- seq(from = today - 3,
                            to = yesterday,
                            by = 1)
} else if ((holiday_det) & (yesterday_day == "Sun")) {
  # Scenario 2: Regular Monday (Need to select 3 files)
  # Save scenario
  scenario <- 2
  include_not_wday <- TRUE
  # Determine weekday and weekend/holiday dates
  wkday_date <- today - 3
  wkend_holiday_date <- seq(from = today - 2,
                            to = yesterday,
                            by = 1)
} else if ((holiday_det) & ((yesterday_day != "Mon") |
                            (yesterday_day != "Sun"))) {
  #Scenario 3: Midweek holiday
  # Save scenario
  scenario <- 3
  include_not_wday <- TRUE
  # Determine weekday and weekend/holiday dates
  wkday_date <- today - 2
  wkend_holiday_date <- yesterday
} else {#Scenario 4: Tue-Fri without holidays
  # Save scenario
  scenario <- 4
  include_not_wday <- FALSE
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday
  wkend_holiday_date <- NULL
}

# Create regular expressions for beginning each file type
scc_pattern_start <- "^(Doc){1}.+"
sq_pattern_start <- "^(KPI_Daily_TAT_Report){1}.*"
pp_pattern_start <- "^(KPI REPORT - RAW DATA V4_V2){1}.*"
epic_pattern_start <- "^(MSHS Pathology Orders EPIC){1}.*"
cyto_backlog_pattern_start <- "^(KPI REPORT - CYTOLOGY PENDING CASES){1}.*"

# Format most recent weekday date as it would appear in saved files
# This will always be the day after the most recent weekday since reports
# show cases resulted on prior day
wkday_date_file_format <- format(wkday_date + 1, "%Y-%m-%d")

# Find SCC file for most recent weekday 
scc_weekday_file <- list.files(
  path = paste0(user_directory, "/SCC CP Reports"),
  pattern = paste0(scc_pattern_start,
                   "(",
                   wkday_date_file_format,
                   ".xlsx)$"))

# Import SCC file for most recent weekday
scc_weekday <- read_excel(path =
                            paste0(user_directory,
                                   "/SCC CP Reports/",
                                   scc_weekday_file),
                          sheet = 1, col_names = TRUE)

# Find Sunquest file for most recent weekday
sq_weekday_file <- list.files(
  path = paste0(user_directory, "/SUN CP Reports"),
  pattern = paste0(sq_pattern_start,
                  "(",
                  wkday_date_file_format,
                  ".xls)$"))

# Import Sunquest file for most recent weekday
sq_weekday <- suppressWarnings(read_excel(path = 
                                            paste0(user_directory,
                                                   "/SUN CP Reports/",
                                                   sq_weekday_file),
                                          sheet = 1, col_names = TRUE))

# Find Powerpath file for signed out cases for most recent weekday
pp_weekday_file <- list.files(
  path = paste0(user_directory, "/AP & Cytology Signed Cases Reports"),
  pattern = paste0(pp_pattern_start,
                   "(",
                   wkday_date_file_format,
                   ".xls)$"))

# Import Powerpath file for most recent weekday
pp_weekday <- read_excel(path =
                           paste0(user_directory,
                                  "/AP & Cytology Signed Cases Reports/",
                                  pp_weekday_file),
                         skip = 1, 1)

pp_weekday <- data.frame(pp_weekday[-nrow(pp_weekday), ],
                         stringsAsFactors = FALSE)

# Find Epic Cytology file for signed out cases for most recent weekday
epic_weekday_file <- list.files(
  path = paste0(user_directory, "/EPIC Cytology"),
  pattern = paste0(epic_pattern_start,
                   "(",
                   wkday_date_file_format,
                   ".xlsx)$"))

# Import Epic Cytology file for most recent weekday
epic_weekday <- read_excel(path =
                             paste0(user_directory,
                                    "/EPIC Cytology/",
                                    epic_weekday_file),
                           1)

# Find Cytology backlog file received today
cyto_backlog_file <- list.files(
  path = paste0(user_directory, "/Cytology Backlog Reports"),
  pattern = paste0(cyto_backlog_pattern_start,
                   "(",
                   format(today, "%Y-%m-%d"),
                   ".xls)$"))

# Import Cytology backlog file received today
cyto_backlog_raw <- read_excel(path =
                                     paste0(user_directory,
                                            "/Cytology Backlog Reports/",
                                            cytology_file),
                                   skip = 1, 1)

cyto_backlog_raw <- data.frame(cyto_backlog_raw[-nrow(cyto_backlog_raw), ],
                               stringsAsFactors = FALSE)

# Find and import weekend/holiday reports
if (is.null(wkend_holiday_date)) {
  wkend_holiday_date_file_format <- NULL
  
  scc_not_weekday <- NULL
  sq_not_weekday <- NULL
  pp_not_weekday <- NULL
  epic_not_weekday <- NULL
  
} else {
  # Determine weekend/holiday dates as formatted in data files
  wkend_holiday_date_file_format <- sapply(wkend_holiday_date + 1,
                                           format, "%Y-%m-%d")
  
  # Find SCC files for weekend/holiday dates
  scc_not_weekday_file <- list.files(
    path = paste0(user_directory, "/SCC CP Reports"),
    pattern = paste0(scc_pattern_start,
                     "(",
                     wkend_holiday_date_file_format,
                     ".xlsx)$",
                     collapse = "|"))
  
  # Read SCC files for weekend/holiday dates and bind into single dataframe
  scc_not_weekday_list <- sapply(scc_not_weekday_file,
                                function(x)
                                  read_excel(path =
                                               paste0(user_directory,
                                                      "/SCC CP Reports/",
                                                      x),
                                             sheet = 1,
                                             col_names = TRUE),
                                simplify = FALSE)
  
  scc_not_weekday <- bind_rows(scc_not_weekday_list)
  
  # Find Sunquest files for weekend/holiday dates
  sq_not_weekday_file <- list.files(
    path = paste0(user_directory, "/SUN CP Reports"),
    pattern = paste0(sq_pattern_start,
                     "(",
                     wkend_holiday_date_file_format,
                     ".xls)$",
                     collapse = "|"))
  
  # Read Sunquest files for weekend/holiday dates and bind into single dataframe
  sq_not_weekday_list <- sapply(sq_not_weekday_file,
                                 function(x)
                                   suppressWarnings(read_excel(path =
                                                paste0(user_directory,
                                                       "/SUN CP Reports/",
                                                       x),
                                              sheet = 1,
                                              col_names = TRUE)),
                                 simplify = FALSE)
  
  sq_not_weekday <- bind_rows(sq_not_weekday_list)
  
  # Find Powerpath signed out cases for weekend/holiday dates
  pp_not_weekday_file <- list.files(
    path = paste0(user_directory, "/AP & Cytology Signed Cases Reports"),
    pattern = paste0(pp_pattern_start,
                     "(",
                     wkend_holiday_date_file_format,
                     ".xls)$",
                     collapse = "|"))
  
  # Read Powerpath files for weekend/holiday dates and bind into single dataframe
  pp_not_weekday_list <- sapply(pp_not_weekday_file,
                                function(x)
                                  read_excel(path =
                                               paste0(user_directory,
                                                      "/AP & Cytology Signed Cases Reports/",
                                                      x),
                                             skip = 1,	  
                                             sheet = 1),
                                simplify = FALSE)
  
  pp_not_weekday_list <- lapply(pp_not_weekday_list,
                                function(x)
                                  data.frame(x[-nrow(x), ],
                                             stringsAsFactors = FALSE))
  
  pp_not_weekday <- bind_rows(pp_not_weekday_list)
  
  # Find Epic Cytology signed out cases for weekend/holiday dates
  epic_not_weekday_file <- list.files(
    path = paste0(user_directory, "/EPIC Cytology"),
    pattern = paste0(epic_pattern_start,
                     "(",
                     wkend_holiday_date_file_format,
                     ".xlsx)$",
                     collapse = "|"))
  # Read Epic Cytology files for weekend/holiday dates and bind into single dataframe
  epic_not_weekday_list <- sapply(epic_not_weekday_file,
                                  function(x)
                                    read_excel(path =
                                                 paste0(user_directory,
                                                        "/EPIC Cytology/",
                                                        x),
                                               sheet = 1),
                                  simplify = FALSE)
  
  epic_not_weekday_list <- lapply(epic_not_weekday_list,
                                  function(x)
                                    data.frame(x, stringsAsFactors = FALSE))
  
  epic_not_weekday <- bind_rows(epic_not_weekday_list)
}


