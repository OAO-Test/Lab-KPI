# Code for converting existing SCC and Sunquest repository from Excel to RDS

#Required packages: run these every time you run the code
library(readxl)
library(dplyr)
library(lubridate)
library(reshape2)
library(stringr)
library(writexl)

# Clear all existing variables
rm(list = ls())

# Set user directory based on drive mappings
if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab KPI/Data")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/", 
                           "Service Lines/Lab KPI/Data")
}
user_path <- paste0(user_directory, "/*.*")

# Import existing repository
existing_repo <- read_excel(
  choose.files(default = paste0(user_directory, "/SCC Sunquest Historical Repo",
                                "/*.*"), 
               caption = "Select historical repository for SCC and Sunquest 
               data"), 
  sheet = 1, col_names = TRUE)

# Save existing repository as .RDS file
saveRDS(existing_repo, file = paste0(user_directory, 
                                     "/SCC Sunquest Historical Repo/",
                                     "Hist Repo ",
                                     format(min(existing_repo$ResultDate),
                                            "%m%d%y"), 
                                     " to ",
                                     format(max(existing_repo$ResultDate),
                                            "%m%d%y"),
                                     " Created ",
                                     format(Sys.Date(), "%m%d%y"),
                                     ".RDS"))

# Read in .RDS file and check structure
rds_repo <- readRDS(file = choose.files(
  default = paste0(user_directory, "/SCC Sunquest Historical Repo",
                   "/*.*")))

test_repo <- rbind(rds_repo, existing_repo)

