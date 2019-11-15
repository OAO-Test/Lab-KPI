# Install packages
# install.packages(ggplot2)

# Load libraries
library(ggplot2)
library(dplyr)
library(readxl)


# Get and set working directory
getwd()
path <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/Service Lines/Lab KPI/Data"
setwd(path)

# Select four daily reports for analysis
scc_path <- grep("SCC CP Reports$", list.dirs(path), value = TRUE)
scc_mask <- paste0(scc_path, "/*.*")
scc_file <- choose.files(default = scc_mask, caption = "Select Today's SCC Report", multi = FALSE)
scc_raw <- read_excel(scc_file)

sun_path <- grep("SUN CP Reports$", list.dirs(path), value = TRUE)
sun_mask <- paste0(sun_path, "/*.*")
sun_file <- choose.files(default = sun_mask, caption = "Select Today's Sunquest Report", multi = FALSE)
sun_raw <- read_excel(sun_file)

path_cyto_path <- grep("*Signed Cases Reports$", list.dirs(path), value = TRUE)
path_cyto_mask <- paste0(path_cyto_path, "/*.*")
path_cyto_file <- choose.files(default = path_cyto_mask, caption = "Select Today's Pathology & Cytology Report", multi = FALSE)
path_cyto_raw <- read_excel(path_cyto_file, skip = 1)
path_cyto_raw <- path_cyto_raw[-nrow(path_cyto_raw), ]

cyto_backlog_path <- grep("*Backlog Reports$", list.dirs(path), value = TRUE)
cyto_backlog_mask <- paste0(cyto_backlog_path, "/*.*")
cyto_backlog_file <- choose.files(default = cyto_backlog_mask, caption = "Select Today's Cytology Backlog Report", multi = FALSE)
cyto_backlog_raw <- read_excel(cyto_backlog_file, skip = 1)
cyto_backlog_raw <- cyto_backlog_raw[-nrow(cyto_backlog_raw), ]

