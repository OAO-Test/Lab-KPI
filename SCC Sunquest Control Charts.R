# Sample control charts

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
library(stringr)
library(writexl)

rm(list = ls())

getwd()
setwd(".\\SCC Sunquest Script Repo")

scc_sun_hist <- read_excel(choose.files(caption = "Select Historical Repository"), sheet = 1, col_names = TRUE)

scc_sun_hist$ResultDate <- as.Date(scc_sun_hist$ResultDate, format = "%m/%d/%y")
start_date <- max(scc_sun_hist$ResultDate) - 29
