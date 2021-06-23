# This code renders the first run markdown for the daily dashboard

# Clear environment
rm(list = ls())

# Determine directory
if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab Kpi/Data")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "Service Lines/Lab Kpi/Data")
}

# Render markdown file with dashboard code and save with today's date
rmarkdown::render("Lab KPI Daily Dashboard/Lab_KPI_First_Run_Dashboard.Rmd", 
                  output_file = paste0(
                    substr(user_directory, 1,
                           nchar(user_directory) - nchar("/Data")),
                    "/Dashboard Drafts",
                    "/Original Lab KPI Dashboard Pre KPI Form ",
                    format(Sys.Date(), "%m-%d-%y")))

