# This code renders the 30-day lookback trended dashboard

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
rmarkdown::render(paste0("OI-Team-Folder/",
                         "ED_Weekly_Dashboard.Rmd"), 
                  output_file = paste0(
                    substr(user_directory, 1,
                           nchar(user_directory) - nchar("/Data")),
                    "/MSHS Lab OI Dashboards",
                    "/MSHS Lab PI Dashboard ",
                    format(Sys.Date(), "%m-%d-%y")))
