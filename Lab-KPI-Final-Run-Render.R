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
rmarkdown::render("lab_kpi_second_run_dashboard.Rmd", 
                  output_file = paste0(
                    substr(user_directory, 1,
                           str_locate(user_directory, "/Data")[1] - 1),
                    "/Dashboard Drafts",
                    "/Lab KPI Dashboard Final ",
                    format(Sys.Date(), "%m-%d-%y")))
