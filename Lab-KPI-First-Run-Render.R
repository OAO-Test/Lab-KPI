rmarkdown::render("lab_kpi_first_run_dashboard.Rmd", 
                  output_file = paste0(
                    substr(user_directory, 1,
                           str_locate(user_directory, "/Data")[1] - 1),
                    "/Dashboard Drafts",
                    "/Lab KPI Dashboard Pre KPI Form ",
                    format(Sys.Date(), "%m-%d-%y")))
