rmarkdown::render("lab_kpi_first_run_dashboard.Rmd", 
                  output_file = paste("Lab KPI Dashboard Pre KPI Form ",
                                      format(Sys.Date(), "%m-%d-%y")))

