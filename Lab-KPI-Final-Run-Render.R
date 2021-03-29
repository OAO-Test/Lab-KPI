rmarkdown::render("lab_kpi_second_run_dashboard.Rmd", 
                  output_file = paste("Lab KPI Dashboard Final ",
                                      format(Sys.Date(), "%m-%d-%y")))
