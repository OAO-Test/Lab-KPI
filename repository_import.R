#######
# Code for importing the REPO data needed to carry-on the second run
# for the lab KPI daily dashboard with different logic based on the DOW.
# Imported data includes:
# 1. CP repo data for clinical pathology
# 2. AP repo data for anatomic pathology
# 3. AP Backlog repo data for anatomic pathology
#######

# Determine dates of most recent weekday and weekend/holiday if applicable
if (((holiday_det) & (yesterday_day == "Mon")) |
    ((yesterday_day == "Sun") & (isHoliday(as.timeDate(yesterday - 2))))) {
  # Scenario 1: Mon Holiday or Friday Holiday
  # Save scenario
  scenario <- 1
  include_not_wday <- TRUE
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday - 3
  wkend_holiday_date <- seq(from = yesterday - 2,
                            to = yesterday,
                            by = 1)
} else if ((holiday_det) & (yesterday_day == "Sun")) {
  # Scenario 2: Regular Monday (Need to select 3 files)
  # Save scenario
  scenario <- 2
  include_not_wday <- TRUE
  # Determine weekday and weekend/holiday dates
  wkday_date <- today - 3
  wkend_holiday_date <- seq(from = yesterday - 1,
                            to = yesterday,
                            by = 1)
} else if ((holiday_det) & ((yesterday_day != "Mon") |
                            (yesterday_day != "Sun"))) {
  #Scenario 3: Midweek holiday
  # Save scenario
  scenario <- 3
  include_not_wday <- TRUE
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday - 1
  wkend_holiday_date <- yesterday
} else {#Scenario 4: Tue-Fri without holidays
  # Save scenario
  scenario <- 4
  include_not_wday <- FALSE
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday
  wkend_holiday_date <- NULL
}

# Import data repositories --------------------
# Import CP historical repository
cp_repo <- readRDS(file = choose.files(
  default = paste0(user_directory, "/CP Historical Repo/*.*"),
  caption = "Select CP Historical Repository"))

# Import AP historical repository
ap_repo <-
  readRDS(file =
            choose.files(
              default =
                paste0(user_directory,
                       "/AP & Cytology Historical Repo",
                       "/*.*"),
              caption = "Select Pathology and Cytology Historical Repository"))

# Import cytology backlog historical repository
backlog_repo <-
  readRDS(file =
            choose.files(
              default =
                paste0(user_directory,
                       "/AP & Cytology Historical Repo",
                       "/*.*"),
              caption = "Select Cytology Backlog Historical Repository"))

# Subset CP data for most recent weekday
cp_wday_summary <- cp_repo %>%
  filter(ResultDate %in% wkday_date)

# Subset CP data for weekend/holidays
cp_not_wday_summary <- cp_repo %>%
  filter(ResultDate %in% wkend_holiday_date)

# Subset AP data for most recent weekday
 ap_wday_summary <- ap_repo %>%
  filter((report_date_only - 1) %in% wkday_date)

# Subset most recent weekday AP data to get cyto
cyto_table_weekday_summarized <- ap_wday_summary %>%
  filter(Spec_group %in% cyto_spec_group)

# Subset most recent weekday AP data to get patho
sp_table_weekday_summarized <- ap_wday_summary %>%
  filter(Spec_group %in% patho_spec_group)

# Subset cyto backlog data for most recent weekday
cyto_backlog_summarized <- backlog_repo %>%
  filter((Report_Date - 1) %in% wkday_date)

# Determine resulted date for weekday labs
wkday_date <- format(wkday_date, format = "%a %m/%d/%y")

# Determine resulted date for weekend labs
if (!is.null(cp_not_wday_summary)) {
  not_wday_result_date <- sort(unique(cp_not_wday_summary$ResultDate))
  wkend_holiday_result_date <- ifelse(
    length(not_wday_result_date) == 1,
    format(not_wday_result_date, format = "%a %m/%d/%y"),
    paste0(format(not_wday_result_date[1],
                  format = "%a %m/%d/%y"),
           "-",
           format(not_wday_result_date[length(not_wday_result_date)],
                  format = "%a %m/%d/%y")))
} else {
  wkend_holiday_result_date <- NULL
}
