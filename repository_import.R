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
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday - 3
  wkend_holiday_date <- seq(from = yesterday - 2,
                            to = yesterday,
                            by = 1)
} else if ((holiday_det) & (yesterday_day == "Sun")) {
  # Scenario 2: Regular Monday (Need to select 3 files)
  # Save scenario
  scenario <- 2
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
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday - 1
  wkend_holiday_date <- yesterday
} else {#Scenario 4: Tue-Fri without holidays
  # Save scenario
  scenario <- 4
  # Determine weekday and weekend/holiday dates
  wkday_date <- yesterday
  wkend_holiday_date <- NULL
}

# Import data repositories --------------------
# Import CP historical repository
cp_repo <- readRDS(file = choose.files(
  default = paste0(user_directory, "/CP Historical Repo/*.*"),
  caption = "Select CP Historical Repository"))

# Subset CP data for most recent weekday
cp_wday_summary <- cp_repo %>%
  filter(ResultDate %in% wkday_date)

# Subset CP data for weekend/holidays
cp_not_wday_summary <- cp_repo %>%
  filter(ResultDate %in% wkend_holiday_date)
