# Code for subsetting CP historical repository data based on relevant dates ----

# Subset CP data for most recent weekday
cp_wday_summary <- cp_repo %>%
  filter(ResultDate %in% wkday_date)

# Subset CP data for weekend/holidays
cp_not_wday_summary <- cp_repo %>%
  filter(ResultDate %in% wkend_holiday_date)
