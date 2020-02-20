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
library(ggplot2)

rm(list = ls())

user_wd <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data\\SCC Sunquest Script Repo"
user_path <- paste0(user_wd, "\\*.*")
setwd(user_wd)

scc_sun_hist <- read_excel(choose.files(default = user_path, caption = "Select Historical Repository"), sheet = 1, col_names = TRUE)

scc_sun_hist$ResultDate <- as.Date(scc_sun_hist$ResultDate, format = "%m/%d/%y")
end_date <- max(scc_sun_hist$ResultDate)
scc_sun_hist <- scc_sun_hist[scc_sun_hist$ResultDate < end_date, ]
start_date <- max(scc_sun_hist$ResultDate) - 29

dashboard_summary <- scc_sun_hist %>%
  group_by(Site, ResultDate, Test, DashboardPriority, DashboardSetting, ReceiveResultTarget, CollectResultTarget) %>%
  summarize(ResultedVolume = sum(TotalResulted), ResultedVolumeWithTAT = sum(TotalResultedTAT), ReceiveResultInTarget = sum(TotalReceiveResultInTarget), CollectResultInTarget = sum(TotalCollectResultInTarget))

dashboard_summary$ReceiveResultPercent <- dashboard_summary$ReceiveResultInTarget / dashboard_summary$ResultedVolumeWithTAT
dashboard_summary$CollectResultPercent <- dashboard_summary$CollectResultInTarget / dashboard_summary$ResultedVolumeWithTAT
dashboard_summary$ReceiveResultDefectRate <- (dashboard_summary$ResultedVolumeWithTAT - dashboard_summary$ReceiveResultInTarget) / dashboard_summary$ResultedVolumeWithTAT
dashboard_summary$CollectResultDefectRate <- (dashboard_summary$ResultedVolumeWithTAT - dashboard_summary$CollectResultInTarget) / dashboard_summary$ResultedVolumeWithTAT

dashboard_summary$SafeThreshold <- ifelse(dashboard_summary$Test == "Rapid Flu" | dashboard_summary$Test == "C. diff", 1.0, 0.95)
dashboard_summary$NotSafeThreshold <- ifelse(dashboard_summary$Test == "Rapid Flu" | dashboard_summary$Test == "C. diff", 0.9, 0.8)

dashboard_summary_volume <- scc_sun_hist %>%
  group_by(Site, ResultDate, Test) %>%
  summarize(ResultedVolume = sum(TotalResulted))

troponin_test <- dashboard_summary[dashboard_summary$Test == "Troponin", ]
troponin_test <- troponin_test[troponin_test$DashboardSetting == "ED & ICU" | troponin_test$DashboardSetting == "IP Non-ICU", ]

troponin_test$ResultDate <- as.Date(troponin_test$ResultDate, format = "%m/%d/%y")


ggplot(data = troponin_test[troponin_test$ResultDate >= start_date & troponin_test$Site == "MSH", ]) +
  geom_point(aes(x = ResultDate, y = ReceiveResultPercent, color = DashboardSetting, shape = DashboardSetting), size = 3) + 
  geom_line(aes(x = ResultDate, y = ReceiveResultPercent, color = DashboardSetting), linetype = "dashed") +
  labs(title = "MSH Troponin Percentage of Labs within Target TAT:\n Receive to Result", x = "Date",  y = "Percentage of Labs", color = "Patient Setting", shape = "Patient Setting") +
  theme_bw() +
  theme(plot.title = element_text(hjust = 0.5), legend.position = "bottom", axis.text.x = element_text(angle = 30, hjust = 1)) +
  scale_color_manual(values = c("ED & ICU" = "#221f72", "IP Non-ICU" = "#00AEEF")) +
  annotate("rect", xmin = start_date-1, xmax = max(troponin_test$ResultDate)+1, ymin = -Inf, ymax = 0.8, fill = "red", alpha = 0.2) +
  annotate("rect", xmin = start_date-1, xmax = max(troponin_test$ResultDate)+1, ymin = 0.8, ymax = 0.95, fill = "orange", alpha = 0.2) +
  annotate("rect", xmin = start_date-1, xmax = max(troponin_test$ResultDate)+1, ymin = 0.95, ymax = Inf, fill = "green", alpha = 0.2) +
  annotate("text", x = max(troponin_test$ResultDate), y = 0.8, label = "Not Safe", color = "red", vjust = 2, hjust = .75) +
  annotate("text", x = max(troponin_test$ResultDate), y = 0.8, label = "At Risk", color = "orange", vjust = -2, hjust = .75) +
  annotate("text", x = max(troponin_test$ResultDate), y = 0.95, label = "Safe", color = "#006600", vjust = -1, hjust = .75) +
  scale_x_date(limits = c(start_date -1, max(troponin_test$ResultDate)+1), breaks = seq(start_date, max(troponin_test$ResultDate), by = 3), date_minor_breaks = "1 day", date_labels = "%m/%d/%y", expand = c(0, 0, 0, 0)) +
  scale_fill_manual(name = "Status Definitions", values = c("Not Safe" = "red", "At Risk" = "orange", "Safe" = "green"))

