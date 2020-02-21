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
lookback_period <- 30
start_date <- max(scc_sun_hist$ResultDate) - (lookback_period - 1)

dashboard_tat_summary <- scc_sun_hist %>%
  group_by(Site, ResultDate, Test, DashboardPriority, DashboardSetting, ReceiveResultTarget, CollectResultTarget) %>%
  summarize(ResultedVolume = sum(TotalResulted), ResultedVolumeWithTAT = sum(TotalResultedTAT), ReceiveResultInTarget = sum(TotalReceiveResultInTarget), CollectResultInTarget = sum(TotalCollectResultInTarget))

dashboard_tat_summary$ReceiveResultPercent <- dashboard_tat_summary$ReceiveResultInTarget / dashboard_tat_summary$ResultedVolumeWithTAT
dashboard_tat_summary$CollectResultPercent <- dashboard_tat_summary$CollectResultInTarget / dashboard_tat_summary$ResultedVolumeWithTAT
dashboard_tat_summary$ReceiveResultDefectRate <- (dashboard_tat_summary$ResultedVolumeWithTAT - dashboard_tat_summary$ReceiveResultInTarget) / dashboard_tat_summary$ResultedVolumeWithTAT
dashboard_tat_summary$CollectResultDefectRate <- (dashboard_tat_summary$ResultedVolumeWithTAT - dashboard_tat_summary$CollectResultInTarget) / dashboard_tat_summary$ResultedVolumeWithTAT

dashboard_tat_summary$SafeThreshold <- ifelse(dashboard_tat_summary$Test == "Rapid Flu" | dashboard_tat_summary$Test == "C. diff", 1.0, 0.95)
dashboard_tat_summary$NotSafeThreshold <- ifelse(dashboard_tat_summary$Test == "Rapid Flu" | dashboard_tat_summary$Test == "C. diff", 0.9, 0.8)

dashboard_vol_summary <- scc_sun_hist %>%
  group_by(Site, ResultDate, Test) %>%
  summarize(ResultedVolume = sum(TotalResulted))

troponin_test <- dashboard_tat_summary[dashboard_tat_summary$Test == "Troponin", ]
troponin_test <- troponin_test[troponin_test$DashboardSetting == "ED & ICU" | troponin_test$DashboardSetting == "IP Non-ICU", ]

troponin_test$ResultDate <- as.Date(troponin_test$ResultDate, format = "%m/%d/%y")


# Try recreating graph using geom_rect instead of annotate (better for legend)
ggplot() +
  geom_rect(aes(xmin = start_date - 1, xmax = max(troponin_test$ResultDate) + 1, ymin = -Inf, ymax = min(troponin_test$NotSafeThreshold), fill = "Not Safe", alpha = "Not Safe")) +
  geom_rect(aes(xmin = start_date - 1, xmax = max(troponin_test$ResultDate) + 1, ymin = min(troponin_test$NotSafeThreshold), ymax = min(troponin_test$SafeThreshold), fill = "At Risk", alpha = "At Risk")) +
  geom_rect(aes(xmin = start_date - 1, xmax = max(troponin_test$ResultDate) + 1, ymin = min(troponin_test$SafeThreshold), ymax = Inf, fill = "Safe", alpha = "Safe")) +
  geom_point(data = troponin_test[troponin_test$ResultDate >= start_date & troponin_test$Site == "MSH", ], aes(x = ResultDate, y = ReceiveResultPercent, color = DashboardSetting, shape = DashboardSetting), size = 3) + 
  geom_line(data = troponin_test[troponin_test$ResultDate >= start_date & troponin_test$Site == "MSH", ], aes(x = ResultDate, y = ReceiveResultPercent, color = DashboardSetting), linetype = "dashed") +
  labs(title = "MSH Troponin Percentage of Labs within Target TAT:\n Receive to Result", x = "Date",  y = "Percentage of Labs", color = "Setting", shape = "Setting") +
  theme_bw() +
  theme(plot.title = element_text(hjust = 0.5), legend.position = "right", legend.box = "vertical", legend.title = element_text(size = 10), legend.text = element_text(size = 8),axis.text.x = element_text(angle = 30, hjust = 1)) +
  scale_x_date(limits = c(start_date -1, max(troponin_test$ResultDate)+1), breaks = seq(start_date, max(troponin_test$ResultDate), by = 3), date_minor_breaks = "1 day", date_labels = "%m/%d/%y", expand = c(0, 0, 0, 0)) +
  scale_fill_manual(name = "Status", values = c("Not Safe" = "red", "At Risk" = "orange", "Safe" = "green"), breaks=c("Not Safe","At Risk","Safe")) + 
  scale_alpha_manual(name = "Status", values = c("Not Safe" = 0.2, "At Risk" = 0.2, "Safe" = 0.2), breaks=c("Not Safe","At Risk","Safe")) +
  scale_color_manual(name = "Setting", values = c("ED & ICU" = "#221f72", "IP Non-ICU" = "#00AEEF"))

