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

mshs_colors = c("#221f72", "#00AEFF", "#B2B3B2", "#D80B8C")

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
  scale_color_manual(name = "Setting", values = mshs_colors)

# Create a custom function for graphing percentage of labs within TAT target
tat_percentage_graph <- function(data, site, metric) {
  metric_name <- ifelse(metric == "ReceiveResultPercent", "Receive to Result TAT", ifelse(metric == "CollectResultPercent", "Collect to Result TAT", "Error"))
  ggplot() +
    geom_rect(aes(xmin = start_date - 1, xmax = max(data$ResultDate) + 1, ymin = -Inf, ymax = min(data$NotSafeThreshold), fill = "Not Safe", alpha = "Not Safe")) +
    geom_rect(aes(xmin = start_date - 1, xmax = max(data$ResultDate) + 1, ymin = min(data$NotSafeThreshold), ymax = min(data$SafeThreshold), fill = "At Risk", alpha = "At Risk")) +
    geom_rect(aes(xmin = start_date - 1, xmax = max(data$ResultDate) + 1, ymin = min(data$SafeThreshold), ymax = Inf, fill = "Safe", alpha = "Safe")) +
    geom_point(data = data[data$ResultDate >= start_date & data$Site == site, ], aes_string(x = "ResultDate", y = metric, color = "DashboardSetting", shape = "DashboardSetting"), size = 3) +
    geom_line(data = data[data$ResultDate >= start_date & data$Site == site, ], aes_string(x = "ResultDate", y = metric, color = "DashboardSetting"), linetype = "dashed") +
    labs(title = paste0("Troponin ", metric_name, ":\n", site, " Percentage of Labs within Target"), x = "Date",  y = "Percentage of Labs", color = "Setting", shape = "Setting") +
    theme_bw() +
    theme(plot.title = element_text(hjust = 0.5), legend.position = "right", legend.box = "vertical", legend.title = element_text(size = 10), legend.text = element_text(size = 8),axis.text.x = element_text(angle = 30, hjust = 1)) +
    scale_x_date(limits = c(start_date -1, max(data$ResultDate)+1), breaks = seq(start_date, max(data$ResultDate), by = 3), date_minor_breaks = "1 day", date_labels = "%m/%d/%y", expand = c(0, 0, 0, 0)) +
    scale_fill_manual(name = "Status", values = c("Not Safe" = "red", "At Risk" = "orange", "Safe" = "green"), breaks=c("Not Safe","At Risk","Safe")) + 
    scale_alpha_manual(name = "Status", values = c("Not Safe" = 0.2, "At Risk" = 0.2, "Safe" = 0.2), breaks=c("Not Safe","At Risk","Safe")) +
    scale_color_manual(name = "Setting", values = mshs_colors)
}

tat_percentage_graph(troponin_test, "MSH", metric = "ReceiveResultPercent")

test2 <- troponin_test[troponin_test$ResultDate >= start_date & troponin_test$Site == "MSH", ]

dashboard_tat_summary$Concate <- paste(dashboard_tat_summary$Test, dashboard_tat_summary$DashboardPriority, dashboard_tat_summary$DashboardSetting)

# Manually remove patient settings or combinations not represented in efficiency indicator dashboards
dashboard_tat_filter <- dashboard_tat_summary
dashboard_tat_filter <- dashboard_tat_filter[!(dashboard_tat_filter$DashboardPriority == "Other" | dashboard_tat_filter$DashboardSetting == "Other" |
                                                             ((dashboard_tat_filter$Test == "Troponin" | dashboard_tat_filter$Test == "Lactate WB") & dashboard_tat_filter$DashboardSetting == "Amb") |
                                                             (dashboard_tat_filter$Test == "C. diff" & dashboard_tat_filter$DashboardSetting == "Amb")), ]

cp_micro_lab_order <- c("Troponin", "Lactate WB", "BUN", "HGB", "PT", "Rapid Flu", "C. diff")
site_order <- c("MSH", "MSQ", "MSBI", "MSB", "MSW", "MSSL")
dashboard_pt_setting <- c("ED & ICU", "IP Non-ICU", "Amb")
dashboard_priority_order <- c("All", "Stat", "Routine")

dashboard_tat_filter$Site <- factor(dashboard_tat_filter$Site, levels = site_order)
dashboard_tat_filter$Test <- factor(dashboard_tat_filter$Test, levels = cp_micro_lab_order)
dashboard_tat_filter$DashboardPriority <- factor(dashboard_tat_filter$DashboardPriority, levels = dashboard_priority_order)
dashboard_tat_filter$DashboardSetting <- factor(dashboard_tat_filter$DashboardSetting, levels = dashboard_pt_setting)
dashboard_tat_filter <- dashboard_tat_filter[order(dashboard_tat_filter$Test, dashboard_tat_filter$Site, dashboard_tat_filter$DashboardPriority, dashboard_tat_filter$DashboardSetting, dashboard_tat_filter$ResultDate), ]

tests <- unique(dashboard_tat_filter$Test)
for (test in tests) {
  print(paste("Test:", test))
  test_df <- dashboard_tat_filter[dashboard_tat_filter$Test == test, ]
  sites <- unique(test_df$Site)
  for (site in sites) {
    print(paste("Site: ", site))
    site_test_df <- test_df[test_df$Site == site, ] 
    priorities <- unique(site_test_df$DashboardPriority)
    for (priority in priorities) {
      print(paste("Priority: ", priority))
      priority_site_test_df <- site_test_df[site_test_df$DashboardPriority == priority, ]
      patient_settings <- unique(priority_site_test_df$DashboardSetting)
      for (patient_setting in patient_settings) {
        print(paste("Patient Setting: ", patient_setting))
      }
    }
  }
}


tests <- unique(dashboard_tat_filter$Test)
for (test in tests) {
  print(paste("Test:", test))
  test_df <- dashboard_tat_filter[dashboard_tat_filter$Test == test, ]
  sites <- unique(test_df$Site)
  for (site in sites) {
    print(paste("Site: ", site))
    site_test_df <- test_df[test_df$Site == site, ] 
    priorities <- unique(site_test_df$DashboardPriority)
    for (priority in priorities) {
      print(paste("Priority: ", priority))
      priority_site_test_df <- site_test_df[site_test_df$DashboardPriority == priority, ]
      patient_settings <- unique(priority_site_test_df$DashboardSetting)
      for (patient_setting in patient_settings) {
        print(paste("Patient Setting: ", patient_setting))
      }
    }
  }
}

tat_percentage_graph2 <- function(data, test_name, site_name, lab_priority, metric) {
  metric_name <- ifelse(metric == "ReceiveResultPercent", "Receive to Result TAT", ifelse(metric == "CollectResultPercent", "Collect to Result TAT", "Error"))
  # df <- data[data$ResultDate >= start_date & data$Test == test & data$Site == site & data$DashboardPriority == lab_priority, ]
  # 
  ggplot() +
    geom_rect(aes(xmin = start_date - 1, xmax = max(data$ResultDate) + 1, ymin = -Inf, ymax = min(data$NotSafeThreshold), fill = "Not Safe", alpha = "Not Safe")) +
    geom_rect(aes(xmin = start_date - 1, xmax = max(data$ResultDate) + 1, ymin = min(data$NotSafeThreshold), ymax = min(data$SafeThreshold), fill = "At Risk", alpha = "At Risk")) +
    geom_rect(aes(xmin = start_date - 1, xmax = max(data$ResultDate) + 1, ymin = min(data$SafeThreshold), ymax = Inf, fill = "Safe", alpha = "Safe")) +
    geom_point(data = data[data$ResultDate >= start_date & data$Test == test_name & data$Site == site_name & data$DashboardPriority == lab_priority, ], aes_string(x = "ResultDate", y = metric, color = "DashboardSetting", shape = "DashboardSetting"), size = 3) +
    geom_line(data = data[data$ResultDate >= start_date & data$Test == test_name & data$Site == site_name & data$DashboardPriority == lab_priority, ], aes_string(x = "ResultDate", y = metric, color = "DashboardSetting"), linetype = "dashed") +
    labs(title = paste0(test_name, " ", metric_name, ":\n", site_name, " Percentage of ", lab_priority, " Labs within Target"), x = "Date",  y = "Percentage of Labs", color = "Setting", shape = "Setting") +
    theme_bw() +
    theme(plot.title = element_text(hjust = 0.5), legend.position = "right", legend.box = "vertical", legend.title = element_text(size = 10), legend.text = element_text(size = 8),axis.text.x = element_text(angle = 30, hjust = 1)) +
    scale_x_date(limits = c(start_date -1, max(data$ResultDate)+1), breaks = seq(start_date, max(data$ResultDate), by = 3), date_minor_breaks = "1 day", date_labels = "%m/%d/%y", expand = c(0, 0, 0, 0)) +
    scale_fill_manual(name = "Status", values = c("Not Safe" = "red", "At Risk" = "orange", "Safe" = "green"), breaks=c("Not Safe","At Risk","Safe")) +
    scale_alpha_manual(name = "Status", values = c("Not Safe" = 0.2, "At Risk" = 0.2, "Safe" = 0.2), breaks=c("Not Safe","At Risk","Safe")) +
    scale_color_manual(name = "Setting", values = mshs_colors)
}

test_df <- dashboard_tat_filter[dashboard_tat_filter$Test == "Troponin" | dashboard_tat_filter$Test == "Lactate WB", ]

tat_percentage_graph2(data = test_df, test_name = "Troponin", site_name = "MSH", lab_priority = "All", metric = "ReceiveResultPercent")

# Create a loop to graph all necessary combinations
for (test in unique(test_df$Test)) {
  # print(paste("Test:", test))
  # test_df <- dashboard_tat_filter[dashboard_tat_filter$Test == test, ]
  dummy_df <- test_df[test_df$Test == test, ]
  sites <- unique(test_df$Site)
  for (site in sites) {
    # print(paste("Site: ", site))
    # site_test_df <- test_df[test_df$Site == site, ] 
    dummy_df <- dummy_df[dummy_df$Site == site, ]
    priorities <- unique(dummy_df$DashboardPriority)
    for (priority in priorities) {
      dummy_df <- dummy_df[dummy_df$DashboardPriority == priority, ]
      tat_percentage_graph2(data = dummy_df, test_name = test, site_name = site, lab_priority = priority, metric = "ReceiveResultPercent")
      # print(paste("Priority: ", priority))
      # priority_site_test_df <- site_test_df[site_test_df$DashboardPriority == priority, ]
      # patient_settings <- unique(priority_site_test_df$DashboardSetting)
      # for (patient_setting in patient_settings) {
      #   print(paste("Patient Setting: ", patient_setting))
      # }
    }
  }
}

tat_percentage_graph(data = dummy_df, site = "MSH", metric = "ReceiveResultPercent")

