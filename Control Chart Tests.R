# install.packages("qcc")
install.packages("qicharts")

library(ggplot2)
library(qcc)
library(readxl)
library(formattable)
library(lubridate)

# data("orangejuice")
# 
# data <- orangejuice
# data$size2 <- round(runif(data$size, min = 45, max = 55), 0)
# 
# data <- data[data$trial == TRUE, ]
# 
# qcc(data = data$D, type = "p", sizes = data$size)
# 
# qcc(data = data$D, type = "p", sizes = data$size2)
# 
# a<-sum(data$D)
# b<-sum(data$size)
# c<-sum(data$size2)

rm(list = ls())

setwd("C:\\Users\\nevink01\\Documents\\OAO GitHub\\Lab-KPI")

dummy_df <- read_excel("Dummy Data Test.xlsx", sheet = "DummyData")
xbar_s_df <- read_excel("Dummy Data Test.xlsx", sheet = "XBar-S Data")


dummy_df$ReceiveResultPercent <- dummy_df$ReceiveResultInTarget/dummy_df$`Volume Resulted with TAT`
dummy_df$CollectResultPercent <- dummy_df$CollectResultInTarget/dummy_df$`Volume Resulted with TAT`

dummy_df$SafePercent <- 0.90
dummy_df$NotSafePercent <- 0.80

dummy_df[ , c("ReceiveResultPercent", "CollectResultPercent", "SafePercent", "NotSafePercent")] <- lapply(dummy_df[ , c("ReceiveResultPercent", "CollectResultPercent", "SafePercent", "NotSafePercent")], percent, digits = 0)

dummy_df$ReceiveResultNP <- dummy_df$`Volume Resulted with TAT` - dummy_df$ReceiveResultInTarget
dummy_df$CollectResultNP <- dummy_df$`Volume Resulted with TAT` - dummy_df$CollectResultInTarget

site1_df <- dummy_df[dummy_df$Site == "MSH", ]
site1_df$Site <- "Site1"

colnames(site1_df)[5] <- "VolumeResulted"

site1_df$Date2 <- as.Date(site1_df$Date, format = "%m/%d")

# Graph option 1: Turnaround Time Compliance --------------------------------
ggplot(data = site1_df) +
  geom_line(mapping = aes(x = Date2, y = SafePercent, color = "Safe", linetype = "Safe")) +
  geom_line(mapping = aes(x = Date2, y = NotSafePercent, color = "Not Safe", linetype = "Not Safe")) +
  geom_line(mapping = aes(x = Date2, y = ReceiveResultPercent, color = "Site Performance", linetype = "Site Performance")) +
  ylim(0.5, 1) +
  labs(title = "Site 1: Test X Receive to Result TAT within Target", x = "Date", y = "Percentage of Labs") +
  theme_bw() +
  theme(plot.title = element_text(hjust = 0.5), legend.position = "bottom") +
  scale_x_date(date_breaks = "2 days", date_labels = "%m/%d") +
  scale_color_manual(name = "", values = c("Safe" = "green", "Not Safe" = "red", "Site Performance" = "black")) +
  scale_linetype_manual(name = "", values = c("Safe" = "dashed", "Not Safe" = "dashed", "Site Performance" = "solid"))

ggplot(data = site1_df) +
  geom_line(mapping = aes(x = Date2, y = SafePercent, color = "Safe", linetype = "Safe")) +
  geom_line(mapping = aes(x = Date2, y = NotSafePercent, color = "Not Safe", linetype = "Not Safe")) +
  geom_line(mapping = aes(x = Date2, y = CollectResultPercent, color = "Site Performance", linetype = "Site Performance")) +
  ylim(0.5, 1) +
  labs(title = "Site 1: Test X Collect to Result TAT within Target", x = "Date", y = "Percentage of Labs") +
  theme_bw() +
  theme(plot.title = element_text(hjust = 0.5), legend.position = "bottom") +
  scale_x_date(date_breaks = "2 days", date_labels = "%m/%d") +
  scale_color_manual(name = "", values = c("Safe" = "green", "Not Safe" = "red", "Site Performance" = "black")) +
  scale_linetype_manual(name = "", values = c("Safe" = "dashed", "Not Safe" = "dashed", "Site Performance" = "solid"))

# Graph option 2: Average turnaround time control chart --------------------------
# Manipulate data to get in correct format
xbar_s_df$Sample <- day(xbar_s_df$Date)
receive_result <- qcc.groups(data = xbar_s_df$ReceiveResultTAT, sample = xbar_s_df$Sample)
collect_result <- qcc.groups(data = xbar_s_df$CollectResultTAT, sample = xbar_s_df$Sample)

receive_result_xbar <- qcc(receive_result, type = "xbar", labels = format(unique(xbar_s_df$Date), "%m/%d"))
plot(receive_result_xbar, title = "Site 1: Receive to Result Average TAT x_bar Chart", xlab = "Date", ylab = "Turnaround Time (min)", axes.las = 2)

receive_result_s <- qcc(receive_result, type = "S", labels = format(unique(xbar_s_df$Date), "%m/%d"))
plot(receive_result_s, title = "Site 1: Receive to Result Average TAT S Chart", xlab = "Date", ylab = "Turnaround Time SD(min)", axes.las = 2)

collect_result_xbar <- qcc(collect_result, type = "xbar", labels = format(unique(xbar_s_df$Date), "%m/%d"))
plot(collect_result_xbar, title = "Site 1: Collect to Result Average TAT x_bar Chart", xlab = "Date", ylab = "Turnaround Time (min)", axes.las = 2)

collect_result_s <- qcc(collect_result, type = "S", labels = format(unique(xbar_s_df$Date), "%m/%d"))
plot(collect_result_s, title = "Site 1: Collect to Result Average TAT S Chart", xlab = "Date", ylab = "Turnaround Time SD(min)", axes.las = 2)


# Graph option 3: Defect rate control chart -----------------------------
receive_result_np <- qcc(data = site1_df$ReceiveResultNP, type = "p", sizes = site1_df$`Volume Resulted with TAT`, labels = format(site1_df$Date, "%m/%d"))
plot(receive_result_np, title = "Site 1: Test X Receive to Result Defect Rate p-Chart", xlab = "Date", ylab = "Defect Rate", axes.las = 2)

collect_result_np <- qcc(data = site1_df$CollectResultNP, type = "p", sizes = site1_df$`Volume Resulted with TAT`, labels = format(site1_df$Date, "%m/%d"))
plot(collect_result_np, title = "Site 1: Test X Collect to Result Defect Rate p-Chart", xlab = "Date", ylab = "Defect Rate", axes.las = 2)

# Volume graph
ggplot(data = site1_df) +
  geom_line(mapping = aes(x = Date2, y = VolumeResulted, group = 1)) + 
  labs(title = "Site 1: Test X Resulted Volume", x = "Date", y = "Volume") +
  theme_bw() +
  theme(plot.title = element_text(hjust = 0.5)) + 
  scale_x_date(date_breaks = "2 days", date_labels = "%m/%d")

