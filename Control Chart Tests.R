# install.packages("qcc")

library(ggplot2)
library(qcc)
library(readxl)
library(formattable)

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

dummy_df$ReceiveResultPercent <- dummy_df$ReceiveResultInTarget/dummy_df$`Volume Resulted with TAT`
dummy_df$CollectResultPercent <- dummy_df$CollectResultInTarget/dummy_df$`Volume Resulted with TAT`

dummy_df$SafePercent <- 0.90
dummy_df$NotSafePercent <- 0.80

dummy_df[ , c("ReceiveResultPercent", "CollectResultPercent", "SafePercent", "NotSafePercent")] <- lapply(dummy_df[ , c("ReceiveResultPercent", "CollectResultPercent", "SafePercent", "NotSafePercent")], percent, digits = 0)

dummy_df$ReceiveResultNP <- dummy_df$`Volume Resulted with TAT` - dummy_df$ReceiveResultInTarget
dummy_df$CollectResultNP <- dummy_df$`Volume Resulted with TAT` - dummy_df$CollectResultInTarget

# Graph option 1: Turnaround Time Compliance


qcc(data = dummy_df$ReceiveResultNP, type = "p", sizes = dummy_df$`Volume Resulted with TAT`, labels = dummy_df$Date)
qcc(data = dummy_df$CollectResultNP, type = "p", sizes = dummy_df$`Volume Resulted with TAT`, labels = dummy_df$Date)


