# Custom function 
preprocess_scc_sun <- function(raw_scc, raw_sun)  {
  
  # SCC DATA PROCESSING --------------------------
  # SCC lookup references ----------------------------------------------
  # Crosswalk labs included and remove out of scope labs
  raw_scc <- left_join(raw_scc, 
                       test_code[ , c("Test", "SCC_TestID", "Division")], 
                       by = c("TEST_ID" = "SCC_TestID"))
  raw_scc$TEST_ID <- as.factor(raw_scc$TEST_ID)
  raw_scc$Division <- as.factor(raw_scc$Division)
  
  raw_scc <- raw_scc %>%
    mutate(TestIncl = ifelse(is.na(raw_scc$Test), FALSE, TRUE))
  
  raw_scc <- raw_scc %>%
    filter(TestIncl)
  
  # Crosswalk units and identify ICUs
  raw_scc <- raw_scc %>%
    mutate(WardandName = paste(Ward, WARD_NAME))
  
  raw_scc <- left_join(raw_scc, scc_icu[ , c("Concatenate", "ICU")], 
                       by = c("WardandName" = "Concatenate"))
  raw_scc[is.na(raw_scc$ICU), "ICU"] <- FALSE
  
  # Crosswalk unit type
  raw_scc <- left_join(raw_scc, scc_setting, 
                       by = c("CLINIC_TYPE" = "Clinic_Type"))
  # Crosswalk site name
  raw_scc <- left_join(raw_scc, mshs_site, 
                       by = c("SITE" = "DataSite"))
  
  # SCC data formatting ----------------------------------------------
  raw_scc[c("Ward", 
            "WARD_NAME", 
            "REQUESTING_DOC",
            "GROUP_TEST_ID",
            "TEST_ID",
            "TEST_NAME",
            "Test",
            "COLLECT_CENTER_ID",
            "SITE",
            "Site",
            "CLINIC_TYPE",
            "Setting",
            "SettingRollUp")] <- lapply(raw_scc[c("Ward",
                                                  "WARD_NAME",
                                                  "REQUESTING_DOC", 
                                                  "GROUP_TEST_ID",
                                                  "TEST_ID", 
                                                  "TEST_NAME", 
                                                  "Test",
                                                  "COLLECT_CENTER_ID",
                                                  "SITE",
                                                  "Site",
                                                  "CLINIC_TYPE",
                                                  "Setting",
                                                  "SettingRollUp")], 
                                        as.factor)
  
  # Fix any timestamps that weren't imported correctly and then format as date/time
  raw_scc[c("ORDERING_DATE", 
            "COLLECTION_DATE", 
            "RECEIVE_DATE", 
            "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE", 
                                                  "COLLECTION_DATE", 
                                                  "RECEIVE_DATE", 
                                                  "VERIFIED_DATE")], 
                                        function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, 
                                                           str_replace(x, "\\*.*\\*", ""), x))
  
  raw_scc[c("ORDERING_DATE",
            "COLLECTION_DATE",
            "RECEIVE_DATE",
            "VERIFIED_DATE")] <- lapply(raw_scc[c("ORDERING_DATE",
                                                  "COLLECTION_DATE",
                                                  "RECEIVE_DATE",
                                                  "VERIFIED_DATE")], 
                                        as.POSIXct, tz = "UTC", format = "%Y-%m-%d %H:%M:%OS", 
                                        options(digits.sec = 1))
  
  # Add a column for Resulted date for later use in repository
  raw_scc <- raw_scc %>%
    mutate(ResultedDate = as.Date(VERIFIED_DATE, format = "%m/%d/%Y"))
  
  # Update patient setting to reflect ICU/Non-ICU and 
  # update priorities for ED and ICU labs
  raw_scc <- raw_scc %>%
    mutate(MasterSetting = ifelse(CLINIC_TYPE == "E", "ED", 
                                  ifelse(CLINIC_TYPE == "O", "Amb", 
                                         ifelse(CLINIC_TYPE == "I" & 
                                                  ICU == TRUE, "ICU",
                                                ifelse(CLINIC_TYPE == "I" & 
                                                         ICU != TRUE, 
                                                       "IP Non-ICU", 
                                                       "Other")))),
           DashboardSetting = ifelse(MasterSetting == "ED" | 
                                       MasterSetting == "ICU", "ED & ICU", 
                                     MasterSetting))
  
  # Update priority to reflect ED/ICU as stat and create Master Priority 
  # for labs where all specimens are treated as stat
  raw_scc <- raw_scc %>%
    mutate(AdjPriority = ifelse(MasterSetting == "ED" | 
                                  MasterSetting == "ICU" | PRIORITY == "S", 
                                "Stat", "Routine"),
           DashboardPriority = ifelse(
             tat_targets$Priority[match(Test, tat_targets$Test)] == "All", 
             "All", AdjPriority))
  
  # Calculate turnaround times
  raw_scc <- raw_scc %>%
    mutate(CollectToReceive = RECEIVE_DATE - COLLECTION_DATE,
           ReceiveToResult = VERIFIED_DATE - RECEIVE_DATE,
           CollectToResult = VERIFIED_DATE - COLLECTION_DATE)
  
  raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- 
    lapply(raw_scc[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], 
           as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min 
  # after specimen received
  # Identify specimens with missing collections times as those with 
  # collection time defaulted to receive time
  raw_scc <- raw_scc %>%
    mutate(AddOnMaster = ifelse(difftime(ORDERING_DATE, 
                                         RECEIVE_DATE, units = "mins") > 5, 
                                "AddOn", "Original"),
           MissingCollect = ifelse(CollectToReceive == 0, TRUE, FALSE))
  
  # Determine target TAT based on test, priority, and patient setting
  raw_scc <- raw_scc %>%
    mutate(Concate1 = paste(Test, DashboardPriority),
           Concate2 = paste(Test, DashboardPriority, MasterSetting),
           ReceiveResultTarget = ifelse(
             !is.na(match(Concate2, tat_targets$Concate)),
             tat_targets$ReceiveToResultTarget[match(Concate2, 
                                                     tat_targets$Concate)],
             ifelse(!is.na(match(Concate1, tat_targets$Concate)), 
                    tat_targets$ReceiveToResultTarget[match(Concate1, 
                                                            tat_targets$Concate)],
                    tat_targets$ReceiveToResultTarget[match(Test, 
                                                            tat_targets$Concate)])),
           CollectResultTarget = ifelse(
             !is.na(match(Concate2, tat_targets$Concate)),
             tat_targets$CollectToResultTarget[match(Concate2, 
                                                     tat_targets$Concate)], 
             ifelse(!is.na(match(Concate1, tat_targets$Concate)), 
                    tat_targets$CollectToResultTarget[match(Concate1, 
                                                            tat_targets$Concate)], 
                    tat_targets$CollectToResultTarget[match(Test, 
                                                            tat_targets$Concate)])),
           ReceiveResultInTarget = ReceiveToResult <= ReceiveResultTarget, 
           CollectResultInTarget = CollectToResult <= CollectResultTarget)
  
  # Identify and remove duplicate tests
  raw_scc <- raw_scc %>%
    mutate(Concate3 = paste(LAST_NAME, FIRST_NAME,
                            ORDER_ID, TEST_NAME, 
                            COLLECTION_DATE, RECEIVE_DATE, VERIFIED_DATE))
  
  raw_scc <- raw_scc[!duplicated(raw_scc$Concate3), ]
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or 
  # receive times after result, or orders with missing collect, receive, or 
  # result timestamps
  raw_scc <- raw_scc %>% 
    mutate(TATInclude = ifelse(AddOnMaster != "Original" | 
                                 MasterSetting == "Other" | 
                                 CollectToResult < 0 | 
                                 ReceiveToResult < 0 | 
                                 is.na(CollectToResult) | 
                                 is.na(ReceiveToResult), FALSE, TRUE))
  
  scc_master <- raw_scc[ , c("Ward", "WARD_NAME", "WardandName",
                             "ORDER_ID", "REQUESTING_DOC NAME", 
                             "MPI", "WORK SHIFT",
                             "TEST_NAME", "Test", "Division", "PRIORITY", 
                             "Site", "ICU", "CLINIC_TYPE", 
                             "Setting", "SettingRollUp", 
                             "MasterSetting", "DashboardSetting",
                             "AdjPriority", "DashboardPriority",
                             "ORDERING_DATE", "COLLECTION_DATE", 
                             "RECEIVE_DATE", "VERIFIED_DATE",
                             "ResultedDate", 
                             "CollectToReceive", "ReceiveToResult", 
                             "CollectToResult", 
                             "AddOnMaster", "MissingCollect", 
                             "ReceiveResultTarget", "CollectResultTarget", 
                             "ReceiveResultInTarget", "CollectResultInTarget", 
                             "TATInclude")]
  
  colnames(scc_master) <- c("LocCode", "LocName", "LocConcat",
                            "OrderID", "RequestMD", 
                            "MSMRN", "WorkShift", 
                            "TestName", "Test", "Division", "OrderPriority", 
                            "Site", "ICU", "LocType", 
                            "Setting", "SettingRollUp", 
                            "MasterSetting", "DashboardSetting",
                            "AdjPriority", "DashboardPriority",
                            "OrderTime", "CollectTime", 
                            "ReceiveTime", "ResultTime", 
                            "ResultDate", 
                            "CollectToReceiveTAT", "ReceiveToResultTAT", 
                            "CollectToResultTAT", 
                            "AddOnMaster", "MissingCollect", 
                            "ReceiveResultTarget", "CollectResultTarget", 
                            "ReceiveResultInTarget", "CollectResultInTarget", 
                            "TATInclude")
  
  
  ## SUNQUEST DATA PROCESSING ----------
  # Sunquest lookup references ----------------------------------------------
  # Crosswalk labs included and remove out of scope labs
  raw_sun <- left_join(raw_sun, test_code[ , c("Test", 
                                               "SUN_TestCode", 
                                               "Division")], 
                       by = c("TestCode" = "SUN_TestCode"))
  raw_sun$TestCode <- as.factor(raw_sun$TestCode)
  raw_sun$Division <- as.factor(raw_sun$Division)
  raw_sun <- raw_sun %>%
    mutate(TestIncl = ifelse(is.na(Test), FALSE, TRUE))
  
  raw_sun <- raw_sun[raw_sun$TestIncl == TRUE, ]
  
  # Crosswalk units and identify ICUs
  raw_sun <- raw_sun %>%
    mutate(LocandName = paste(LocCode, LocName))
  
  raw_sun <- left_join(raw_sun, sun_icu[ , c("Concatenate", "ICU")], 
                       by = c("LocandName" = "Concatenate"))
  raw_sun[is.na(raw_sun$ICU), "ICU"] <- FALSE
  
  # Crosswalk unit type
  raw_sun <- left_join(raw_sun, sun_setting, by = c("LocType" = "LocType"))
  
  # Crosswalk site name
  #raw_sun <- left_join(raw_sun, mshs_site, by = c("HospCode" = "DataSite"))
  #Asala Added this to tackle the issue of the other site temporarly until Kate comes back
  raw_sun <- inner_join(raw_sun, mshs_site, by = c("HospCode" = "DataSite"))
  
  # Sunquest data formatting --------------------------------------------
  raw_sun[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
            "LocType", "LocCode", "LocName", 
            "PhysName", "SHIFT",
            "ReceiveTech", "ResultTech", "PerformingLabCode",
            "Test", "LocandName", "Setting", "SettingRollUp", "Site")] <- 
    lapply(raw_sun[c("HospCode", "BatTstCode", "BATName", "TestCode", "TSTName",
                     "LocType", "LocCode", "LocName", 
                     "PhysName", "SHIFT", 
                     "ReceiveTech", "ResultTech", "PerformingLabCode", 
                     "Test", "LocandName", "Setting", "SettingRollUp", "Site")],
           as.factor)
  
  # Fix any timestamps that weren't imported correctly and then format as date/time
  raw_sun[c("OrderDateTime", 
            "CollectDateTime", 
            "ReceiveDateTime", 
            "ResultDateTime")] <- 
    lapply(raw_sun[c("OrderDateTime", 
                     "CollectDateTime", 
                     "ReceiveDateTime", 
                     "ResultDateTime")], 
           function(x) ifelse(!is.na(x) & str_detect(x, "\\*.*\\*")  == TRUE, 
                              str_replace(x, "\\*.*\\*", ""), x))
  
  raw_sun[c("OrderDateTime", 
            "CollectDateTime", 
            "ReceiveDateTime", 
            "ResultDateTime")] <- 
    lapply(raw_sun[c("OrderDateTime", 
                     "CollectDateTime", 
                     "ReceiveDateTime", 
                     "ResultDateTime")], 
           as.POSIXct, tz = "UTC", format = "%m/%d/%Y %H:%M:%S")
  
  # Add a column for Resulted date for later use in repository
  raw_sun <- raw_sun %>%
    mutate(ResultedDate = as.Date(ResultDateTime, format = "%m/%d/%Y"))
  
  # Sunquest data preprocessing -----------------------------------------------
  # Update patient setting to reflect ICU/Non-ICU and update priorities for 
  # ED and ICU labs
  raw_sun <- raw_sun %>%
    mutate(MasterSetting = ifelse(SettingRollUp == "ED", "ED", 
                                  ifelse(SettingRollUp == "Amb", "Amb", 
                                         ifelse(SettingRollUp == "IP" & 
                                                  ICU == TRUE, "ICU",
                                                ifelse(SettingRollUp == "IP" & 
                                                         ICU == FALSE, 
                                                       "IP Non-ICU", "Other")))), 
           DashboardSetting = ifelse(MasterSetting == "ED" | 
                                       MasterSetting == "ICU", "ED & ICU", 
                                     MasterSetting))
  
  # Update priority to reflect ED/ICU as stat and create Master Priority for 
  # labs where all specimens are treated as stat
  raw_sun <- raw_sun %>%
    mutate(AdjPriority = ifelse(MasterSetting != "ED" & MasterSetting != "ICU" &
                                  is.na(SpecimenPriority), "Routine",
                                ifelse(MasterSetting == "ED" | 
                                         MasterSetting == "ICU" | 
                                         SpecimenPriority == "S", "Stat", 
                                       "Routine")),
           DashboardPriority = ifelse(tat_targets$Priority[match(
             Test, tat_targets$Test)] == "All", "All", AdjPriority))
  
  # Calculate turnaround times
  raw_sun <- raw_sun %>%
    mutate(CollectToReceive = ReceiveDateTime - CollectDateTime,
           ReceiveToResult = ResultDateTime - ReceiveDateTime,
           CollectToResult = ResultDateTime - CollectDateTime)
  
  raw_sun[c("CollectToReceive", "ReceiveToResult", "CollectToResult")] <- 
    lapply(raw_sun[c("CollectToReceive", "ReceiveToResult", "CollectToResult")], 
           as.numeric, units = "mins")
  
  # Identify add on orders as orders placed more than 5 min after 
  # specimen received. 
  # Identify specimens with missing collections times as those with collection 
  # time defaulted to order time
  raw_sun <- raw_sun %>%
    mutate(AddOnMaster = ifelse(difftime(OrderDateTime, ReceiveDateTime, 
                                         units = "mins") > 5, "AddOn", 
                                "Original"),
           MissingCollect = ifelse(CollectDateTime == OrderDateTime, TRUE, 
                                   FALSE))
  
  # Determine target TAT based on test, priority, and patient setting
  raw_sun <- raw_sun %>%
    mutate(Concate1 = paste(Test, DashboardPriority), 
           Concate2 = paste(Test, DashboardPriority, MasterSetting),
           ReceiveResultTarget = ifelse(!is.na(match(Concate2, 
                                                     tat_targets$Concate)), 
                                        tat_targets$ReceiveToResultTarget[
                                          match(Concate2, tat_targets$Concate)], 
                                        ifelse(!is.na(match(Concate1,
                                                            tat_targets$Concate)),
                                               tat_targets$ReceiveToResultTarget[
                                                 match(Concate1, 
                                                       tat_targets$Concate)],
                                               tat_targets$ReceiveToResultTarget[
                                                 match(Test, 
                                                       tat_targets$Concate)])),
           CollectResultTarget = ifelse(!is.na(match(Concate2, 
                                                     tat_targets$Concate)), 
                                        tat_targets$CollectToResultTarget[
                                          match(Concate2, tat_targets$Concate)], 
                                        ifelse(!is.na(match(Concate1, 
                                                            tat_targets$Concate)), 
                                               tat_targets$CollectToResultTarget[
                                                 match(Concate1, 
                                                       tat_targets$Concate)], 
                                               tat_targets$CollectToResultTarget[
                                                 match(Test, 
                                                       tat_targets$Concate)])),
           ReceiveResultInTarget = ReceiveToResult <= ReceiveResultTarget,
           CollectResultInTarget = CollectToResult <= CollectResultTarget)
  
  # Identify and remove duplicate tests
  raw_sun <- raw_sun %>%
    mutate(Concate3 = paste(PtNumber, HISOrderNumber, TSTName, 
                            CollectDateTime, ReceiveDateTime, ResultDateTime))
  
  raw_sun <- raw_sun[!duplicated(raw_sun$Concate3), ]
  
  # Identify which labs to include in TAT analysis
  # Exclude add on orders, orders from "other" settings, orders with collect or 
  # receive times after result, or orders with missing collect, receive, or 
  # result timestamps
  raw_sun <- raw_sun %>%
    mutate(TATInclude = ifelse(AddOnMaster != "Original" | 
                                 MasterSetting == "Other" | 
                                 CollectToResult < 0 | 
                                 ReceiveToResult < 0 | 
                                 is.na(CollectToResult) | 
                                 is.na(ReceiveToResult), FALSE, TRUE))
  
  sun_master <- raw_sun[ ,c("LocCode", "LocName", "LocandName", 
                            "HISOrderNumber", "PhysName", 
                            "PtNumber", "SHIFT",            
                            "TSTName", "Test", "Division", "SpecimenPriority", 
                            "Site", "ICU", "LocType",               
                            "Setting", "SettingRollUp", 
                            "MasterSetting", "DashboardSetting", 
                            "AdjPriority", "DashboardPriority", 
                            "OrderDateTime", "CollectDateTime", 
                            "ReceiveDateTime", "ResultDateTime", 
                            "ResultedDate", 
                            "CollecttoReceive", "ReceivetoResult", 
                            "CollecttoResult", 
                            "AddOnMaster", "MissingCollect", 
                            "ReceiveResultTarget", "CollectResultTarget", 
                            "ReceiveResultInTarget", "CollectResultInTarget", 
                            "TATInclude")]
  
  colnames(sun_master) <- c("LocCode", "LocName", "LocConcat", 
                            "OrderID", "RequestMD", 
                            "MSMRN", "WorkShift", 
                            "TestName", "Test", "Division", "OrderPriority", 
                            "Site", "ICU", "LocType", 
                            "Setting", "SettingRollUp", 
                            "MasterSetting", "DashboardSetting", 
                            "AdjPriority", "DashboardPriority",
                            "OrderTime", "CollectTime", 
                            "ReceiveTime", "ResultTime", 
                            "ResultDate", 
                            "CollectToReceiveTAT", "ReceiveToResultTAT", 
                            "CollectToResultTAT", 
                            "AddOnMaster", "MissingCollect", 
                            "ReceiveResultTarget", "CollectResultTarget", 
                            "ReceiveResultInTarget", "CollectResultInTarget", 
                            "TATInclude")
  
  
  scc_sun_master <- rbind(scc_master, sun_master)
  
  scc_sun_master[c("LocConcat", "RequestMD", "WorkShift", 
                   "AdjPriority", "AddOnMaster")] <-
    lapply(scc_sun_master[c("LocConcat", "RequestMD", "WorkShift",
                            "AdjPriority", "AddOnMaster")], as.factor)
  
  scc_sun_master$Site <- factor(scc_sun_master$Site, 
                                levels = site_order)
  scc_sun_master$Test <- factor(scc_sun_master$Test, 
                                levels = cp_micro_lab_order)
  scc_sun_master$MasterSetting <- factor(scc_sun_master$MasterSetting, 
                                         levels = pt_setting_order)
  scc_sun_master$DashboardSetting <- factor(scc_sun_master$DashboardSetting, 
                                            levels = pt_setting_order2)
  scc_sun_master$DashboardPriority <- factor(scc_sun_master$DashboardPriority, 
                                             levels = dashboard_priority_order)
  
  scc_sun_list <- list(raw_scc, raw_sun, scc_sun_master)
  
}