# Code for preprocessing, analyzing, and displaying Clinical Pathology KPI -----
# CP includes Chemistry, Hematology, and Microbiology RRL divisions

# Custom function for preprocessing raw data from SCC and Sunquest -------
preprocess_cp <- function(raw_scc, raw_sun)  {

  # Preprocess SCC data -------------------------------
  # Remove any duplicates
  raw_scc <- unique(raw_scc)
  # Correct and format any timestamps that were not imported correctly
  raw_scc[c("ORDERING_DATE",
            "COLLECTION_DATE",
            "RECEIVE_DATE",
            "VERIFIED_DATE")] <-
    lapply(raw_scc[c("ORDERING_DATE",
                     "COLLECTION_DATE",
                     "RECEIVE_DATE",
                     "VERIFIED_DATE")],
           function(x)
             ifelse(!is.na(x) & str_detect(x, "\\*.*\\*"),
                    str_replace(x, "\\*.*\\*", ""), x))

  raw_scc[c("ORDERING_DATE",
            "COLLECTION_DATE",
            "RECEIVE_DATE",
            "VERIFIED_DATE")] <-
    lapply(raw_scc[c("ORDERING_DATE",
                     "COLLECTION_DATE",
                     "RECEIVE_DATE",
                     "VERIFIED_DATE")],
           as.POSIXct, tz = "UTC",
           format = "%Y-%m-%d %H:%M:%OS",
           options(digits.sec = 1))

  # SCC lookup references ----------------------------------------------
  # Crosswalk in scope labs
  raw_scc <- left_join(raw_scc,
                       scc_test_code,
                       by = c("TEST_ID" = "SCC_TestID"))

  # Determine if test is included based on crosswalk results
  raw_scc <- raw_scc %>%
    mutate(TestIncl = !is.na(Test)) %>%
    filter(TestIncl)

  # Crosswalk unit type
  raw_scc <- left_join(raw_scc, scc_setting,
                       by = c("CLINIC_TYPE" = "Clinic_Type"))
  # Crosswalk site name
  raw_scc <- left_join(raw_scc, mshs_site,
                       by = c("SITE" = "DataSite"))

  # Crosswalk units and identify ICUs
  raw_scc <- raw_scc %>%
    mutate(WardandName = paste(Ward, WARD_NAME))

  raw_scc <- left_join(raw_scc, scc_icu[, c("Concatenate", "ICU")],
                       by = c("WardandName" = "Concatenate"))

  # Preprocess SCC data and add any necessary columns
  raw_scc <- raw_scc %>%
    mutate(
      # Determine if unit is an ICU based on crosswalk results
      ICU = ifelse(is.na(ICU), FALSE, ICU),
      # Create a column for resulted date
      ResultedDate = date(VERIFIED_DATE),
      # Create master setting column to identify ICU and IP Non-ICU units
      MasterSetting = ifelse(SettingRollUp == "IP" & ICU, "ICU",
                             ifelse(SettingRollUp == "IP" & !ICU,
                                    "IP Non-ICU", SettingRollUp)),
      # Create dashboard setting column to roll up master settings based on
      # desired dashboard grouping (ie, group ED and ICU together)
      DashboardSetting = ifelse(MasterSetting %in% c("ED", "ICU"),
                                "ED & ICU", MasterSetting),
      # Create column with adjusted priority based on assumption that all ED and
      # ICU labs are treated as stat per operational leadership
      AdjPriority = ifelse(MasterSetting %in% c("ED", "ICU") |
                             PRIORITY %in% "S", "Stat", "Routine"),
      # Create dashboard priority column
      DashboardPriority = ifelse(
        tat_targets$Priority[match(Test, tat_targets$Test)] == "All",
        "All", AdjPriority),
      # Calculate turnaround times
      CollectToReceive =
        as.numeric(RECEIVE_DATE - COLLECTION_DATE, units = "mins"),
      ReceiveToResult =
        as.numeric(VERIFIED_DATE - RECEIVE_DATE, units = "mins"),
      CollectToResult =
        as.numeric(VERIFIED_DATE - COLLECTION_DATE, units = "mins"),
      #
      # Determine if order was an add on or original order based on time between
      # order and receive times
      AddOnMaster = ifelse(as.numeric(ORDERING_DATE - RECEIVE_DATE,
                                      units = "mins")
                           > 5, "AddOn", "Original"),
      # Determine if collection time is missing
      MissingCollect = CollectToReceive == 0,
      #
      # Determine TAT based on test, priority, and patient setting
      # Create column concatenating test and priority to determine TAT targets
      Concate1 = paste(Test, DashboardPriority),
      # Create column concatenating test, priority, and setting to determine
      # TAT targets
      Concate2 = paste(Test, DashboardPriority, MasterSetting),
      # Determine Receive to Result TAT target using this logic:
      # 1. Try to match test, priority, and setting (applicable for labs with
      # different TAT targets based on patient setting and order priority)
      # 2. Try to match test and priority (applicable for labs with different
      # TAT targets based on order priority)
      # 3. Try to match test - this is for tests with (applicable for labs with
      # TAT targets that are independent of patient setting or priority)
      #
      # Determine Receive to Result TAT target based on above logic/scenarios
      ReceiveResultTarget =
        # Match on scenario 1
        ifelse(!is.na(match(Concate2, tat_targets$Concate)),
               tat_targets$ReceiveToResultTarget[
                 match(Concate2, tat_targets$Concate)],
               # Match on scenario 2
               ifelse(!is.na(match(Concate1, tat_targets$Concate)),
                      tat_targets$ReceiveToResultTarget[
                        match(Concate1, tat_targets$Concate)],
                      # Match on scenario 3
                      tat_targets$ReceiveToResultTarget[
                        match(Test, tat_targets$Concate)])),
      #
      # Determine Collect to Result TAT target based on above logic/scenarios
      CollectResultTarget =
        # Match on scenario 1
        ifelse(!is.na(match(Concate2, tat_targets$Concate)),
               tat_targets$CollectToResultTarget[
                 match(Concate2, tat_targets$Concate)],
               # Match on scenario 2
               ifelse(!is.na(match(Concate1, tat_targets$Concate)),
                      tat_targets$CollectToResultTarget[
                        match(Concate1, tat_targets$Concate)],
                      # Match on scenario 3
                      tat_targets$CollectToResultTarget[
                        match(Test, tat_targets$Concate)])),
      #
      # Determine if Receive to Result and Collect to Result TAT meet targets
      ReceiveResultInTarget = ReceiveToResult <= ReceiveResultTarget,
      CollectResultInTarget = CollectToResult <= CollectResultTarget,
      # Create column with patient name, order ID, test, collect, receive, and
      # result date and determine if there is a duplicate; order time excluded
      Concate3 = paste(LAST_NAME, FIRST_NAME,
                       ORDER_ID, TEST_NAME,
                       COLLECTION_DATE, RECEIVE_DATE, VERIFIED_DATE),
      DuplTest = duplicated(Concate3),
      # Determine whether or not to include this particular lab in TAT analysis
      # Exclusion criteria:
      # 1. Add on orders
      # 2. Orders from "Other" settings
      # 3. Orders with collect or receive times after result time
      # 4. Orders with missing collect, receive, or result timestamps
      TATInclude = ifelse(AddOnMaster == "AddOn" |
                            MasterSetting == "Other" |
                            CollectToReceive < 0 |
                            CollectToResult < 0 |
                            ReceiveToResult < 0 |
                            is.na(CollectToResult) |
                            is.na(ReceiveToResult), FALSE, TRUE))

  # Remove duplicate tests
  raw_scc <- raw_scc %>%
    filter(!DuplTest)

  # Select columns
  scc_master <- raw_scc[, c("Ward", "WARD_NAME", "WardandName",
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
  # Rename columns
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

  # Preprocess Sunquest data --------------------------------
  # Remove any duplicates
  raw_sun <- unique(raw_sun)
  # Correct and format any timestamps that were not imported correctly
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

  # Sunquest lookup references
  # Crosswalk labs included and remove out of scope labs
  raw_sun <- left_join(raw_sun, sun_test_code,
                       by = c("TestCode" = "SUN_TestCode"))

  # Determine if test is included based on crosswalk results
  raw_sun <- raw_sun %>%
    mutate(TestIncl = !is.na(Test)) %>%
    filter(TestIncl)

  # Crosswalk unit type
  raw_sun <- left_join(raw_sun, sun_setting,
                       by = c("LocType" = "LocType"))

  # Crosswalk site name
  raw_sun <- left_join(raw_sun, mshs_site,
                       by = c("HospCode" = "DataSite"))

  # Crosswalk units and identify ICUs
  raw_sun <- raw_sun %>%
    mutate(LocandName = paste(LocCode, LocName))

  raw_sun <- left_join(raw_sun, sun_icu[, c("Concatenate", "ICU")],
                       by = c("LocandName" = "Concatenate"))

  raw_sun[is.na(raw_sun$ICU), "ICU"] <- FALSE

  # # Sunquest data formatting-----------------------------
  # Preprocess Sunquest data and add any necessary columns
  raw_sun <- raw_sun %>%
    mutate(
      # Determine if unit is an ICU based on crosswalk results
      ICU = ifelse(is.na(ICU), FALSE, ICU),
      # Create a column for resulted date
      ResultedDate = as.Date(ResultDateTime, format = "%m/%d/%Y"),
      # Create master setting column to identify ICU and IP Non-ICU units
      MasterSetting = ifelse(SettingRollUp == "IP" & ICU, "ICU",
                             ifelse(SettingRollUp == "IP" & !ICU,
                                    "IP Non-ICU", SettingRollUp)),
      # Create dashboard setting column to roll up master settings based on
      # desired dashboard grouping(ie, group ED and ICU together)
      DashboardSetting = ifelse(MasterSetting %in% c("ED", "ICU"), "ED & ICU",
                                MasterSetting),
      #
      # Create column with adjusted priority based on operational assumption
      # that all ED and ICU labs are treated as stat
      AdjPriority = ifelse(MasterSetting %in% c("ED", "ICU") |
                             SpecimenPriority %in% "S", "Stat", "Routine"),
      #
      # Create dashboard priority column
      DashboardPriority = ifelse(
        tat_targets$Priority[match(Test, tat_targets$Test)] == "All", "All",
        AdjPriority),
      #
      # Calculate turnaround times
      CollectToReceive =
        as.numeric(ReceiveDateTime - CollectDateTime, units = "mins"),
      ReceiveToResult =
        as.numeric(ResultDateTime - ReceiveDateTime, units = "mins"),
      CollectToResult =
        as.numeric(ResultDateTime - CollectDateTime, units = "mins"),
      #
      # Determine if order was an add on or original order based on time between
      # order and receive times
      AddOnMaster = ifelse(as.numeric(OrderDateTime - ReceiveDateTime,
                                      units = "mins") > 5, "AddOn", "Original"),
      #
      # Determine if collection time is missing
      MissingCollect = CollectDateTime == OrderDateTime,
      #
      # Determine TAT target based on test, priority, and patient setting
      # Create column concatenating test and priority to determine TAT targets
      Concate1 = paste(Test, DashboardPriority),
      # Create column concatenating test, priority, and setting to determine
      # TAT targets
      Concate2 = paste(Test, DashboardPriority, MasterSetting),
      # Determine Receive to Result TAT target using this logic:
      # 1. Try to match test, priority, and setting (applicable for labs with
      # different TAT targets based on patient setting and order priority)
      # 2. Try to match test and priority (applicable for labs with different
      # TAT targets based on order priority)
      # 3. Try to match test - this is for tests with (applicable for labs with
      # TAT targets that are independent of patient setting or priority)
      #
      # Determine Receive to Result TAT target based on above logic/scenarios
      ReceiveResultTarget =
        # Match on scenario 1
        ifelse(!is.na(match(Concate2, tat_targets$Concate)),
               tat_targets$ReceiveToResultTarget[
                 match(Concate2, tat_targets$Concate)],
               # Match on scenario 2
               ifelse(!is.na(match(Concate1, tat_targets$Concate)),
                      tat_targets$ReceiveToResultTarget[
                        match(Concate1, tat_targets$Concate)],
                      # Match on scenario 3
                      tat_targets$ReceiveToResultTarget[
                        match(Test, tat_targets$Concate)])),
      #
      # Determine Collect to Result TAT target based on above logic/scenarios
      CollectResultTarget =
        # Match on scenario 1
        ifelse(!is.na(match(Concate2, tat_targets$Concate)),
               tat_targets$CollectToResultTarget[
                 match(Concate2, tat_targets$Concate)],
               # Match on scenario 2
               ifelse(!is.na(match(Concate1, tat_targets$Concate)),
                      tat_targets$CollectToResultTarget[
                        match(Concate1, tat_targets$Concate)],
                      # Match on scenario 3
                      tat_targets$CollectToResultTarget[
                        match(Test, tat_targets$Concate)])),
      #
      # Determine if Receive to Result and Collect to Result TAT meet targets
      ReceiveResultInTarget = ReceiveToResult <= ReceiveResultTarget,
      CollectResultInTarget = CollectToResult <= CollectResultTarget,
      #
      # Create column with patient name, order ID, test, collect, receive, and
      # result date and determine if there is a duplicate; order time excluded
      Concate3 = paste(PtNumber, HISOrderNumber, TSTName,
                       CollectDateTime, ReceiveDateTime, ResultDateTime),
      DuplTest = duplicated(Concate3),
      #
      # Determine whether or not to include this particular lab in TAT analysis
      # Exclusion criteria:
      # 1. Add on orders
      # 2. Orders from "Other" settings
      # 3. Orders with collect or receive times after result time
      # 4. Orders with missing collect, receive, or result timestamps
      TATInclude = ifelse(AddOnMaster == "AddOn" |
                            MasterSetting == "Other" |
                            CollectToReceive < 0 |
                            CollectToResult < 0 |
                            ReceiveToResult < 0 |
                            is.na(CollectToResult) |
                            is.na(ReceiveToResult), FALSE, TRUE))

  # Remove duplicate tests
  raw_sun <- raw_sun %>%
    filter(!DuplTest)

  # Select columns
  sun_master <- raw_sun[, c("LocCode", "LocName", "LocandName",
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

  # Save output data to list
  scc_sun_list <- list(raw_scc, raw_sun, scc_sun_master)
  #
  return(scc_sun_list)

}

# Custom function to determine correct number of dates -------
# Remove labs from master data frame that were resulted at exactly midnight on
# next morning or prior day
correct_scc_result_dates <- function(data, number_days) {
  all_resulted_dates_vol <- data %>%
    group_by(ResultDate) %>%
    summarize(VolLabs = n()) %>%
    arrange(desc(VolLabs)) %>%
    ungroup()

  correct_dates <- all_resulted_dates_vol$ResultDate[1:number_days]

  new_data <- data %>%
    filter(ResultDate %in% correct_dates)

  return(new_data)
}


# Custom function to subset and summarize data for each lab division ----------
summarize_cp_tat <- function(x, lab_division) {
  # Subset data to be included based on lab division, whether or not TAT
  # meets inclusion criteria, and site location
  lab_div_df <- x %>%
    filter(Division == lab_division &
             Site %in% city_sites)
  #
  # Summarize data based on test, site, priority, setting, and TAT targets.
  lab_summary <- lab_div_df %>%
    group_by(Test,
             Site,
             DashboardPriority,
             DashboardSetting,
             ReceiveResultTarget,
             CollectResultTarget) %>%
    summarize(ResultedVolume = sum(TotalResultedTAT),
              ReceiveResultInTarget = sum(TotalReceiveResultInTarget),
              CollectResultInTarget = sum(TotalCollectResultInTarget),
              ReceiveResultPercent = round(
                ReceiveResultInTarget / ResultedVolume, digits = 3),
              CollectResultPercent = round(
                CollectResultInTarget / ResultedVolume, digits = 3),
              .groups = "keep") %>%
    ungroup()
  #
  # Subset template data frame for this division
  lab_div_df_templ <- tat_dashboard_templ %>%
    mutate(Incl = NULL) %>%
    filter(Division == lab_division)
  #
  # Combine lab summary with template data frame for this division for
  # dashboard visualization
  lab_summary <- left_join(lab_div_df_templ, lab_summary,
                           by = c("Test" = "Test",
                                  "Site" = "Site",
                                  "DashboardPriority" = "DashboardPriority",
                                  "DashboardSetting" = "DashboardSetting"))
  #
  # Format relevant columns as factors, look up target TAT for labs with 0
  # resulted volume, add formatting for percent within targets
  lab_summary <- lab_summary %>%
    mutate(
      #
      # Set test, site, priority, and setting as factors
      Test = droplevels(factor(Test, levels = test_names, ordered = TRUE)),
      Site = droplevels(factor(Site, levels = city_sites, ordered = TRUE)),
      DashboardPriority = droplevels(factor(DashboardPriority,
                                            levels = dashboard_priority_order,
                                            ordered = TRUE)),
      DashboardSetting = droplevels(factor(DashboardSetting,
                                           levels = dashboard_pt_setting,
                                           ordered = TRUE)),
      #
      # Determine TAT target for sites with 0 resulted labs
      # Create column concatenating test and priority to determine TAT targets
      Concate1 = paste(Test, DashboardPriority),
      # Create column concatenating test, priority, and setting to determine
      # TAT targets
      Concate2 = paste(Test, DashboardPriority, DashboardSetting),
      # Determine Receive to Result TAT target using this logic:
      # 1. Try to match test, priority, and setting (applicable for labs with
      # different TAT targets based on patient setting and order priority)
      # 2. Try to match test and priority (applicable for labs with different
      # TAT targets based on order priority)
      # 3. Try to match test - this is for tests with (applicable for labs with
      # TAT targets that are independent of patient setting or priority)
      #
      # Determine Receive to Result TAT target based on above logic/scenarios
      ReceiveResultTarget =
        # If TAT target is known, keep TAT target
        ifelse(!is.na(ReceiveResultTarget), ReceiveResultTarget,
               # Try to match on scenario 1
               ifelse(
                 !is.na(match(Concate2, tat_targets$Concate)),
                 tat_targets$ReceiveToResultTarget[
                   match(Concate2, tat_targets$Concate)],
                 # Try to match on scenario 2
                 ifelse(
                   !is.na(match(Concate1, tat_targets$Concate)),
                   tat_targets$ReceiveToResultTarget[
                     match(Concate1, tat_targets$Concate)],
                   # Try to match on scenario 3
                   tat_targets$ReceiveToResultTarget[
                     match(Test, tat_targets$Concate)]))),
      #
      # Determine Collect to Result TAT target based on above logic/scenarios
      # Determine Receive to Result TAT target based on above logic/scenarios
      CollectResultTarget =
        # If TAT target is known, keep TAT target
        ifelse(!is.na(CollectResultTarget), CollectResultTarget,
               # Try to match on scenario 1
               ifelse(
                 !is.na(match(Concate2, tat_targets$Concate)),
                 tat_targets$CollectToResultTarget[
                   match(Concate2, tat_targets$Concate)],
                 # Try to match on scenario 2
                 ifelse(
                   !is.na(match(Concate1, tat_targets$Concate)),
                   tat_targets$CollectToResultTarget[
                     match(Concate1, tat_targets$Concate)],
                   # Try to match on scenario 3
                   tat_targets$CollectToResultTarget[
                     match(Test, tat_targets$Concate)]))),
      #
      # Format target TAT for tables from numbers to "<=X min"
      ReceiveResultTarget = paste0("<=", ReceiveResultTarget, " min"),
      CollectResultTarget = paste0("<=", CollectResultTarget, " min"),
      #
      # Format percentage of labs in target
      ReceiveResultPercent = percent(ReceiveResultPercent, digits = 0),
      CollectResultPercent = percent(CollectResultPercent, digits = 0),
      #
      # Apply conditional color formatting to TAT percentages based on status
      # definitions for each lab division
      #
      # Chemistry & Hematology:
      # Green: >= 95%, Yellow: >= 80% & < 95%, Red: < 80%
      # Microbiology:
      # Green: 100%, Yellow: >= 90% & < 100%, Red: < 90%
      #
      ReceiveResultPercent = cell_spec(
        ReceiveResultPercent, "html",
        color = ifelse(is.na(ReceiveResultPercent), "lightgray",
                       ifelse(
                         (ReceiveResultPercent >= 0.95 &
                            lab_division %in% c("Chemistry", "Hematology")) |
                           (ReceiveResultPercent == 1.00 &
                              lab_division %in% c("Microbiology RRL")),
                         "green",
                         ifelse(
                           (ReceiveResultPercent >= 0.8 &
                              lab_division %in% c("Chemistry", "Hematology")) |
                             (ReceiveResultPercent >= 0.9 &
                                lab_division %in% c("Microbiology RRL")),
                           "orange", "red")))),
      CollectResultPercent = cell_spec(
        CollectResultPercent, "html",
        color = ifelse(is.na(CollectResultPercent), "lightgray",
                       ifelse(
                         (CollectResultPercent >= 0.95 &
                            lab_division %in% c("Chemistry", "Hematology")) |
                           (CollectResultPercent == 1.00 &
                              lab_division %in% c("Microbiology RRL")),
                         "green",
                         ifelse(
                           (CollectResultPercent >= 0.8 &
                              lab_division %in% c("Chemistry", "Hematology")) |
                             (CollectResultPercent >= 0.9 &
                                lab_division %in% c("Microbiology RRL")),
                           "orange", "red")))),
      #
      # Create a new column with test and priority to be used in tables later
      TestAndPriority = paste(Test, "-", DashboardPriority, "Labs"),
      #
      # Remove concatenated columns used for matching
      Concate1 = NULL,
      Concate2 = NULL) %>%
    arrange(Test, Site, DashboardPriority, DashboardSetting)
  #
  # Melt summarized data into a long dataframe
  lab_dashboard_melt <- melt(lab_summary,
                             id.var = c("Test",
                                        "Site",
                                        "DashboardPriority",
                                        "TestAndPriority",
                                        "DashboardSetting",
                                        "ReceiveResultTarget",
                                        "CollectResultTarget"),
                             measure.vars = c("ReceiveResultPercent",
                                              "CollectResultPercent"))
  #
  # Case dataframe into wide format for use in tables later
  lab_dashboard_cast <- dcast(lab_dashboard_melt,
                              Test +
                                DashboardPriority +
                                TestAndPriority +
                                DashboardSetting +
                                ReceiveResultTarget +
                                CollectResultTarget ~
                                variable +
                                Site,
                              value.var = "value")
  #
  # Rearrange columns based on desired dashboard aesthetics
  lab_dashboard_cast <- lab_dashboard_cast %>%
    mutate(DashboardSetting2 = DashboardSetting) %>%
    select(Test, DashboardPriority, TestAndPriority,
           ReceiveResultTarget, DashboardSetting,
           ReceiveResultPercent_MSH, ReceiveResultPercent_MSQ,
           ReceiveResultPercent_MSBI, ReceiveResultPercent_MSB,
           ReceiveResultPercent_MSW, ReceiveResultPercent_MSM,
           CollectResultTarget, DashboardSetting2,
           CollectResultPercent_MSH, CollectResultPercent_MSQ,
           CollectResultPercent_MSBI, CollectResultPercent_MSB,
           CollectResultPercent_MSW, CollectResultPercent_MSM)
  #
  # Save outputs in a list
  lab_sub_output <- list(lab_div_df,
                         lab_summary,
                         lab_dashboard_melt,
                         lab_dashboard_cast)
  #
  lab_sub_output

}

# Custom function for creating kables for each CP lab division ----------------
kable_cp_tat <- function(x) {
  #
  # Select columns 3 and on
  data <- x[, c(3:ncol(x))]
  #
  # Format kable
  kable(data, format = "html", escape = FALSE, align = "c",
        col.names = c("Test & Priority",
                      "Target", "Setting",
                      "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM",
                      "Target", "Setting",
                      "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM")) %>%
    kable_styling(bootstrap_options = "hover", position = "center",
                  font_size = 11) %>%
    column_spec(column = c(1, 9, 17),
                border_right = "thin solid lightgray") %>%
    add_header_above(c(" " = 1,
                       "Receive to Result Within Target" =
                         (ncol(data) - 1) / 2,
                       "Collect to Result Within Target" =
                         (ncol(data) - 1) / 2),
                     background = c("white", "#00AEEF", "#221f72"),
                     color = "white", line = FALSE, font_size = 13) %>%
    column_spec(column = 2:9, background = "#E6F8FF", color = "black") %>%
    column_spec(column = 10:17, background = "#EBEBF9", color = "black") %>%
    #column_spec(column = 2:17, background = "inherit", color = "inherit") %>%
    column_spec(column = 1, width_min = "125px") %>%
    column_spec(column = c(3, 11), width_min = "100px") %>%
    row_spec(row = 0, font_size = 13) %>%
    collapse_rows(columns = c(1, 2, 10))
}

# Custom function for summarizing resulted lab volume from prior day(s) --------
summarize_cp_vol <- function(x, lab_division) {
  # Subset data to be included based on lab division and site location
  lab_div_vol_df <- x %>%
    filter(Division == lab_division &
             Site %in% city_sites) %>%
    group_by(Site,
             Test,
             DashboardPriority,
             MasterSetting) %>%
    summarize(ResultedLabs = sum(TotalResulted),
              .groups = "keep")
  #
  # Subset volume dataframe template for this division
  lab_div_vol_templ <- vol_dashboard_templ %>%
    filter(Division == lab_division) %>%
    mutate(Incl = NULL)
  #
  # Combine two dataframes to ensure all combinations are accounts for
  lab_div_vol_df <- left_join(lab_div_vol_templ, lab_div_vol_df,
                              by = c("Test" = "Test",
                                     "Site" = "Site",
                                     "DashboardPriority" = "DashboardPriority",
                                     "PtSetting" = "MasterSetting"))
  #
  lab_div_vol_df <- lab_div_vol_df %>%
    mutate(
      # Set test, site, priority, and setting as factors
      Test = droplevels(factor(Test, levels = test_names, ordered = TRUE)),
      Site = droplevels(factor(Site, levels = city_sites, ordered = TRUE)),
      DashboardPriority = droplevels(factor(DashboardPriority,
                                            levels = dashboard_priority_order,
                                            ordered = TRUE)),
      PtSetting = droplevels(factor(PtSetting,
                                    levels = pt_setting_order,
                                    ordered = TRUE)),
      #
      # Replace NA with 0
      ResultedLabs = ifelse(is.na(ResultedLabs), 0, ResultedLabs),
      #
      # Create column with test name and priority
      TestAndPriority = paste(Test, "-", DashboardPriority, "Labs"))
  #
  # Cast dataframe
  lab_div_vol_cast <- dcast(lab_div_vol_df,
                            Test +
                              DashboardPriority +
                              TestAndPriority +
                              PtSetting ~
                              Site,
                            value.var = "ResultedLabs")
  # Remove test and priority columns
  lab_div_vol_cast <- lab_div_vol_cast[, c(3:ncol(lab_div_vol_cast))]
  #
  return(lab_div_vol_cast)
}

# Custom function for creating a kable of lab volume from prior day(s)----------
kable_cp_vol <- function(x) {
  kable(x, format = "html", escape = FALSE, align = "c",
        col.names = c("Test & Priority", "Setting",
                      "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM")) %>%
    kable_styling(bootstrap_options = "hover",
                  position = "center",
                  font_size = 11) %>%
    column_spec(column = c(1, 8), border_right = "thin solid lightgray") %>%
    add_header_above(c(" " = 1,
                       "Resulted Lab Volume" = (ncol(x) - 1)),
                     background = c("white", "#00AEEF"),
                     color = "white",
                     line = FALSE,
                     font_size = 13) %>%
    column_spec(column = 2:8, background = "#E6F8FF", color = "black") %>%
    # column_spec(column = 2:8,
    #             background = "inherit",
    #             color = "inherit") %>%
    # column_spec(column = 1,
    #             width_min = "125px",
    #             include_thead = TRUE) %>%
    # column_spec(column = c(3, 11),
    #             width_min = "100px",
    #             include_thead = TRUE) %>%
    row_spec(row = 0, font_size = 13) %>%
    collapse_rows(columns = c(1, 2))
}

# Custom function for creating a kable of labs with missing collections --------
kable_missing_collections <- function(x) {
  # Filter data for city sites and summarize
  missing_collect <- x %>%
    filter(Site %in% city_sites) %>%
    group_by(Site) %>%
    summarize(ResultedVolume = sum(TotalResultedTAT),
              MissingCollection = sum(TotalMissingCollections, na.rm = TRUE),
              Percent = percent(MissingCollection / ResultedVolume,
                                digits = 0),
              .groups = "keep") %>%
    ungroup() %>%
    mutate(
      # Apply conditional formatting based on percentage of labs with missing
      # collections
      Percent = cell_spec(
        Percent, "html",
        color = ifelse(is.na(Percent), "grey",
                       ifelse(Percent <= 0.05, "green",
                              ifelse(Percent <= 0.15, "orange", "red")))),
      # Format site as factors
      Site = factor(Site, levels = city_sites, ordered = TRUE))
  #
  # Create template to ensure all sites are included
  missing_collect <- left_join(data.frame("Site" = factor(city_sites,
                                                          levels = city_sites,
                                                          ordered = TRUE)),
                               missing_collect,
                               by = c("Site" = "Site"))
  #
  # Cast missing collections into table format
  missing_collect_table <- dcast(missing_collect,
                                 "Percentage of Specimens" ~ Site,
                                 value.var = "Percent")
  # Create kable with summarized data
  missing_collect_table %>%
    kable(format = "html", escape = FALSE, align = "c",
          col.names = c("Site", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM")) %>%
    kable_styling(
      bootstrap = "hover",
      position = "float_left",
      font_size = 11,
      full_width = FALSE) %>%
    add_header_above(
      c(" " = 1,
        "Percentage of Labs Missing Collect Times" =
          ncol(missing_collect_table) - 1),
      background = c("white", "#00AEEF"),
      color = "white",
      line = FALSE,
      font_size = 13) %>%
    column_spec(column = c(1, ncol(missing_collect_table)),
                border_right = "thin solid lightgray") %>%
    column_spec(column = c(2:ncol(missing_collect_table)),
                background = "#E6F8FF",
                color = "black") %>%
    # column_spec(column = c(2:ncol(missing_collect_table)),
    #             background = "inherit",
    #             color = "inherit",
    #             width_max = 0.15) %>%
    row_spec(row = 0, font_size = 13)
}

# Custom function for creating a kable of add-on order volume
kable_add_on_volume <- function(x) {
  # Filter data for city sites and summarize
  add_on_volume <- x %>%
    filter(Site %in% city_sites) %>%
    group_by(Test, Site) %>%
    summarize(AddOnVolume = sum(TotalAddOnOrder, na.rm = TRUE),
              .groups = "keep") %>%
    ungroup()

  add_on_volume <- left_join(test_site_comb, add_on_volume,
                             by = c("Site" = "Site",
                                    "Test" = "Test"))

  add_on_volume <- add_on_volume %>%
    mutate(
      # Set test and site as factors
      Test = droplevels(factor(Test, levels = test_names, ordered = TRUE)),
      Site = factor(Site, levels = city_sites, ordered = TRUE),
      AddOnVolume = ifelse(is.na(AddOnVolume), 0, AddOnVolume))

  add_on_table <- dcast(add_on_volume, Test ~ Site, value.var = "AddOnVolume")

  # Create kable of add on orders
  add_on_table %>%
    kable(format = "html", escape = FALSE, align = "c",
          col.names = c("Test", "MSH", "MSQ", "MSBI", "MSB", "MSW", "MSM"),
          color = "gray") %>%
    kable_styling(
      bootstrap = "hover",
      position = "right",
      font_size = 11,
      full_width = FALSE) %>%
    add_header_above(
      c(" " = 1,
        "Volume of Add On Labs" = ncol(add_on_table) - 1),
      background = c("white", "#00AEEF"),
      color = "white",
      line = FALSE,
      font_size = 13) %>%
    column_spec(
      column = c(1, ncol(add_on_table)),
      border_right = "thin solid lightgray") %>%
    column_spec(
      column = c(2:ncol(add_on_table)),
      background = "#E6F8FF", color = "black") %>%
    # column_spec(column = c(2:ncol(add_on_table)),
    #             background = "inherit",
    #             color = "inherit") %>%
    row_spec(row = 0, font_size = 13)
}