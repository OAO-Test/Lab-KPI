#######
# Code for importing the raw data needed to carry-on the first run
# for the lab KPI daily dashboard with different logic based on the DOW.
# Imported data includes:
# 1. SCC data for clinical pathology
# 2. SunQuest data for clinical pathology
# 3. PowerPath data for anatomic pathology
# 4. Epic data for anatomic pathology especially cytology
# 5. Backlog data for anatomic pathology especially cytology-----
#######

#------------------------------Read Excel sheets------------------------------#

#Import weekday files
scc_weekday <- read_excel(
  choose.files(
    default = paste0(user_directory, "/SCC CP Reports/*.*"),
    caption = "Select SCC Report for Labs Resulted on Most Recent Weekday"),
  sheet = 1, col_names = TRUE)

sq_weekday <- suppressWarnings(read_excel(
  choose.files(
    default = paste0(user_directory, "/SUN CP Reports/*.*"),
    caption =
      "Select Sunquest Report for Labs Resulted on Most Recent Weekday"),
  sheet = 1, col_names = TRUE))

pp_weekday <- read_excel(
  choose.files(
    default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
    caption =
      "Select PowerPath Report for Spec Signed Out on Most Recent Weekday"),
  skip = 1, 1)

pp_weekday <- data.frame(pp_weekday[-nrow(pp_weekday), ],
                         stringsAsFactors = FALSE)

epic_weekday <- read_excel(
  choose.files(
    default = paste0(user_directory, "/EPIC Cytology/*.*"),
    caption = "Select Epic Report for Spec Signed Out on Most Recent Weekday"),
  1)

# The if-statement below determines the number of raw data files to import
# based on the day of week and holidays. The user is then prompted to select
# each file from the relevant folder.

if (((holiday_det) & (yesterday_day == "Mon")) |
    ((yesterday_day == "Sun") & (isHoliday(as.timeDate(yesterday) - (86400 * 2))))) {
  # Scenario 1: Mon Holiday or Friday Holiday (Need to select 4 files)
  # Save scenario
  scenario <- 1
  # Import SCC data
  scc_hol_mon_fri <- read_excel(
    choose.files(
      default = paste0(user_directory, "/SCC CP Reports/*.*"),
      caption = "Select SCC Report for Labs Resulted on Recent Holiday"),
    sheet = 1, col_names = TRUE)
  scc_sun <- read_excel(
    choose.files(
      default = paste0(user_directory, "/SCC CP Reports/*.*"),
      caption = "Select SCC Report for Labs Resulted on Sunday"),
    sheet = 1, col_names = TRUE)
  scc_sat <- read_excel(
    choose.files(
      default = paste0(user_directory, "/SCC CP Reports/*.*"),
      caption = "Select SCC Report for Labs Resulted on Saturday"),
    sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  scc_not_weekday <- rbind(scc_hol_mon_fri,
                           scc_sun,
                           scc_sat)
  #
  # Import Sunquest data
  sq_hol_mon_fri <- suppressWarnings(read_excel(
    choose.files(
      default = paste0(user_directory, "/SUN CP Reports/*.*"),
      caption = "Select Sunquest Report for Labs Resulted on Recent Holiday"),
    sheet = 1, col_names = TRUE))
  sq_sun <- suppressWarnings(read_excel(
    choose.files(
      default = paste0(user_directory, "/SUN CP Reports/*.*"),
      caption = "Select Sunquest Report for Labs Resulted on Sunday"),
    sheet = 1, col_names = TRUE))
  sq_sat <- suppressWarnings(read_excel(
    choose.files(
      default = paste0(user_directory, "/SUN CP Reports/*.*"),
      caption = "Select Sunquest Report for Labs Resulted on Saturday"),
    sheet = 1, col_names = TRUE))
  #Merge the weekend data with the holiday data in one data frame
  sq_not_weekday <- rbind(sq_hol_mon_fri, sq_sun, sq_sat)
  #
  # Import Powerpath data
  pp_hol_mon_fri <- read_excel(
    choose.files(
      default =
        paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
      caption =
        "Select PowerPath Report for Spec Signed Out on Recent Holiday"),
    skip = 1, 1)
  pp_hol_mon_fri  <- pp_hol_mon_fri[
    -nrow(pp_hol_mon_fri), ]
  pp_sun <- read_excel(
    choose.files(
      default =
        paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
      caption =
        "Select PowerPath Report for Spec Signed Out on Sunday"),
    skip = 1, 1)
  pp_sun <- pp_sun[-nrow(pp_sun), ]
  pp_sat <- read_excel(
    choose.files(
      default =
        paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
      caption =
        "Select PowerPath Report for Spec Signed Out on Saturday"),
    skip = 1, 1)
  pp_sat <- pp_sat[-nrow(pp_sat), ]
  #Merge the weekend data with the holiday data in one data frame
  pp_not_weekday <- data.frame(rbind(pp_hol_mon_fri,
                                     pp_sun,
                                     pp_sat), stringsAsFactors = FALSE)
  #
  # Import EPIC Cytology data
  epic_hol_mon_fri <- read_excel(
    choose.files(
      default = paste0(user_directory, "/EPIC Cytology/*.*"),
      caption = "Select Epic Report for Spec Signed Out on Recent Holiday"),
    1)
  epic_sun <- read_excel(
    choose.files(
      default = paste0(user_directory, "/EPIC Cytology/*.*"),
      caption = "Select Epic Report for Spec Signed Out on Sunday"), 1)
  epic_sat <- read_excel(
    choose.files(
      default = paste0(user_directory, "/EPIC Cytology/*.*"),
      caption = "Select PowerPath Report for Spec Signed Out on Saturday"), 1)
  #Merge the weekend data with the holiday data in one data frame
  epic_not_weekday <- data.frame(rbind(epic_hol_mon_fri,
                                       epic_sun,
                                       epic_sat), stringsAsFactors = FALSE)
  
} else if ((holiday_det) & (yesterday_day == "Sun")) {
  # Scenario 2: Regular Monday (Need to select 3 files)
  # Save scenario
  scenario <- 2
  #
  # Import SCC data
  scc_sun <- read_excel(
    choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"),
                 caption = "Select SCC Report for Labs Resulted on Sunday"),
    sheet = 1, col_names = TRUE)
  scc_sat <- read_excel(
    choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"),
                 caption = "Select SCC Report for Labs Resulted on Saturday"),
    sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  scc_not_weekday <- rbind(scc_sun, scc_sat)
  #
  # Import Sunquest data
  sq_sun <- suppressWarnings(read_excel(
    choose.files(
      default = paste0(user_directory, "/SUN CP Reports/*.*"),
      caption = "Select Sunquest Report for Labs Resulted on Sunday"),
    sheet = 1, col_names = TRUE))
  sq_sat <- suppressWarnings(read_excel(
    choose.files(
      default = paste0(user_directory, "/SUN CP Reports/*.*"),
      caption = "Select Sunquest Report for Labs Resulted on Saturday"),
    sheet = 1, col_names = TRUE))
  #Merge the weekend data with the holiday data in one data frame
  sq_not_weekday <- rbind(sq_sun, sq_sat)
  #
  # Import Powerpath data
  pp_sun <- read_excel(
    choose.files(
      default =
        paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
      caption =
        "Select PowerPath Report for Spec Signed Out on Sunday"),
    skip = 1, 1)
  pp_sun <- pp_sun[-nrow(pp_sun), ]
  pp_sat <- read_excel(
    choose.files(
      default =
        paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
      caption =
        "Select PowerPath Report for Spec Signed Out on Saturday"),
    skip = 1, 1)
  pp_sat <- pp_sat[-nrow(pp_sat), ]
  #Merge the weekend data with the holiday data in one data frame
  pp_not_weekday <- data.frame(rbind(pp_sun,
                                     pp_sat), stringsAsFactors = FALSE)
  #
  # Import Epic Cytology data
  epic_sun <- read_excel(
    choose.files(
      default = paste0(user_directory, "/EPIC Cytology/*.*"),
      caption = "Select Epic Report for Spec Signed Out on Sunday"), 1)
  epic_sat <- read_excel(
    choose.files(
      default = paste0(user_directory, "/EPIC Cytology/*.*"),
      caption = "Select Epic Report for Spec Signed Out on Saturday"), 1)
  #Merge the weekend data with the holiday data in one data frame
  epic_not_weekday <- data.frame(rbind(epic_sun,
                                       epic_sat),
                                 stringsAsFactors = FALSE)
  
} else if ((holiday_det) & ((yesterday_day != "Mon") |
                            (yesterday_day != "Sun"))) {
  #Scenario 3: Midweek holiday (Need to select 2 files)
  # Save scenario
  scenario <- 3
  #
  # Import SCC data
  scc_hol_weekday <- read_excel(
    choose.files(
      default = paste0(user_directory, "/SCC CP Reports/*.*"),
      caption = "Select SCC Report for Labs Resulted on Recent Holiday"),
    sheet = 1, col_names = TRUE)
  scc_not_weekday <- scc_hol_weekday
  #
  # Import Sunquest data
  sq_hol_weekday <- suppressWarnings(read_excel(
    choose.files(
      default = paste0(user_directory, "/SUN CP Reports/*.*"),
      caption = "Select Sunquest Report for Labs Resulted on Recent Holiday"),
    sheet = 1, col_names = TRUE))
  sq_not_weekday <- sq_hol_weekday
  #
  # Import Powerpath data
  pp_hol_weekday <- read_excel(
    choose.files(
      default =
        paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"),
      caption =
        "Select PowerPath Report for Spec Signed Out on Recent Holiday"),
    skip = 1, 1)
  pp_hol_weekday <- pp_hol_weekday[-nrow(pp_hol_weekday), ]
  pp_not_weekday <- data.frame(pp_hol_weekday, stringsAsFactors = FALSE)
  #
  # Import Epic Cytology data
  epic_hol_weekday <- read_excel(
    choose.files(
      default = paste0(user_directory, "/EPIC Cytology/*.*"),
      caption = "Select Epic Report for Spec Signed Out on Recent Holiday"), 1)
  epic_not_weekday <- data.frame(epic_hol_weekday, stringsAsFactors = FALSE)
  
} else {#Scenario 4: Tue-Fri without holidays (Need to select 1 file)
  # Save scenario
  scenario <- 4
  #
  scc_not_weekday <- NULL
  #
  sq_not_weekday <- NULL
  #
  pp_not_weekday <- NULL
  #
  epic_not_weekday <- NULL
}

#-----------Cytology Backlog Excel Files-----------#
#For the backlog files the read excel is starting from the second row and
#remove last row
cyto_backlog_raw <- read_excel(
  choose.files(default = paste0(user_directory,
                                "/Cytology Backlog Reports/*.*"),
               caption = "Select Cytology Backlog Report"),
  skip = 1, 1)
cyto_backlog_raw <- data.frame(
  cyto_backlog_raw[-nrow(cyto_backlog_raw), ], stringsAsFactors = FALSE)
