# Libraries ---------------------------------------------------------------
library(tidyverse)
library(rstudioapi)
library(zip)
library(readxl)

# Assigning Directory(ies) ------------------------------------------------
prod_path <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/",
                    "Productivity/")

#User selected directory where output files could be sent
output_folder <- rstudioapi::selectDirectory(caption = "select output folder")

# Constants ---------------------------------------------------------------
# Define constants that will be used throughout the code. These are the
# variables that are calculated here and not changed in the rest of the code.


# Data Import -------------------------------------------------------------
#Selecting and unzipping batch download file to get file paths
bd_zip_file <- choose.files(caption = "select batch download zip file")

unzip(bd_zip_file, exdir = output_folder)

zip_folders <-
  list.dirs(
    path = output_folder,
    full.names = TRUE
  )

zip_file_paths <-
  list.files(
    path = zip_folders,
    full.names = TRUE
  )

zip_file_paths <- as.data.frame(zip_file_paths)

#Importing data files
rep_def <- read_xlsx(paste0(prod_path, "Universal Data/Mapping/",
                            "MSHS_Reporting_Definition_Mapping.xlsx"))

admin_def_field_all <- read.csv(paste0(prod_path, "Universal Data/Premier/",
                                       "Dictionary Exports/AdminDef(All Entity).csv"), header = F)

admin_def_field_multi <- read.csv(paste0(prod_path, "Universal Data/Premier/",
                                         "Dictionary Exports/AdminDef(Multi-Entity).csv"), header = F) 

admin_rep_dep_all <- read.csv(paste0(prod_path, "Universal Data/Premier/",
                                     "Dictionary Exports/",
                                     "AdministrativeReportingDepts(All Entity).csv"), header = F)

admin_rep_dep_multi <- read.csv(paste0(prod_path, "Universal Data/Premier/",
                                       "Dictionary Exports/",
                                       "AdministrativeReportingDepts(Multi-Entity).csv"), header = F)

dep_def_all <- read.csv(paste0(prod_path, "Universal Data/Premier/",
                               "Dictionary Exports/DepartmentDef.csv"), header = F)

dates <- read_xlsx(paste0(prod_path, "Universal Data/Mapping/",
                          "MSHS_Pay_Cycle.xlsx"))

#Adding column names
colnames(admin_def_field_all) <- c("Corporation ID", "Hospital ID",	"Admin Definition Name", 
                                   "Admin Definition Code", "Threshold Type",	"Effective Date",
                                   "Time Period Cycle Code", "Single or Multi Entity")

colnames(admin_def_field_multi) <- c("Corporation ID", "Hospital ID",	"Admin Definition Name", 
                                     "Admin Definition Code", "Threshold Type",	"Effective Date",
                                     "Time Period Cycle Code", "Single or Multi Entity")

colnames(admin_rep_dep_all) <- c("Corporation ID", "Hospital ID", "Admin Definition Code",
                                 "Effective Date", "Expiration Date", "Entity Code", "DRD Code",
                                 "Action", "Single or Multi Entity")

colnames(admin_rep_dep_multi) <- c("Corporation ID", "Hospital ID", "Admin Definition Code",
                                   "Effective Date", "Expiration Date", "Entity Code", "DRD Code",
                                   "Action", "Single or Multi Entity")

colnames(dep_def_all) <- c("Corporation ID", "Hospital ID", "Reporting Definition Name",
                           "Reporting Definition Code", "Department Code",
                           "Internal Department Category", "Threshold Type", "Target Type",
                           "Exclude from Rollup Report", "Effective Date", "Pay Cycle Code",
                           "Exclude from Admin Rollup Report", "Exclude from Action Plan")

#Combining dictionaries into data frame with only the columns needed
System_Dep_Mapping <- left_join(admin_rep_dep_multi, admin_def_field_multi, 
                                by = c("Admin Definition Code" = "Admin Definition Code")) %>%
  select(`Corporation ID.x`, `Hospital ID.x`, `Entity Code`, `DRD Code`, `Admin Definition Code`,
         `Admin Definition Name`)

#Table of distribution dates
dist_dates <- dates %>%
  select(END.DATE, PREMIER.DISTRIBUTION) %>%
  distinct() %>%
  drop_na() %>%
  arrange(END.DATE) %>%
  #filter only on distribution end dates
  filter(PREMIER.DISTRIBUTION %in% c(TRUE, 1),
         #filter 3 weeks from run date (21 days) for data collection lag before run date
         END.DATE < as.POSIXct(Sys.Date() - 21))
#Table of non-distribution dates
non_dist_dates <- dates %>%
  select(END.DATE, PREMIER.DISTRIBUTION) %>%
  distinct() %>%
  drop_na() %>%
  arrange(END.DATE) %>%
  #filter only on distribution end dates
  filter(PREMIER.DISTRIBUTION %in% c(FALSE, 0),
         #filter 3 weeks from run date (21 days) for data collection lag before run date
         END.DATE < as.POSIXct(Sys.Date() - 21))
#Selecting current and previous distribution dates
distribution <- format(dist_dates$END.DATE[nrow(dist_dates)],"%m/%d/%Y")
previous_distribution <- format(dist_dates$END.DATE[nrow(dist_dates)-1],"%m/%d/%Y")
#Confirming distribution dates
cat("Current distribution is", distribution,
    "\nPrevious distribution is", previous_distribution)
answer <- select.list(choices = c("Yes", "No"),
                      preselect = "Yes",
                      multiple = F,
                      title = "Correct distribution?",
                      graphics = T)
if (answer == "No") {
  distribution <- select.list(choices =
                                format(sort.POSIXlt(dist_dates$END.DATE, decreasing = T),
                                       "%m/%d/%Y"),
                              multiple = F,
                              title = "Select current distribution",
                              graphics = T)
  which(distribution == format(dist_dates$END.DATE, "%m/%d/%Y"))
  previous_distribution <- format(dist_dates$END.DATE[which(distribution == format(dist_dates$END.DATE, "%m/%d/%Y"))-1],"%m/%d/%Y")
}


# Data References ---------------------------------------------------------
# (aka Mapping Tables)
# Files that need to be imported for mappings and look-up tables.
# (This section may be combined into the Data Import section.)


# Creation of Functions --------------------------------------------------
# These are functions that will be commonly used within the rest of the script.
# It might make sense to keep these files in a separate file that is sourced
# in the "Source Global Functions" section above.


# Data Pre-processing -----------------------------------------------------
# Cleaning raw data and ensuring that all values are accounted for such as
# blanks and NA. As well as excluding data that may not be used or needed. This
# section can be split into multiple ones based on the data pre-processing
# needed.
# One of the first steps could be to perform initial checks to make sure data is
# in the correct format.  This might also be done as soon as the data is
# imported.


# Data Formatting ---------------------------------------------------------
# How the data will look during the output of the script.
# For example, if you have a data table that needs the numbers to show up as
# green or red depending on whether they meet a certain threshold.


# Quality Checks ----------------------------------------------------------
# Checks that are performed on the output to confirm data consistency and
# expected outputs.


# Visualization -----------------------------------------------------------
# How the data will be plotted or how the data table will look including axis
# titles, scales, and color schemes of graphs or data tables.
# (This section may be combined with the Data Formatting section.)


# File Saving -------------------------------------------------------------
# Writing files or data for storage


# Script End --------------------------------------------------------------
