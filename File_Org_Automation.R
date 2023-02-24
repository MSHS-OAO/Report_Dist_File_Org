# Libraries ---------------------------------------------------------------
library(tidyverse)
library(rstudioapi)
library(zip)
library(readxl)
library(zoo)

# Assigning Directory(ies) ------------------------------------------------
prod_path <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/",
                    "Productivity/")

#User selected directory where output files could be sent
unzip_output_folder <- rstudioapi::selectDirectory(caption = "select unzip output folder")

final_output_folder <- rstudioapi::selectDirectory(caption = "select final output folder")

# Data Import -------------------------------------------------------------
#Selecting and unzipping batch download file to get file paths
bd_zip_file <- choose.files(caption = "select batch download zip file")

unzip(bd_zip_file, exdir = unzip_output_folder)

zip_folders <-
  list.dirs(
    path = unzip_output_folder,
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

dates <- read_xlsx(paste0(prod_path, "Universal Data/Mapping/",
                          "MSHS_Pay_Cycle.xlsx"))

VP_list <- rep_def %>%
  select(VP) %>%
  distinct() %>%
  filter(!grepl("Administrative Summary Report", VP)) %>%
  na.omit()

VP_list <- toupper(VP_list$VP)

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
distribution <- format(dist_dates$END.DATE[nrow(dist_dates)],"%m.%d.%y")
previous_distribution <- format(dist_dates$END.DATE[nrow(dist_dates)-1],"%m.%d.%y")
previous_cpt_distribution <- format(dist_dates$END.DATE[nrow(dist_dates)-2],"%m.%d.%y")
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
                                       "%m.%d.%y"),
                              multiple = F,
                              title = "Select current distribution",
                              graphics = T)
  which(distribution == format(dist_dates$END.DATE, "%m.%d.%y"))
  previous_distribution <- format(dist_dates$END.DATE[which(distribution == format(dist_dates$END.DATE, "%m.%d.%y"))-1],"%m.%d.%y")
}


# Data References ---------------------------------------------------------
# (aka Mapping Tables)
# Files that need to be imported for mappings and look-up tables.
# (This section may be combined into the Data Import section.)

#Copying folders to output folder
file.copy(from = "J:/deans/Presidents/SixSigma/Individual Folders/Current Employees/Engineers/Matthew Miesle/LPM/Report Distribution Prep/Default Folders v2/MSHS", 
          to = final_output_folder, recursive = T)

#Creating first mapping dataframe for admin reports and some department reports
Mapping_df <- data.frame('Source Paths' = zip_file_paths)
Mapping_df <- Mapping_df %>%
  data.frame('File Name' = basename(Mapping_df$zip_file_paths),
             APR = ifelse(str_detect(Mapping_df$zip_file_paths, "_APR_"), 1, 0),
             DPR = ifelse(str_detect(Mapping_df$zip_file_paths, "_DPR_"), 1, 0),
             Folder = ifelse(!str_detect(Mapping_df$zip_file_paths, ".pdf"), 1, 0),
             'Case Management' = ifelse(str_detect(Mapping_df$zip_file_paths, "CASE MANAGEMENT"), 1, 0),
             ED = ifelse(str_detect(Mapping_df$zip_file_paths, "EMERGENCY") | str_detect(Mapping_df$zip_file_paths, "_ER REGISTRATION"), 1, 0),
             Nursing = ifelse(str_detect(Mapping_df$zip_file_paths, "NURSING") & 
                                !str_detect(Mapping_df$zip_file_paths, "BOLIVER"), 1, 0),
             PLamb = ifelse(str_detect(Mapping_df$zip_file_paths, "PLAMB") & 
                                         !str_detect(Mapping_df$zip_file_paths, "CPT"), 1, 0),
             Radiology = ifelse(str_detect(Mapping_df$zip_file_paths, "RADIOLOGY"), 1, 0),
             'Supply Chain and Support Services' = ifelse(str_detect(Mapping_df$zip_file_paths, 
                                                                     "SUPPLY CHAIN AND SUPPORT SERVICES"), 1, 0), 
             'MSM CPT' = ifelse(str_detect(Mapping_df$zip_file_paths, "MSM_41") 
                                & str_detect(Mapping_df$zip_file_paths, previous_distribution) 
                                | str_detect(Mapping_df$zip_file_paths, "MSM_42") 
                                & str_detect(Mapping_df$zip_file_paths, previous_distribution), 1, 0),
             'MSW CPT' = ifelse(str_detect(Mapping_df$zip_file_paths, "MSW_15") 
                                & str_detect(Mapping_df$zip_file_paths, previous_distribution), 1, 0)) %>%
  filter(!str_detect(File.Name, "DNU"))

#Adding engineering column that includes all reports within the FACILITIES admin report
Engineering_subset <- subset(Mapping_df, str_detect(Mapping_df$zip_file_paths, "FACILITIES") & 
                        !str_detect(Mapping_df$zip_file_paths, ".pdf")) %>%
  select(File.Name)

Mapping_df <- Mapping_df %>%
  data.frame(Engineering = ifelse(str_detect(Mapping_df$zip_file_paths, paste(Engineering_subset$File.Name, collapse = '|')), 1, 0))

#Reading in APR Mapping and only selecting the first 9 rows and removing the first 8 columns 
APR_Mapping <- read_excel(paste0(prod_path, "R Programming/Report Distribution File Org Automation/", 
                                 "Brainstorming/Mapping example.xlsx"))[(1:9),-(1:8)] %>%
  select(-c(`System.COO`, `COO.Folders`))

Na_in_mapping <- Mapping_df[rowSums(is.na(Mapping_df)) > 0,]

APR_mapping_join <- left_join(Mapping_df, APR_Mapping) %>%
  na.omit()

APR_mapping_join$`Full Destination Path` = paste0(final_output_folder, APR_mapping_join$`Short Folder`, 
                                                    "/", APR_mapping_join$File.Name)

duplicated_rows <- duplicated(APR_mapping_join$File.Name)
APR_mapping_join <- APR_mapping_join[!duplicated_rows, ]

file.copy(from = APR_mapping_join$zip_file_paths, to = APR_mapping_join$`Full Destination Path`)

Mapping_df_pt2 <- Mapping_df %>%
  data.frame(MSB = ifelse(!str_detect(Mapping_df$File.Name, "MSBI") & str_detect(Mapping_df$File.Name, "MSB"), 1, 0),
             MSBI = ifelse(str_detect(Mapping_df$File.Name, "MSBI"), 1, 0),
             MSH = ifelse(str_detect(Mapping_df$File.Name, "MSH"), 1, 0),
             MSM = ifelse(str_detect(Mapping_df$File.Name, "MSM"), 1, 0),
             MSQ = ifelse(str_detect(Mapping_df$File.Name, "MSQ"), 1, 0),
             MSW = ifelse(str_detect(Mapping_df$File.Name, "MSW"), 1, 0),
             MSUS = ifelse(str_detect(Mapping_df$File.Name, "MSUS"), 1, 0),
             CCW = ifelse(str_detect(Mapping_df$File.Name, "CCW"), 1, 0),
             System.COO = ifelse(str_detect(Mapping_df$File.Name, paste(VP_list, collapse = '|')), 1, 0),
             COO.Folders = ifelse(str_detect(Mapping_df$File.Name, "ADMINISTRATIVE SUMMARY"), 1, 0))

DPR_Mapping <- read_excel(paste0(prod_path, "R Programming/Report Distribution File Org Automation/", 
                                 "Brainstorming/Mapping example.xlsx"))

DPR_Mapping <- DPR_Mapping[10:nrow(DPR_Mapping), ]

DPR_Mapping_join <- left_join(Mapping_df_pt2, DPR_Mapping) %>%
  na.omit() 

DPR_Mapping_join$`Destination Path` = paste0(final_output_folder, DPR_Mapping_join$`Short Folder`)
DPR_Mapping_join$`Full Destination Path` = paste0(final_output_folder, DPR_Mapping_join$`Short Folder`,
                                                  "/", DPR_Mapping_join$File.Name)

DPR_Mapping_join$`Full Destination Path` = gsub("\\\\", "/", DPR_Mapping_join$`Full Destination Path`)

System_COO_VP_remove_list <- paste(c(" CV ", " CPT ", " IP "), collapse = '|')
System_COO_VP <- DPR_Mapping_join %>%
  filter(Folder == "0") %>%
  filter(!str_detect(File.Name, System_COO_VP_remove_list))
  
Folder_Copy_df <- DPR_Mapping_join %>%
  filter(Folder == "1")

file.copy(from = System_COO_VP$zip_file_paths, to = System_COO_VP$`Full Destination Path`)


for (i in 1:nrow(Folder_Copy_df)) {
  source_files <- list.files(Folder_Copy_df$zip_file_paths[i], recursive = T)
  for (j in 1:length(source_files)) {
    file.copy(from = file.path(Folder_Copy_df$zip_file_paths[i], source_files[j]),
              to = file.path(Folder_Copy_df$`Destination Path`[i], source_files[j]),
              overwrite = T)
    
  }
}

#-------------MSMW CPT Reports---------------------------------------------

#-------------Renaming-----------------------------------------------------
Renaming_mapping <- read_xlsx(paste0(prod_path, "/R Programming/Report Distribution File Org Automation/",
                                     "Brainstorming/Renaming Mapping2.xlsx"))

Renaming_mapping$Old.Name.Path <- paste0(final_output_folder, Renaming_mapping$Old.Name, distribution, ".pdf")
Renaming_mapping$Old.Name.Path = gsub("\\\\", "/", Renaming_mapping$Old.Name.Path)
Renaming_mapping$Dir.Name <- dirname(Renaming_mapping$Old.Name.Path)
Renaming_mapping$New.Name.Path <- paste0(Renaming_mapping$Dir.Name, "/", Renaming_mapping$New.Name, ".pdf")

file.rename(Renaming_mapping$Old.Name.Path, Renaming_mapping$New.Name.Path)

#-------------Removing empty files------------------------------------------
output_files <- list.files(path = final_output_folder, recursive = TRUE, full.names = TRUE)
for (file in output_files) {
  if (file.size(file) < 1000) {
    file.remove(file)
    print(paste("Removed file:", file))
  }
}

final_output_files <- list.files(path = final_output_folder, recursive = TRUE, full.names = TRUE)
#------Quality Check (comparing output folder to previous month)-------------
previous_month_output <- rstudioapi::selectDirectory(caption = "select the previous months output folder")

previous_month_files <- basename(list.files(path = previous_month_output, recursive = TRUE))

current_month_files <- basename(final_output_files)

previous_names <- gsub(previous_distribution, "", previous_month_files)
current_names <- gsub(distribution, "", current_month_files)


# Get the files that are in previous month but not in current month
diff1 <- setdiff(previous_names, current_names)
# Get the files that are in current month but not in previous month
diff2 <- setdiff(current_names, previous_names)

# Print the differences
cat("Files in", previous_month_output, "but not in", final_output_folder, ":\n")
cat(paste(diff1, collapse = "\n"), "\n")

cat("Files in", final_output_folder, "but not in", previous_month_output, ":\n")
cat(paste(diff2, collapse = "\n"), "\n")

#-------Zipping files-----------------------------------------------------
subfolder_names_to_zip <- c("MSHS/Engineering/Department Reports_Engineering", "MSHS/Nursing/Site Administrative Reports_Nursing", 
                            "MSHS/Nursing/Site Department Reports_Nursing", "MSHS/PLamb/Site Administrative Reports",
                            "MSHS/PLamb/Site Department Reports", "MSHS/Radiology/Site Administrative Reports_Radiology",
                            "MSHS/Radiology/Site Department Reports_Radiology", "MSHS/Supply Chain and Support Services/Site Administrative Reports",
                            "MSHS/Supply Chain and Support Services/Site Department Reports", "MSHS/System COO/Site Department Reports",
                            "MSHS/System COO/Site VP Reports")

for (subfolder_name in subfolder_names_to_zip) {
  zip_file_name <- paste0(subfolder_name, ".zip")
  zip::zipr(zipfile = file.path(final_output_folder, zip_file_name), files = file.path(final_output_folder, subfolder_name), recurse = TRUE)
}
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
