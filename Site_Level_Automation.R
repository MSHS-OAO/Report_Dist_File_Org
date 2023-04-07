# Libraries ---------------------------------------------------------------
library(tidyverse)
library(rstudioapi)
library(zip)
library(readxl)
library(zoo)
library(writexl)
library(fs)

# Assigning Directory(ies) ------------------------------------------------
prod_path <- paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/",
                    "Productivity/")

#User selected directory where output files could be sent
unzip_output_folder <-
  rstudioapi::selectDirectory(caption = "select unzip output folder")

final_output_folder <-
  rstudioapi::selectDirectory(caption = "select final output folder")

# Data Import -------------------------------------------------------------
#Selecting and unzipping batch download file to get file paths
#File must be zip file containing all the APR and DPR files for a site
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

#Site selection------------------------------------------------------------
output_site <-
  select.list(
    choices = c("MSHS", "MSB", "MSBI", "MSH", "MSM", "MSQ", "MSW"),
    title = "Select Output Site(s)",
    multiple = T,
    graphics = T,
    preselect = "MSHS"
  )

# Data Import -------------------------------------------------------------
dates <- read_xlsx(paste0(prod_path, "Universal Data/Mapping/",
                          "MSHS_Pay_Cycle.xlsx"))

#Table of distribution dates
dist_dates <- dates %>%
  select(END.DATE, PREMIER.DISTRIBUTION) %>%
  distinct() %>%
  drop_na() %>%
  arrange(END.DATE) %>%
  #filter only on distribution end dates
  filter(PREMIER.DISTRIBUTION %in% c(TRUE, 1),
         END.DATE < as.POSIXct(Sys.Date() - 21))
#Table of non-distribution dates
non_dist_dates <- dates %>%
  select(END.DATE, PREMIER.DISTRIBUTION) %>%
  distinct() %>%
  drop_na() %>%
  arrange(END.DATE) %>%
  #filter only on distribution end dates
  filter(PREMIER.DISTRIBUTION %in% c(FALSE, 0),
         END.DATE < as.POSIXct(Sys.Date() - 21))
#Selecting current and previous distribution dates
distribution <- format(dist_dates$END.DATE[nrow(dist_dates)], "%m.%d.%y")
previous_distribution <-
  format(dist_dates$END.DATE[nrow(dist_dates) - 1], "%m.%d.%y")
previous_cpt_distribution <-
  format(dist_dates$END.DATE[nrow(dist_dates) - 2], "%m.%d.%y")
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
                                format(sort.POSIXlt
                                       (dist_dates$END.DATE,
                                         decreasing = T),
                                       "%m.%d.%y"),
                              multiple = F,
                              title = "Select current distribution",
                              graphics = T)
  which(distribution == format(dist_dates$END.DATE, "%m.%d.%y"))
  previous_distribution <-
    format(dist_dates$END.DATE[which(distribution ==
                                       format(dist_dates$END.DATE,
                                              "%m.%d.%y")) - 1], "%m.%d.%y")
}

# Data References ---------------------------------------------------------

#Copying folders to output folder
file.copy(from = paste0(prod_path, 
"/R Programming/Report Distribution File Org Automation/Site Folder Structure/",
output_site),
          to = final_output_folder, recursive = T)

Folders <-
  list.dirs(
    path = paste0(prod_path, 
                  "R Programming/Report Distribution File Org Automation/",
                  "Site Folder Structure/",
                  output_site), recursive = FALSE)

#Creating first mapping dataframe 
Mapping_df <- data.frame("Source Paths" = zip_file_paths)
Mapping_df <- Mapping_df %>%
  data.frame("File Name" = basename(Mapping_df$zip_file_paths))
Mapping_df <- Mapping_df %>%
  data.frame(APR = ifelse(str_detect(Mapping_df$zip_file_paths, "_APR_"), 1, 0),
             DPR = ifelse(str_detect(Mapping_df$zip_file_paths, "_DPR_"), 1, 0),
             Folder = ifelse(str_detect(Mapping_df$zip_file_paths, ".pdf"), 0, 1),
             BRabill = ifelse(str_detect(Mapping_df$zip_file_paths, "BRABILL"), 1, 0),
             CDejesus = ifelse(str_detect(Mapping_df$zip_file_paths, "CDEJESUS"), 1, 0),
             LBorenstein = ifelse(str_detect(Mapping_df$zip_file_paths, "LBORENSTEIN"), 1, 0),
             MSinanvasishta = ifelse(str_detect(Mapping_df$zip_file_paths, "MSINANAVASISHTA"), 1, 0),
             TKnox = ifelse(str_detect(Mapping_df$zip_file_paths, "TKNOX"), 1, 0),
             JConnolly = ifelse(str_detect(Mapping_df$zip_file_paths, "JCONNOLLY"), 1, 0),
             LValentino = ifelse(str_detect(Mapping_df$zip_file_paths, "LVALENTINO"), 1, 0),
             MJablons = ifelse(str_detect(Mapping_df$zip_file_paths, "MJABLONS"), 1, 0),
             CGarcenot = ifelse(str_detect(Mapping_df$zip_file_paths, "CGARCENOT"), 1, 0),
             MOgnibene = ifelse(str_detect(Mapping_df$zip_file_paths, "MOGNIBEBE"), 1, 0),
             SHonig = ifelse(str_detect(Mapping_df$zip_file_paths, "SHONIG"), 1, 0),
             VLeyko = ifelse(str_detect(Mapping_df$zip_file_paths, "VLEYKO"), 1, 0),
             BBarnett = ifelse(str_detect(Mapping_df$zip_file_paths, "BBARNETT"), 1, 0),
             CBerner = ifelse(str_detect(Mapping_df$zip_file_paths, "CBERNER"), 1, 0),
             CCastillo = ifelse(str_detect(Mapping_df$zip_file_paths, "CCASTILLO"), 1, 0),
             CMahoney = ifelse(str_detect(Mapping_df$zip_file_paths, "CMAHONEY"), 1, 0),
             ESellman = ifelse(str_detect(Mapping_df$zip_file_paths, "ESELLMAN"), 1, 0),
             JCoffin = ifelse(str_detect(Mapping_df$zip_file_paths, "JCOFFIN"), 1, 0),
             MMcCarry = ifelse(str_detect(Mapping_df$zip_file_paths, "MMCCARRY"), 1, 0),
             BOliver = ifelse(str_detect(Mapping_df$zip_file_paths, "BOLIVER"), 1, 0),
             FCartwright = ifelse(str_detect(Mapping_df$zip_file_paths, "FCARTWRIGHT"), 1, 0),
             CHernandez = ifelse(str_detect(Mapping_df$zip_file_paths, "CHERNANDEZ"), 1, 0),
             EBabar = ifelse(str_detect(Mapping_df$zip_file_paths, "EBABAR"), 1, 0),
             JGoldstein = ifelse(str_detect(Mapping_df$zip_file_paths, "JGOLDSTEIN"), 1, 0),
             AdminSummary = ifelse(str_detect(Mapping_df$File.Name,
                                             "ADMINISTRATIVE SUMMARY"), 1, 0),
             MSB = ifelse(!str_detect(Mapping_df$File.Name, "MSBI") &
                            str_detect(Mapping_df$File.Name, "MSB"), 1, 0),
             MSBI = ifelse(str_detect(Mapping_df$File.Name, "MSBI"), 1, 0),
             MSH = ifelse(str_detect(Mapping_df$File.Name, "MSH"), 1, 0),
             MSM = ifelse(str_detect(Mapping_df$File.Name, "MSM"), 1, 0),
             MSQ = ifelse(str_detect(Mapping_df$File.Name, "MSQ"), 1, 0),
             MSW = ifelse(str_detect(Mapping_df$File.Name, "MSW"), 1, 0))
        
Site_Mapping <- read_excel(paste0(prod_path,
                                 "R Programming/Report Distribution File Org Automation/",
                                 "Mapping/Site Mapping.xlsx"))

Mapping_join <- left_join(Mapping_df, Site_Mapping) %>%
  na.omit()

#Removing any DNU reports
Mapping_join <- Mapping_join[!grepl("DO NOT USE", Mapping_join$File.Name), ]

Mapping_join$`Destination Path` <- paste0(final_output_folder,
                                           Mapping_join$`Short Folder`, "/",
                                           Mapping_join$File.Name)
#Copying over pdf files
File_Copy_df <- Mapping_join %>%
  filter(Folder == "0")

file.copy(from = File_Copy_df$zip_file_paths,
          to = File_Copy_df$`Destination Path`)

#Copying over folders and contents
Folder_Copy_df <- Mapping_join %>%
  filter(Folder == "1")

for (g in 1:nrow(Folder_Copy_df)) {
  dir_copy(file.path(Folder_Copy_df$zip_file_paths[g]),
                     Folder_Copy_df$`Destination Path`[g])
}

#RENAMING

