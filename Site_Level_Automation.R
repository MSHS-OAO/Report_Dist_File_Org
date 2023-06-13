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
    choices = c("MSBIB", "MSH", "MSM", "MSQ", "MSW"),
    title = "Select Output Site(s)",
    multiple = T,
    graphics = T,
    preselect = "MSBIB"
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
             MSinananvasishta = ifelse(str_detect(Mapping_df$zip_file_paths, "MSINANANVASISHTA"), 1, 0),
             TKnox = ifelse(str_detect(Mapping_df$zip_file_paths, "TKNOX"), 1, 0),
             JConnolly = ifelse(str_detect(Mapping_df$zip_file_paths, "JCONNOLLY"), 1, 0),
             LValentino = ifelse(str_detect(Mapping_df$zip_file_paths, "LVALENTINO"), 1, 0),
             MJablons = ifelse(str_detect(Mapping_df$zip_file_paths, "MJABLONS"), 1, 0),
             CGarcenot = ifelse(str_detect(Mapping_df$zip_file_paths, "CGARCENOT"), 1, 0),
             MOgnibene = ifelse(str_detect(Mapping_df$zip_file_paths, "MOGNIBENE"), 1, 0),
             SHonig = ifelse(str_detect(Mapping_df$zip_file_paths, "SHONIG"), 1, 0),
             VLeyko = ifelse(str_detect(Mapping_df$zip_file_paths, "VLEYKO"), 1, 0),
             DMazique = ifelse(str_detect(Mapping_df$zip_file_paths, "DMAZIQUE"), 1, 0),
             CGirdusky = ifelse(str_detect(Mapping_df$zip_file_paths, "CGIRDUSKY"), 1, 0),
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
             MSW = ifelse(str_detect(Mapping_df$File.Name, "MSW"), 1, 0),
             MSUS = ifelse(str_detect(Mapping_df$File.Name, "MSUS200") |
                             str_detect(Mapping_df$File.Name, "MSUS100"), 1, 0),
             CCW = ifelse(str_detect(Mapping_df$File.Name, "CCW"), 1, 0),
             MSH.IP.CV = ifelse(str_detect(Mapping_df$File.Name, "MSH02CV") |
                                  str_detect(Mapping_df$File.Name, "MSH03IP"), 1, 0),
             SmallFTE = ifelse(str_detect(Mapping_df$File.Name, 
                                          "SMALL REPORTING DEFINITIONS"), 1, 0))

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

Folder_Copy_df$test <- shortPathName(Folder_Copy_df$zip_file_paths)

for (g in 1:nrow(Folder_Copy_df)) {
  dir_copy(file.path(Folder_Copy_df$test[g]),
           Folder_Copy_df$`Destination Path`[g])
}

#-------------Renaming-----------------------------------------------------
#File renaming
APR_Renaming_mapping <- read_xlsx(paste0(prod_path,
                                         "/R Programming/",
                                         "Report Distribution File Org Automation/",
                                         "Mapping/Site Renaming.xlsx"), sheet = "Admin")

APR_Renaming_mapping$Old.Name.Path <- paste0(final_output_folder,
                                             APR_Renaming_mapping$Old.Name,
                                             distribution, ".pdf")
APR_Renaming_mapping$Old.Name.Path <- gsub("\\\\", "/",
                                           APR_Renaming_mapping$Old.Name.Path)
APR_Renaming_mapping$Dir.Name <- dirname(APR_Renaming_mapping$Old.Name.Path)
APR_Renaming_mapping$New.Name.Path <- paste0(APR_Renaming_mapping$Dir.Name, "/",
                                             APR_Renaming_mapping$New.Name, ".pdf")

file.rename(APR_Renaming_mapping$Old.Name.Path, APR_Renaming_mapping$New.Name.Path)

#Folder renaming
Folder_Renaming <- read_xlsx(paste0(prod_path,
                                    "/R Programming/",
                                    "Report Distribution File Org Automation/",
                                    "Mapping/Site Renaming.xlsx"), sheet = "Folder")

Folder_Renaming$Old.Name.Path <- paste0(final_output_folder, 
                                        Folder_Renaming$Old.Name)
Folder_Renaming$Dir.Name <- dirname(Folder_Renaming$Old.Name.Path)
Folder_Renaming$New.Name.Path <- paste0(Folder_Renaming$Dir.Name, "/", 
                                        Folder_Renaming$New.Name)
file.rename(Folder_Renaming$Old.Name.Path, Folder_Renaming$New.Name.Path)

#MSH renaming
MSH_Renaming <- read_xlsx(paste0(prod_path,
                                 "/R Programming/",
                                 "Report Distribution File Org Automation/",
                                 "Mapping/Site Renaming.xlsx"), sheet = "MSH")

MSH_Renaming$Old.Name.Path <- paste0(final_output_folder, MSH_Renaming$Old.Name)
MSH_Renaming$Dir.Name <- dirname(MSH_Renaming$Old.Name.Path)
MSH_Renaming$New.Name.Path <- paste0(MSH_Renaming$Dir.Name, "/",
                                     MSH_Renaming$New.Name, distribution)
file.rename(MSH_Renaming$Old.Name.Path, MSH_Renaming$New.Name.Path)


#-------------Removing empty files------------------------------------------
output_files <- list.files(path = final_output_folder, recursive = TRUE,
                           full.names = TRUE)

removed_files <- c()
for (file in output_files) {
  if (file.size(file) < 1000) {
    file.remove(file)
    removed_files <- c(removed_files, basename(file))
    print(paste("Removed file:", file))
  }
}

removed_files_df <- data.frame("Empty Files Removed" = removed_files)
write_xlsx(removed_files_df, 
           path = paste0(prod_path, 
                         "/R Programming/",
                         "Report Distribution File Org Automation/",
                         "Quality Checks/Site/", output_site, "/", 
                         "Empty Files Removed ", output_site, " ",
                         format(Sys.time(), '%d%b%y'),
                         ".xlsx"))

#-------Zipping files-----------------------------------------------------
MSBIB_zipfolders <- 
  c("MSBIB/MSB/CGarcenot Reports",
    "MSBIB/MSB/MOgnibene Reports",
    "MSBIB/MSB/MSB Department Reports",
    "MSBIB/MSB/MSB VP Reports",
    "MSBIB/MSB/SHonig Reports",
    "MSBIB/MSB/VLeyko Reports",
    "MSBIB/MSBI/CCastillo Reports",
    "MSBIB/MSBI/CGirdusky Reports",
    "MSBIB/MSBI/CMahoney Reports",
    "MSBIB/MSBI/DMazique Reports",
    "MSBIB/MSBI/ESellman Reports",
    "MSBIB/MSBI/JCoffin Reports",
    "MSBIB/MSBI/MSBI Department Reports",
    "MSBIB/MSBI/MSBI VP Reports",
    "MSBIB/MSUS/MSUS Department Reports",
    "MSBIB/MSUS2/630571_CCW_RAD CCW RADIOLOGY",
    "MSBIB/MSUS2/630571_MSUS200 MSUS PERIOP")

MSH_zipfolders <-
  c(paste0("MSH/MSH_Admin Reports_", distribution),
    paste0("MSH/MSH_BOliver Departments_", distribution),
    paste0("MSH/MSH_Department Reports_", distribution),
    paste0("MSH/MSH_FCartwright Departments_", distribution),
    paste0("MSH/MSH_MMcCarry Departments_", distribution))

MSQ_zipfolders <-
  c("MSQ/CHernandez Reports",
    "MSQ/EBabar Reports",
    "MSQ/JGoldstein Reports",
    "MSQ/MSQ VP Reports",
    "MSQ/MSQ Department Reports")

MSM_zipfolders <-
  c("MSM/BRabill Departments",
    "MSM/CDejesus Departments",
    "MSM/LBorenstein Departments",
    "MSM/MSinananvasishta Departments",
    "MSM/MSM Department Reports",
    "MSM/TKnox Departments")

MSW_zipfolders <-
  c("MSW/JConnolly Departments",
    "MSW/LBorenstein Departments",
    "MSW/LValentino Departments",
    "MSW/MJablons Departments")

if (output_site == "MSBIB") {
  for (subfolder_name in MSBIB_zipfolders) {
    zip_file_name <- paste0(subfolder_name, ".zip")
    zip::zipr(zipfile = file.path(final_output_folder, zip_file_name),
              files = file.path(final_output_folder, subfolder_name),
              recurse = TRUE)
  }
} else if (output_site == "MSH") {
  for (subfolder_name in MSH_zipfolders) {
    zip_file_name <- paste0(subfolder_name, ".zip")
    zip::zipr(zipfile = file.path(final_output_folder, zip_file_name),
              files = file.path(final_output_folder, subfolder_name),
              recurse = TRUE)
  }
} else if (output_site == "MSM") {
  for (subfolder_name in MSM_zipfolders) {
    zip_file_name <- paste0(subfolder_name, ".zip")
    zip::zipr(zipfile = file.path(final_output_folder, zip_file_name),
              files = file.path(final_output_folder, subfolder_name),
              recurse = TRUE)
  }
} else if (output_site == "MSW") {
  for (subfolder_name in MSW_zipfolders) {
    zip_file_name <- paste0(subfolder_name, ".zip")
    zip::zipr(zipfile = file.path(final_output_folder, zip_file_name),
              files = file.path(final_output_folder, subfolder_name),
              recurse = TRUE)
  }
} else if (output_site == "MSQ") {
  for (subfolder_name in MSQ_zipfolders) {
    zip_file_name <- paste0(subfolder_name, ".zip")
    zip::zipr(zipfile = file.path(final_output_folder, zip_file_name),
              files = file.path(final_output_folder, subfolder_name),
              recurse = TRUE)
  }
}

#------Quality Check (comparing output folder to previous month)-------------
final_output_files <- list.files(path = paste0(final_output_folder, "/",
                                               output_site),
                                 recursive = TRUE, full.names = TRUE)

previous_month_output <-
  rstudioapi::selectDirectory(caption =
                                "select the previous months output folder")

previous_month_files <- list.files(path = previous_month_output,
                                   recursive = TRUE, full.names = TRUE)

previous_names <- gsub(previous_distribution, "", previous_month_files)
current_names <- gsub(distribution, "", final_output_files)

# Get the files that are in previous month but not in current month
diff1 <- setdiff(basename(previous_names), basename(current_names))

full_paths_only_in_previous <- 
  previous_month_files[which(basename(previous_names) %in% diff1)]
cat("Files in", previous_month_output, "but not in", final_output_folder, ":\n")
previous_distribution_only <- data.frame(File.Names = 
                                           basename(full_paths_only_in_previous), 
                                         File.Paths = full_paths_only_in_previous)
write_xlsx(previous_distribution_only, 
           path = paste0(prod_path, 
                         "/R Programming/",
                         "Report Distribution File Org Automation/",
                         "Quality Checks/Site/", output_site, "/",
                          output_site, " Reports in Previous Distribution Only ",
                         format(Sys.time(), '%d%b%y'), ".xlsx"))
# Get the files that are in current month but not in previous month
diff2 <- setdiff(basename(current_names), basename(previous_names))

full_paths_only_in_current <- 
  final_output_files[which(basename(current_names) %in% diff2)]
cat("Files in", final_output_folder, "but not in", previous_month_output, ":\n")
current_distribution_only <- data.frame(File.Names = 
                                          basename(full_paths_only_in_current), 
                                        File.Paths = full_paths_only_in_current)

write_xlsx(current_distribution_only, 
           path = paste0(prod_path, 
                         "/R Programming/",
                         "Report Distribution File Org Automation/",
                         "Quality Checks/Site/", output_site, "/",
                         output_site, " Reports in Current Distribution Only ",
                         format(Sys.time(), '%d%b%y'), ".xlsx"))
# Script End --------------------------------------------------------------

