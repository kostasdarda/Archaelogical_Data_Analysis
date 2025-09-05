##### THIS SCRIPT IS USED TO TRANSFORM THE "WorkingDataSheets" FILE INTO A 
##### USABLE FORMAT. THE OUTPUT OF THE SCRIPT IS A NEW FILE CALLED
##### "Organized_WorkingDataSheets"
      


library(readxl)
library(dplyr)
library(openxlsx)
# library(stringr)

file_path <- "./WorkingDataSheets.xlsx"
raw_data <- read_excel(file_path, sheet = 1)      # I am reading the RawData sheet

# I am creating a working copy, so that the original is not modified
cleaned_data <- raw_data

# I am defining the target genera and associated scientific names 
# (used for pattern matching column names)
# Each genus gets a list of possible scientific names to look for in column headers
target_patterns <- list(
  "Alces" = c("Alces", "Alces alces"),
  "Bison" = c("Bison", "Bison bison", "Bison priscus"),
  "Bos" = c("Bos", "Bos taurus", "Bos primigenius", "Bibos", "Bibos gaurus"),
  "Canis" = c("Canis", "Canis lupus", "Canis latrans", "Canis familiaris", "Canis aureus"),
  "Capra" = c("Capra", "Capra hircus", "Capra ibex"),
  "Capreolus" = c("Capreolus", "Capreolus capreolus"),
  "Cervus" = c("Cervus", "Cervus elaphus", "Cervus canadensis"),
  "Coelodonta" = c("Coelodonta", "Coelodonta antiquitatis"),
  "Dama" = c("Dama", "Dama dama"),
  "Equus" = c("Equus", "Equus ferus", "Equus caballus", "Equus hydruntinus", "Equus hemionus"),
  "Mammuthus" = c("Mammuthus", "Mammuthus primigenius"),
  "Megaloceros" = c("Megaloceros", "Megaloceros giganteus"),
  "Ovibos" = c("Ovibos", "Ovibos moschatus"),
  "Ovis" = c("Ovis", "Ovis aries", "Ovis ammon"),
  "Palaeoloxodon" = c("Palaeoloxodon", "Palaeoloxodon antiquus"),
  "Rangifer" = c("Rangifer", "Rangifer tarandus"),
  "Saiga" = c("Saiga", "Saiga tatarica"),
  "Stephanorhinus" = c("Stephanorhinus", "Stephanorhinus hemitoechus", "Stephanorhinus kirchbergensis"),
  "Sus" = c("Sus", "Sus scrofa", "Sus domesticus"),
  "Ursus" = c("Ursus", "Ursus arctos", "Ursus spelaeus"),
  "Panthera" = c("Panthera", "Panthera leo", "Panthera spelaea", "Panthera pardus")
)

# I loop through each genus to identify and sum the related columns
for (genus in names(target_patterns)) {
  # Finds all column names that contain any of the genus-related terms
  match_cols <- grep(
    paste(target_patterns[[genus]], collapse = "|"), # combines the patterns with "or"
    names(cleaned_data),
    ignore.case = TRUE,   # ignores capitalization
    value = TRUE
  )
  
  # If a matching column is found
  if (length(match_cols) > 0) {
    # Adding in a line that just prints match_cols so that we can double check what species are being IDd in each genus. 
    print(match_cols)
    # It replaces "–" with 0 and converts values to numeric
    cleaned_data[match_cols] <- lapply(cleaned_data[match_cols], function(x) {
      x <- ifelse(x == "–", 0, x)
      as.numeric(x)
    })
    # It creates new column with the sum of all related species for each one of the target genera
    cleaned_data[[paste0(genus, "_NISP")]] <- rowSums(select(cleaned_data, all_of(match_cols)), na.rm = TRUE)
  } else {
    # Just added something extra if no matching columns are found for a genus
    message(paste("No columns matched for genus:", genus))
  }
}

# I distinguish columns 1 to 16, which contain ID, Locality, etc. with the new cols
non_zooarch_cols <- names(raw_data)[1:16]
# nisp_cols are all new NISP summary columns created during the loop
nisp_cols <- grep("_NISP$", names(cleaned_data), value = TRUE)

# I create the final cleaned dataset (cols 1:16 + new cols)
final_cleaned <- cleaned_data %>%
  select(all_of(non_zooarch_cols), all_of(nisp_cols))

# I remove the second row (ID:1), which is unnecessary for the CleanedData
final_cleaned <- final_cleaned[-1, ]

# I load the full workbook, in order to modify it
wb <- loadWorkbook(file_path)

#I did not use this after all 
#sheets <- excel_sheets(file_path)

# I replace the existing "CleanedData" with the new one
removeWorksheet(wb, "CleanedData")
addWorksheet(wb, "CleanedData")
writeData(wb, "CleanedData", final_cleaned)

# Auto-adjust column widths for this new sheet (not necessary)
setColWidths(wb, "CleanedData", cols = 1:ncol(final_cleaned), widths = "auto")

# I save the modified workbook to a new file
saveWorkbook(wb, "./Organized_WorkingDataSheets.xlsx", overwrite = TRUE)




























