# 1. Introduction ==============================================================
#' Converts a series of csv files to xlsx format for the ICASA DD.
#' 
print ("Starting the R script")
print(getwd())
# 2. Set working directory to script location. Define file name and paths.======

library(openxlsx2)
subfolder_for_csv <- "Data"

dictionary_sheets <- c("ReadMe", "Metadata", "Management_info", "Soils_data", "Weather_data", 
    "Measured_data", "Metadata_codes", "Management_codes", "Crop_codes",
    "Pest_codes", "Other_codes", "Glossary")        
  
# 3. Load the target wb and obtain sheet names =================================
print("Creating workbook")
# Create a new workbook
wb <- wb_workbook()

# 4. Loop through CSV files and add each as a worksheet ========================
for (file in dictionary_sheets) {
  # Read CSV into data frame
  csv_file <- paste0(file, ".csv")
  csv_file_path <- file.path(subfolder_for_csv, csv_file)
  df <- read.csv(csv_file_path, stringsAsFactors = FALSE)
  
  # Create a sheet name from the file name (remove folder and extension)
  # Add worksheet and write data
  wb$add_worksheet(sheet = file)
  wb$add_data(sheet = file, x = df, na.strings = "")
  
}

# 5. Add further formatting as needed. =========================================
# 5.1 Freeze at top line =======================================================
for (file in dictionary_sheets) {
  wb$freeze_pane(sheet = file, first_active_row = 2)  
}

# 5.2 Set column widths and text wrapping in ReadMe ============================

for (file in dictionary_sheets) {
  print(file)
  if (file == "ReadMe") {
    wb$set_col_widths(sheet = file, cols = c(1, 2), widths = c(90, 20))
    # Apply wrap style to the column
    # wb$set_row_heights(sheet = file, rows = 1:20, heights = 40)
    wb$add_cell_style(sheet=file, dims = "A1:A20", 
                      horizontal = "left", vertical = "top", wrap_text = TRUE)  
    wb$add_font(sheet = file, dims = "A1:A1", bold = TRUE)
    wb$add_font(sheet = file, dims = "A6:A6", bold = TRUE)
  } else {
    #wb$set_col_widths(sheet = file, cols = 1:22, widths = "auto")
    wb$add_font(sheet = file, dims = "A1:V1", bold = TRUE)
  }
}

# 5.2 Set column widths and text wrapping in sheets for data ===================
data_sheets <- c("Metadata", "Management_info", "Soils_data", "Weather_data", 
                  "Measured_data")
col_widths <- rep(12, 8)
for (file in data_sheets) {
  wb$set_col_widths(sheet = file, cols = c(1:18), widths = c(22,18,18,14,14,60,12,12,15,15,col_widths))
  wb$add_cell_style(sheet=file, dims = "F1:F100", 
                    horizontal = "left", vertical = "top", wrap_text = TRUE)  
  }

# 5.3 Set column widths and text wrapping in sheet for Metadata_codes ==========
file <- "Metadata_codes"
  wb$set_col_widths(sheet = file, cols = c(1:5), widths = c(20,12,60,40,15))
  wb$add_cell_style(sheet=file, dims = "C1:D100", 
                   horizontal = "left", vertical = "top", wrap_text = TRUE)  

# 5.4 Set column widths and text wrapping in sheet for Management_codes ========
file <- "Management_codes"
  wb$set_col_widths(sheet = file, cols = c(1:7), widths = c(12,12,12,40,20,40,15))
  wb$add_cell_style(sheet=file, dims = "D1:F600", 
                    horizontal = "left", vertical = "top", wrap_text = TRUE)  

# 5.5 Set column widths and text wrapping in sheet for Crop_codes ========
file <-"Crop_codes"
  wb$set_col_widths(sheet = file, cols = c(1:21), 
                    widths = c(20,12,30,30,30,30,30,12,12,12,15,12,12,12,12,12,12,12,12,40,25))
  wb$add_cell_style(sheet=file, dims = "G1:G200", 
                    horizontal = "left", vertical = "top", wrap_text = TRUE)  

# 5.6 Set column widths and text wrapping in sheet for Pest_codes ========
file <- "Pest_codes"
  wb$set_col_widths(sheet = file, cols = c(1:9), widths = c(12,10,18,40,12,12,30,12,12))
  wb$add_cell_style(sheet=file, dims = "D1:D100", 
                    horizontal = "left", vertical = "top", wrap_text = TRUE)  

# 5.7 Set column widths and text wrapping in sheet for Other_codes ========
file <- "Other_codes"
  wb$set_col_widths(sheet = file, cols = c(1:6), widths = c(15,12,12,40,30,15))
  wb$add_cell_style(sheet=file, dims = "D1:D100", 
                    horizontal = "left", vertical = "top", wrap_text = TRUE)  

# 5.7 Set column widths and text wrapping in sheet for the Glossary ========
  file <- "Glossary"
  wb$set_col_widths(sheet = file, cols = c(1:5), widths = c(20,20,60,12,12))
  wb$add_cell_style(sheet=file, dims = "C1:C200", 
                    horizontal = "left", vertical = "top", wrap_text = TRUE)  
  
# 6. Save the final, formatted ICASA DD workbook ===============================
# 6.1 Before saving select ReadMe and set cell A1 as active. ===================
wb$set_selected(sheet = 1)  # Necessary to avoid multiple sheets being selected.
wb$set_bookview(active_tab = 0, first_sheet = 0)

# 6.2 Save the final file
print("Saving workbook")
saved_xlsx <- file.path("ICASA_DATA_Dictionary.xlsx")
wb_save(wb, saved_xlsx)

# 5. End script with message giving location of the CSV files. =================
cat("XLSX files saved in", saved_xlsx, "\n")
