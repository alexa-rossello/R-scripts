# ---- Import Data Aggregation ----
# Purpose: This script aggregates import data and saves the data into a spreadsheet. 
# Author: Lex Rossello

# Libraries -----
library(odbc)
library(DBI)
library(lubridate)
library(dplyr)
library(openxlsx)

# Query data from database -----

# Set current week variable
# Date is automatically set to 'current week' after Friday ends
current.week <- Sys.Date() - wday(Sys.Date() + 1)

### Option to manually set week
#current.week <- as.Date("m/d/yy", "%m/%d/%y")

# Connect to SQL server 
# Note: Database name has been anonymized to maintain confidentiality.
con <- dbConnect(odbc(), "DB")

# Query data from database
# Note: Field and table names have been anonymized or removed to maintain confidentiality.
imports <- dbGetQuery(con, paste("SELECT fields FROM table WHERE WEEK_ENDING='",current.week,"'",sep=""))

# Disconnect from database
dbDisconnect(con)

# Reset row index
rownames(imports) <- 1:nrow(imports)

# Aggregate data -----

# Aggregate data by Product Code and Region
imports.agg <- aggregate(Volume ~ Product + Region, data = imports, FUN = sum, na.rm = TRUE)

# Reshape data frame to wide structure
imports.agg <- reshape(imports.agg, idvar = "Region", timevar = "PRODUCT_CODE", direction = "wide")

# Change NAs to 0
imports.agg[is.na(imports.agg)] <- 0

# Reset row index
rownames(imports.agg) <- 1:nrow(imports.agg)

# Add columns of zeros for all products
# Note: Field names have been anonymized or removed to maintain confidentiality.
x <- c("VOLUME.10","VOLUME.11","VOLUME.12","VOLUME.13","VOLUME.14",
       "VOLUME.15","VOLUME.16","VOLUME.17","VOLUME.18","VOLUME.19",
       "VOLUME.20","VOLUME.21","VOLUME.22","VOLUME.23","VOLUME.24",
       "VOLUME.25","VOLUME.26","VOLUME.27","VOLUME.28","VOLUME.29")
imports.agg[x[!(x %in% colnames(imports.agg))]] = 0

# Combine Regions 2-4
imports.agg <- imports.agg %>%
  mutate(Region = recode(Region, '2 ' = '2-4', '3 ' = '2-4', '4 ' = '2-4'))

imports.agg <- imports.agg %>% group_by(Region) %>% 
  summarize(across(everything(), sum),
            .groups = 'drop')  %>%
  as.data.frame()

# Remove certain products
# Note: Field names have been anonymized or removed to maintain confidentiality.
imports.agg <- subset(imports.agg, select=-c(VOLUME.12, VOLUME.16, VOLUME.19, VOLUME.25, VOLUME.28))

# Add rows of zero for each Region, remove if that Region exists
# Blank data frame
Region.blank <- data.frame(Region = c('1X','1Y','1Z','2-4','5 '), #do not remove space from the 5
                         "VOLUME.10"=0,"VOLUME.11"=0,"VOLUME.13"=0,"VOLUME.14"=0,"VOLUME.15"=0,
                         "VOLUME.17"=0,"VOLUME.18"=0,"VOLUME.20"=0,"VOLUME.21"=0,"VOLUME.22"=0,
                         "VOLUME.23"=0,"VOLUME.24"=0,"VOLUME.26"=0,"VOLUME.27"=0,"VOLUME.29"=0)

imports.agg <- rbind(imports.agg, Region.blank)

# Add totals to rows
imports.agg <- imports.agg %>%
  mutate(Total = rowSums(.[-1]))

# Remove rows with zeros based on Region if it is a duplicate
imports.agg <- imports.agg[order(imports.agg[,'Region'],-imports.agg[,'Total']),]
imports.agg <- imports.agg[!duplicated(imports.agg$Region),]

# Calculate totals for parent products
# Calculate total product 100
imports.agg$VOLUME.134 <- rowSums(imports.agg[, c("VOLUME.10", "VOLUME.11", "VOLUME.13", "VOLUME.14")])

# Calculate total product 200
imports.agg$VOLUME.152 <- rowSums(imports.agg[, c("VOLUME.15", "VOLUME.17")])

# Calculate total product 300
imports.agg$VOLUME.461 <- rowSums(imports.agg[, c("VOLUME.18", "VOLUME.20")])

# Calculate total product 400
imports.agg$VOLUME.462 <- rowSums(imports.agg[, c("VOLUME.21", "VOLUME.22")])

# Select columns with products needed for estimation
imports.agg <- subset(imports.agg, select = c('Region','VOLUME.100','VOLUME.200','VOLUME.300',
                                              'VOLUME.400','VOLUME.23','VOLUME.24',
                                              'VOLUME.26','VOLUME.27','VOLUME.29'))

# Re-calculate the row totals
imports.agg <- imports.agg %>%
  mutate(Total = rowSums(.[-1]))

# Rename Region 6 to Total
imports.agg["Region"][imports.agg["Region"] == '6 '] <- "Total"

# Rename columns
names(imports.agg) <- sub('^VOLUME.', 'Product ', names(imports.agg))

# Reset row index
rownames(imports.agg) <- 1:nrow(imports.agg)

# Add week ending column
imports.agg$Week_Ending <- as.Date(current.week)

# Reorder columns
imports.agg <- imports.agg[, c(12,1,2,3,4,5,6,7,8,9,10,11)]

# Export results -----

# Set directory 
# Note: Directory has been anonymized to maintain confidentiality.
mainDir <- "REMOVED"
subDir <- paste(current.week)
dir.create(file.path(mainDir, subDir), showWarnings = FALSE)
setwd(file.path(mainDir, subDir))

#getwd()

# Set file name
current_datetime <- format(Sys.time(), "%H-%M")

print(current_datetime)

file.name <- paste0("IMP_DIFZ ", current.week, " run_", current_datetime, ".xlsx")

# Set header style
hs <- createStyle(textDecoration = "BOLD")

# Export to Excel only if the file does not exist
# If the file exists, warning message appears - rename the existing file
write.file <- if(file.exists(file.name)) {
    message(paste(file.name, "already exists! Rename file before exporting results."))
    } else {
      openxlsx::write.xlsx(imports.agg, file = file.name, colNames = TRUE, headerStyle = hs)
    }
