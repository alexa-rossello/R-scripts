# ---- Missing Report List ----
# Purpose: This script identifies missing reports and saves the list into a spreadsheet by report number. 
# Author: Lex Rossello

# Libraries -----
library(odbc)
library(DBI)
library(lubridate)
library(dplyr)
library(tidyr)
library(stringr)
library(openxlsx)

# Query data from database -----
# Set current week variable
# Date is automatically set to 'current week' after Friday ends
current.week <- Sys.Date() - wday(Sys.Date() + 1)
prev.week <- current.week - 7

###option to manually set weeks
#current.week <- as.Date("m/d/yy", "%m/%d/%Y")
#prev.week <- current.week - 7

# Connect to SQL server 
con <- dbConnect(odbc(), "db")

# Queries data from database
# Note: Field and table names have been anonymized or removed to maintain confidentiality.
data_current <- dbGetQuery(con, paste("SELECT fields FROM db.table1 WHERE week ='", current.week,"'",sep=""))
data_prev <- dbGetQuery(con, paste("SELECT fields FROM db.table1 WHERE week='", prev.week,"'",sep=""))
company_name <-  dbGetQuery(con, paste("SELECT fields FROM db.table2", sep=""))

# Disconnect from db
dbDisconnect(con)

# Restructure data frames -----

# Rename columns in data frames
data_current <- data_current %>% rename(Week_current = week, type_current = type)

data_prev <- data_prev %>% rename(week_prev = week, type_prev = type)

# Merge current and previous form_ent data frames by comp_ID
missing_forms <- merge(x=data_prev, y=data_current, by="comp_ID", all.x=TRUE)

# Drop extra form_ID column and prev week date
missing_forms <- subset(missing_forms, select = -c(form_ID.y,week_prev))

# Rename form_ID.x to form_ID
missing_forms <- missing_forms %>% rename(form_ID = form_ID.x)

# Modify Date type of week_current
missing_forms$week_current <- as.Date(missing_forms$week_current, 'UTC')

# Add company names from company_name data frame
missing_forms2 <- merge(x=missing_forms, y=company_name, by="comp_ID", all.x=TRUE)

# Trim whitespace in data frame
missing_forms2 <- missing_forms2 %>% mutate(across(where(is.character), str_trim))

# Add form IDs
# Note: Field names have been anonymized or removed to maintain confidentiality.
missing_forms2 <- missing_forms2 %>% mutate(report_number = case_when(
  form_ID == "A" ~ 10,
  form_ID == "B" ~ 20,
  form_ID == "C" ~ 30,
  form_ID == "D" ~ 40,
  form_ID == "E" ~ 50
))

# Reorder columns
final_missing <- missing_forms2[, c(4,2,6,1,7,3,5)]

# Sort by form number then company name
final_missing <- final_missing %>% arrange((report_number), company_name)

# Identify missing reports -----

# Replace NA with 'missing'
final_missing$type_current <- final_missing$type_current %>% replace_na("missing")

# Drop if report received
final_missing <- final_missing %>% filter(type_current == "missing")

# Remove exclusion list -----

# Set directory 
# Note: Directory has been anonymized to maintain confidentiality.
setwd("REMOVED")
exclude <- read.csv("file.csv", header = TRUE)
excl_list <- as.list(exclude$comp_ID)
final_missing_excl <- subset(final_missing, !(comp_ID %in% excl_list))

# Subset missing by form -----

A <- final_missing_excl %>% filter(form_ID == "A")
B <- final_missing_excl %>% filter(form_ID == "B")
C <- final_missing_excl %>% filter(form_ID == "C")
D <- final_missing_excl %>% filter(form_ID == "D")
E <- final_missing_excl %>% filter(form_ID == "E")

# Sort data -----

final_missing_excl <- final_missing_excl[order(final_missing_excl$report_number, final_missing_excl$COMPANY),]
A <- A[order(A$report_number, A$COMPANY),]
B <- B[order(B$report_number, B$COMPANY),]
C <- C[order(C$report_number, C$COMPANY),]
D <- D[order(D$report_number, D$COMPANY),]
E <- E[order(E$report_number, E$COMPANY),]

# Export results -----

# Set directory 
# Note: Directory has been anonymized to maintain confidentiality.
mainDir <- "Report"
subDir <- paste(current.week)
dir.create(file.path(mainDir, subDir), showWarnings = FALSE)
setwd(file.path(mainDir, subDir))

#getwd()

# Set file name
current_datetime <- format(Sys.time(), "%m-%d %H-%M")

print(current_datetime)

file.name <- paste0("Missing ", current.week, " run_", current_datetime, ".xlsx")

print(file.name)

# Set header style
hs <- createStyle(textDecoration = "BOLD")

# List of data frames
df_list <- list('All Missing' = final_missing_excl, 'A' = A, 'B' = B, 'C' = C, 'D' = D, 'E' = E)

# Export to Excel only if the file does not exist
# If the file exists, warning message appears - rename the existing file
write.file <- if(file.exists(file.name)) {
  message(paste(file.name, "already exists! Rename file before exporting results."))
} else {
    openxlsx::write.xlsx(df_list, file = file.name, colNames = TRUE, 
                         colWidths = "auto", headerStyle = hs) 
}
