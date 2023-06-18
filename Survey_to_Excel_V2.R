library(tidyverse)
library(openxlsx)
library(scales)

#--------------------------
#If the standardized residual is greater than 2 or less than -2 in absolute value, then the cell is considered to be significantly contributing to the chi-squared statistic at a 5% level of significance.
#The standardized residual represents the number of standard deviations that the observed frequency deviates from the expected frequency
#--------------------------
setwd("C:/Users/JoeKea/OneDrive - NZ Transport Agency/R/Survey")

rm(list = ls())

#--------------------------

#--------------------------
setClass("Sheet_Design",
         slots = list(
           data = "data.frame",
           rowVars = "list",
           colVars = "list"
         )
)

#--------------------------
sheet_designs <- list(
  new("Sheet_Design", data = mtcars, rowVars = list("gear"), colVars = list("vs", "am", "carb")),
  new("Sheet_Design", data = starwars, rowVars = list("eye_color"), colVars = list("homeworld", "gender"))
)
#--------------------------
Workbook_Maker <- function(){
  
  workbook <- list()
  for (sheet_design in sheet_designs) {
    sheet <- Sheet_Maker(sheet_design)
    workbook <- c(workbook,  list(sheet))
  }

  return(workbook)
}

#--------------------------
Sheet_Maker <- function(sheet_design){
  
  tables <- list()
  for (rowVar in sheet_design@rowVars) {
    table <- Table_Maker(sheet_design, rowVar)
    tables <- c(tables, list(table))
  }
  return(tables)
  
}
#--------------------------
Table_Maker <- function(sheet_design, current_rowVar){
  #have the first two columns be Overall and Net
  
  data <- sheet_design@data
  
  #Overall
  overallTable <- as.data.frame(prop.table(table(data[[current_rowVar]])) * 100)
  overallTable$Freq <- round(overallTable$Freq)
  colnames(overallTable) <- c("Column %", "Overall")
  overallTable <- format(overallTable, justify = "right")
  overallTable <- apply(overallTable, 2, trimws)
  
  #Net
  netData <- data[complete.cases(data[, unlist(sheet_design@colVars)]), ]
  netTable <- as.data.frame(prop.table(table(netData[[current_rowVar]])) * 100)
  netTable$Freq <- round(netTable$Freq)
  colnames(netTable) <- c("Column %", "Net")
  netTable <- format(netTable, justify = "right")
  netTable <- apply(netTable, 2, trimws)
  
  table <- merge(overallTable, netTable, by = "Column %", all = T)
  
  print(table)
  
  for (colVar in sheet_design@colVars) {
    
    subtable <- Subtable_Maker(sheet_design, current_rowVar, colVar)
    table <- merge(table, subtable, by = "Column %", all = T)
  }
  
  return(table)
  
}
#--------------------------
Subtable_Maker <- function(sheet_design, current_rowVar, current_colVar){ 
  
  data <- sheet_design@data

  subtable <- xtabs(~data[[current_rowVar]]+data[[current_colVar]],data=data)
  
  chisq <- chisq.test(subtable)
  std_res <- chisq$stdres
  expected <- chisq$expected
  subtable <- prop.table(subtable, margin = 2)
  subtable <- apply(subtable, 2, round, digits = 2)
  subtable <- subtable * 100
  
  subtable <- format(subtable, justify = "right")
  subtable <- apply(subtable, 2, trimws)
  
  upIndices <- std_res > 1.96 & std_res != "NAN" & expected > 4
  downIndices <- std_res < -1.96 & std_res != "NAN" & expected > 4
  subtable[upIndices] <- paste(subtable[upIndices], intToUtf8(8593))
  subtable[downIndices] <- paste(subtable[downIndices], intToUtf8(8593))

  subtable <- as.data.frame(subtable)
  
  colnames(subtable) <- paste(current_colVar, colnames(subtable), sep = '.')
  
  subtable <- rownames_to_column(subtable, "Column %")
  
  return(subtable)
}
#--------------------------
#--------------------------

workbook <- Workbook_Maker()

df <- data.frame(A = c(1, 2, NA, 4),
                 B = c(NA, 5, 6, 7),
                 C = c(8, 9, 10, NA),
                 D = c(11, NA, 13, 14))

# Specify the three specific columns
specific_columns <- c("A", "B", "C")

# Extract rows without missing values in specific columns
filtered_df <- df[complete.cases(df[, specific_columns]), ]

# Print the filtered dataframe
print(filtered_df)



