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
setClass("Table_Group",
         slots = list(
           title = "character",
           table = "data.frame",
           footer = "character"
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
  for (sheetDesign in sheet_designs) {
    sheet <- Sheet_Maker(sheetDesign)
    workbook <- c(workbook,  sheet)
  }

  return(workbook)
}

#--------------------------
Sheet_Maker <- function(sheetDesign){
  
  tables <- list()
  for (rowVar in sheetDesign@rowVars) {
    title <- rowVar
    table <- Table_Maker(sheetDesign, rowVar)
    footer <- Footer_Maker(sheetDesign, rowVar)
      
    tableGroup <- new("Table_Group", title = title, table = table, footer = footer)
    tables <- c(tables, tableGroup)
  }
  return(tables)
  
}
#--------------------------
Table_Maker <- function(sheetDesign, current_rowVar){
  
  data <- sheetDesign@data
  data <- data[complete.cases(data[, unlist(sheetDesign@colVars)]), ]
  
  table <- as.data.frame(prop.table(table(data[[current_rowVar]])) * 100)
  table$Freq <- round(table$Freq)
  colnames(table) <- c("Column %", "Net")
  table <- format(table, justify = "right")
  table <- apply(table, 2, trimws)
  
  for (colVar in sheetDesign@colVars) {
    
    subtable <- Subtable_Maker(sheetDesign, current_rowVar, colVar)
    table <- merge(table, subtable, by = "Column %", all = T)
  }
  
  table$Net[which(table$`Column %` == "TOTAL n")] <- sum(table(data[[current_rowVar]]))
  total_row <- table[table[, 1] == "TOTAL n", ]
  table <- table[-which(table[, 1] == "TOTAL n"), ]
  table <- rbind(table, total_row)
  table$Net[-nrow(table)] <- paste0(table$Net[-nrow(table)], "%")
  
  return(table)
  
}
#--------------------------
Subtable_Maker <- function(sheetDesign, current_rowVar, current_colVar){ 
  
  data <- sheetDesign@data
  data <- data[complete.cases(data[, unlist(sheetDesign@colVars)]), ]

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
  subtable[upIndices] <- paste0(subtable[upIndices], "% ", intToUtf8(8593))
  subtable[downIndices] <- paste0(subtable[downIndices], "% ", intToUtf8(8593))
  subtable[!upIndices & !downIndices] <- paste0(subtable[!upIndices & !downIndices], "%")
  
  subtable <- as.data.frame(subtable)
  
  colnames(subtable) <- paste(current_colVar, colnames(subtable), sep = '.')
  
  subtable <- rownames_to_column(subtable, "Column %")
  
  subtable <- rbind(subtable, c("TOTAL n", colSums(chisq$observed)))
  
  return(subtable)
}
#--------------------------
Footer_Maker <- function(sheetDesign, current_rowVar){ 
  
  data <- sheetDesign@data
  data <- data[complete.cases(data[, unlist(sheetDesign@colVars)]), ]
  
  return(paste0("Total n = ", nrow(data), "; ", nrow(sheetDesign@data)-nrow(data), " missing" ))
  
}
#--------------------------
#--------------------------

workbook <- Workbook_Maker()



