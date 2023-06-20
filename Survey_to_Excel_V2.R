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
setClass("Sheet_Design",
         slots = list(
           data = "data.frame",
           rowVars = "list",
           colVars = "list",
           sheetName = "character"
         )
)
setClass("Work_Book", #beacause workbook is being used by openxlsx... grr
         slots = list(
           sheets = "list"
         )
)
setClass("Sheet",
         slots = list(
           tableGroups = "list",
           sheetName = "character"
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
filename <- "testV2.xlsx"

wb <- createWorkbook()

sheet_designs <- list(
  new("Sheet_Design", data = mtcars, rowVars = list("gear"), colVars = list("vs", "am", "carb"), sheetName = "mtcars"),
  new("Sheet_Design", data = starwars, rowVars = list("eye_color"), colVars = list("homeworld", "gender"), sheetName = "starwars")
)
#--------------------------
Main <- function(){
  workbook <- Workbook_Maker()
  
  Printer(workbook)
}
#--------------------------
Workbook_Maker <- function(){
  
  workbook <- new("Work_Book", sheets = list())
  for (sheetDesign in sheet_designs) {
    sheet <- Sheet_Maker(sheetDesign)
    workbook@sheets <- c(workbook@sheets, sheet)
  }

  return(workbook)
}

#--------------------------
Sheet_Maker <- function(sheetDesign){
  
  sheet <- new("Sheet", tableGroups = list(), sheetName = sheetDesign@sheetName)
  for (rowVar in sheetDesign@rowVars) {
    title <- rowVar
    table <- Table_Maker(sheetDesign, rowVar)
    footer <- Footer_Maker(sheetDesign, rowVar)
      
    tableGroup <- new("Table_Group", title = title, table = table, footer = footer)
    sheet@tableGroups <- c(sheet@tableGroups, tableGroup)
  }
  return(sheet)
  
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
Printer <- function(workbook){
  
  currentRow <- 1
  
  for(i in 1:length(workbook@sheets)){
    
    print(i)
    
    for(j in 1:length(workbook@sheets[[i]]@tableGroups)){

      print(j)
      
      title <- workbook@sheets[[i]]@tableGroups[[j]]@title
      table <- workbook@sheets[[i]]@tableGroups[[j]]@table
      footer <- workbook@sheets[[i]]@tableGroups[[j]]@footer
      sheetName <- workbook@sheets[[i]]@sheetName
      
      tocRow <- currentRow
      headerRow <- tocRow + 1
      spanRow <- headerRow + 1
      tableHeaderRow <- spanRow + 1
      tableRowsStart <- tableHeaderRow+ 1
      tableRowsEnd <- tableRowsStart + nrow(table) - 1
      tableColsEnd <- length(table)
      
      ######## ADD TOC HYPERLINK ######## 
      
      # writeFormula(wb, sheet = sheetname, startCol = 1, startRow = tocrow,
      #              x = makeHyperlinkString(sheet = "Table of Contents", col = 1, row = 1, text = "Back to ToC"))
      
      
    }
    
  }
  
}


#--------------------------

workbook <- Workbook_Maker()

#access it by going workbook@sheets[[1]]@tableGroups[[1]]@table

Main()

