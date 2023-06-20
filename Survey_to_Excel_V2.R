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
  
  saveWorkbook(wb, filename, overwrite = TRUE)
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
    
    sheetName <- workbook@sheets[[i]]@sheetName
    
    addWorksheet(wb, sheetName = sheetName)
    
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
      footerRow <- tableRowsEnd + 1
      
      ######## ADD TOC HYPERLINK ######## 
      
      # writeFormula(wb, sheet = sheetName, startCol = 1, startRow = tocRow,
      #              x = makeHyperlinkString(sheet = "Table of Contents", col = 1, row = 1, text = "Back to ToC"))

      ######## ADD TITLE ########
      
      writeData(wb, sheetName, paste(str_to_title(title), " by Banner"), startCol = 1, startRow = headerRow)
      addStyle(wb, sheetName, cols = 1, rows = headerRow, style = createStyle(textDecoration = "Bold"))
      
      ######## ADD SPANS ########
      
      spans <- names(table)[-c(1,2)]
      spans <- sub("\\..*", "", spans)
      currentSpan <- ""

      for (k in 1:length(spans)) {
        if(k == 1){ currentSpan <- spans[k]}
        else{
          if(spans[k] == currentSpan){spans[k] <- ""}
          else{currentSpan <- spans[k]}
        }
      }
      for (k in 1:length(spans)) {
        writeData(wb, sheetName, str_to_title(spans[k]), startCol = k+2, startRow = spanRow)
      }
      
      ####### ADD TABLE ########
      
      writeData(wb, sheetName, table, startRow = tableHeaderRow)
      
      colnames <- sub(".*\\.", "", colnames(table))
      for (k in 1:length(colnames)) {
        writeData(wb, sheetName, colnames[[k]], startCol = k, startRow = tableHeaderRow)
      }
      
      writeData(wb, sheetName, "Column %", startCol = 1, startRow = tableHeaderRow)
      
      ####### ADD FOOTER ########
      
      writeData(wb, sheetName, footer, startCol = 1, startRow = footerRow)

      ####### ADD STYLES ########
      rowColour <- "#dee3e2"
      boldStyle <- createStyle(textDecoration = "Bold")
      shadeAlignStyle <- createStyle(fgFill = "#d5dbda", halign = "right")
      boldShadeAlignStyle <- createStyle(textDecoration = "Bold", fgFill = rowColour, halign = "right")
      alignStyle <- createStyle(halign = "right")
      boldAlignStyle <- createStyle(textDecoration = "Bold", halign = "right")
      
      addStyle(wb, sheetName, cols = 1:tableColsEnd, rows = tableHeaderRow, style = boldStyle)
      addStyle(wb, sheetName, cols = 1:tableColsEnd, rows = spanRow, style = boldStyle)
      addStyle(wb, sheetName, cols = 1, rows = tableRowsStart:tableRowsEnd, style = boldStyle)
      
      for (i in seq(from = tableRowsStart, to = tableRowsEnd)) {
        if(i %% 2 != 0){ #odd
          addStyle(wb, sheetName, cols = 1:tableColsEnd, rows = i, style = shadeAlignStyle)
          addStyle(wb, sheetName, cols = 1, rows = i, style = boldShadeAlignStyle)
        }
        else{
          addStyle(wb, sheetName, cols = 1:tableColsEnd, rows = i, style = alignStyle)
          addStyle(wb, sheetName, cols = 1, rows = i, style = boldAlignStyle)
        }
      }
      
      ######## ADD CONDITIONAL FORMATTING ########
      
      upArrowFormat <- createStyle(fontColour = "blue")
      downArrowFormat <- createStyle(fontColour = "red")
      
      conditionalFormatting(wb, sheetName, type = "contains", rule = "↑", cols = 1:1000, rows = 1:1000, style = upArrowFormat)
      conditionalFormatting(wb, sheetName, type = "contains", rule = "↓", cols = 1:1000, rows = 1:1000, style = downArrowFormat) 
    }
  }
}


#--------------------------

workbook <- Workbook_Maker()

#access it by going workbook@sheets[[1]]@tableGroups[[1]]@table

Main()

