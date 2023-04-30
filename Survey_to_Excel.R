library(tidyverse)
library(openxlsx)

#--------------------------
#If the standardized residual is greater than 2 or less than -2 in absolute value, then the cell is considered to be significantly contributing to the chi-squared statistic at a 5% level of significance.
#The standardized residual represents the number of standard deviations that the observed frequency deviates from the expected frequency
#--------------------------

rm(list = ls())

#--------------------------

sig_table <- function(rowData,data, cols){
  
  tablelist <- list()
  
  for (i in 1:length(cols)) {
    
    xtabs <- xtabs(~data[[rowData]]+data[[cols[[i]]]],data=data)
    
    chisq <- chisq.test(xtabs)
    
    std_res <- chisq$stdres
    
    xtabs <- prop.table(xtabs, margin = 2)
    
    xtabs <- apply(xtabs, 2, round, digits = 2)
    
    mytable <- format(xtabs, justify = "right")
    mytable[std_res > 2] <- paste(mytable[std_res > 2], intToUtf8(8593)) 
    mytable[std_res < -2] <- paste(mytable[std_res < -2], intToUtf8(8595))
    
    mytable <- as.data.frame(mytable)
    
    tablelist[[i]] <- mytable
    names(tablelist)[i] <- cols[i]
    
  }
  
  table <- do.call(cbind, tablelist)
  
  table <- tibble::rownames_to_column(table, rowData)
  
  return(table)
}  

#--------------------------

thedata <- mtcars
thecols <- list("vs", "am", "carb")
therows <- list("gear", "cyl")
filename <- "test.xlsx"

wb <- createWorkbook()
addWorksheet(wb, sheetName = "Sheet 1")

for (i in 1:length(therows)) {
  sigtable <- sig_table(therows[[i]], thedata, thecols)
  
  rownum <- (((i-2) * 7) + 8)
  
  writeData(wb, "Sheet 1", "Title Goes In This Row", startCol = 1, startRow = rownum)
  
  currentspancell <- 2 #to keep track of which cell to put the span in
  
  #go through columns and add span titles
  for (j in 1:length(thecols)) {
    
    spanlength <- length(unique(thedata[[thecols[[j]]]]))
    
    writeData(wb, "Sheet 1", "Span Title", startCol = currentspancell, startRow = rownum  + 1)
    
    currentspancell <- currentspancell + spanlength
  }
  
  writeDataTable(wb, "Sheet 1", sigtable, startRow = rownum + 2,
                 tableStyle = "TableStyleLight1",
                 withFilter = FALSE)
}

up_arrow_format <- createStyle(fontColour = "blue")
down_arrow_format <- createStyle(fontColour = "red")
conditionalFormatting(wb, "Sheet 1", type = "contains", rule = "↑", cols = 1:1000, rows = 1:1000, style = up_arrow_format)
conditionalFormatting(wb, "Sheet 1", type = "contains", rule = "↓", cols = 1:1000, rows = 1:1000, style = down_arrow_format) 

saveWorkbook(wb, filename, overwrite = TRUE)
