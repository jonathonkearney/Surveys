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
  
  writeDataTable(wb, "Sheet 1", sigtable, startRow = (6 + (i-2)*5), tableStyle = "TableStyleLight1")
}

up_arrow_format <- createStyle(fontColour = "blue")
down_arrow_format <- createStyle(fontColour = "red")
conditionalFormatting(wb, "Sheet 1", type = "contains", rule = "↑", cols = 1:1000, rows = 1:1000, style = up_arrow_format)
conditionalFormatting(wb, "Sheet 1", type = "contains", rule = "↓", cols = 1:1000, rows = 1:1000, style = down_arrow_format) 

saveWorkbook(wb, filename, overwrite = TRUE)
