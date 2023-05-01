library(tidyverse)
library(openxlsx)
library(scales)

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
    
    xtabs <- xtabs * 100
    
    mytable <- format(xtabs, justify = "right")
    mytable[std_res > 2] <- paste(mytable[std_res > 2], intToUtf8(8593)) 
    mytable[std_res < -2] <- paste(mytable[std_res < -2], intToUtf8(8595))
    
    mytable <- as.data.frame(mytable)
    
    tablelist[[i]] <- mytable
    names(tablelist)[i] <- cols[i]
    
  }
  
  table <- do.call(cbind, tablelist)
  
  table <- tibble::rownames_to_column(table, "Column %")

  nettable <- as.data.frame(prop.table(table(data[[rowData]])))
  colnames(nettable) <- c("Column %", "NET")
  
  nettable$NET <- round(nettable$NET ,digit=2)
  nettable$NET <- nettable$NET * 100
  
  
  table <- full_join(table,nettable, by = "Column %")
  
  table <- table %>% select(NET, everything())
  
  table <- table %>% relocate(NET)
  table <- table %>% relocate("Column %")
  
  table$NET <- format(table$NET, justify = "right")
  
  for (i in 2:length(table)) { #starts at 2 to skip first column
    table[,i] <- str_trim(table[,i])
    table[,i] <- gsub("(\\d)(?!.*\\d)", "\\1%", table[,i], perl = TRUE)
  }
  
  return(table)
}  

#--------------------------

thedata <- mtcars
thecols <- list("vs", "am", "carb")
therows <- list("gear", "cyl")
filename <- "test.xlsx"

wb <- createWorkbook()

modifyBaseFont(wb, fontSize = 11, fontName = "Calibri")

addWorksheet(wb, sheetName = "Sheet 1")

#--------------------------

for (i in 1:length(therows)) {
  sigtable <- sig_table(therows[[i]], thedata, thecols)

  boldstyle <- createStyle(textDecoration = "Bold")
  shadestyle <- createStyle(fgFill = "#d5dbda")
  boldandshadestyle <- createStyle(textDecoration = "Bold", fgFill = "#d5dbda")
  
  headerrow <- (((i-2) * 7) + 8)
  spanrow <- headerrow + 1
  tableheaderrow <- spanrow + 1
  tablerowsstart <- tableheaderrow + 1
  tablerowsend <- tablerowsstart + length(unique(thedata[[therows[[i]]]]))
  tablecolsend <- length(sigtable)
  
  ######## ADD HEADER ########
  
  writeData(wb, "Sheet 1", paste(str_to_title(therows[[i]]), " by Banner"), startCol = 1, startRow = headerrow)
  
  addStyle(wb, "Sheet 1", cols = 1, rows = headerrow, style = boldstyle)

  ######## ADD SPANS ########
  
  spanstart <- 3
  spanend <- 0
  
  for (j in 1:length(thecols)) {
    
    writeData(wb, "Sheet 1", str_to_title(thecols[[j]]), startCol = spanstart, startRow = spanrow)
    
    spanlength <- length(unique(thedata[[thecols[[j]]]]))
    spanend <- (spanstart+spanlength -1)

    addStyle(wb, "Sheet 1", cols = spanstart:spanend, rows = spanrow, style = boldstyle)
    
    spanstart <- spanend + 1
  }
  
  ######## ADD TABLE ########
  
  writeData(wb, "Sheet 1", sigtable, startRow = tableheaderrow)
  
  colnames <- sub(".*\\.", "", colnames(sigtable))
  for (j in 1:length(colnames)) {
    writeData(wb, "Sheet 1", colnames[[j]], startCol = j, startRow = tableheaderrow)
  }
  
  writeData(wb, "Sheet 1", "Column %", startCol = 1, startRow = tableheaderrow)
  
  addStyle(wb, "Sheet 1", cols = 1:tablecolsend, rows = tableheaderrow, style = boldstyle)
  
  addStyle(wb, "Sheet 1", cols = 1, rows = tablerowsstart:tablerowsend, style = boldstyle)
  
  for (i in seq(from = tablerowsstart, to = tablerowsend, by = 2)) {
    addStyle(wb, "Sheet 1", cols = 1:tablecolsend, rows = i, style = shadestyle)
    addStyle(wb, "Sheet 1", cols = 1, rows = i, style = boldandshadestyle)
  }
  
  
  
}

up_arrow_format <- createStyle(fontColour = "blue")
down_arrow_format <- createStyle(fontColour = "red")
conditionalFormatting(wb, "Sheet 1", type = "contains", rule = "↑", cols = 1:1000, rows = 1:1000, style = up_arrow_format)
conditionalFormatting(wb, "Sheet 1", type = "contains", rule = "↓", cols = 1:1000, rows = 1:1000, style = down_arrow_format) 

saveWorkbook(wb, filename, overwrite = TRUE)
