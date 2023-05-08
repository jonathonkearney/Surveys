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

filename <- "test.xlsx"

wb <- createWorkbook()

sheet1 <- list(list("gear", "cyl"), mtcars, list("vs", "am", "carb"), "mtcars")
addWorksheet(wb, sheetName = sheet1[[4]])
sheet2 <- list(list("eye_color"), starwars, list("homeworld", "gender"), "starwars")
addWorksheet(wb, sheetName = sheet2[[4]])
sheet3 <- list(list("relig"), gss_cat, list("partyid", "race", "marital"), "gss_cat")
addWorksheet(wb, sheetName = sheet3[[4]])

sheets <- list(sheet1, sheet2, sheet3)

#--------------------------

modifyBaseFont(wb, fontSize = 11, fontName = "Calibri")
rowcolour <- "#dee3e2"
boldstyle <- createStyle(textDecoration = "Bold")
shadestyle <- createStyle(fgFill = "#d5dbda")
boldandshadestyle <- createStyle(textDecoration = "Bold", fgFill = rowcolour)

up_arrow_format <- createStyle(fontColour = "blue")
down_arrow_format <- createStyle(fontColour = "red")

#--------------------------

sheets_maker <- function(sheetList){
  
  for (i in 1:length(sheetList)) {
    
    tables_maker(sheetList[[i]])  
  }
  
}

#--------------------------

tables_maker <- function(sheet){
  
  currentrow <- 1
  
  for (i in 1:length(sheet[[1]])) {
    
    tabledata <- tabledata_maker(sheet[[1]][[i]], sheet[[2]], sheet[[3]])
    table_printer(tabledata, sheet[[4]], currentrow, sheet[[1]][[i]])
    currentrow <- currentrow + nrow(tabledata) + 4
  }
  
}

#--------------------------

tabledata_maker <- function(rowData, data, cols){

  
  # tablelist <- list()
  table <- data.frame()
  
  for (i in 1:length(cols)) {
    
    xtabs <- xtabs(~data[[rowData]]+data[[cols[[i]]]],data=data)
    
    chisq <- chisq.test(xtabs)
    
    std_res <- chisq$stdres
    
    xtabs <- prop.table(xtabs, margin = 2)
    
    xtabs <- apply(xtabs, 2, round, digits = 2)
    
    xtabs <- xtabs * 100
    
    #consider only applying the arrow if the cell n is larger than some amount?
    
    mytable <- format(xtabs, justify = "right")
    mytable <- apply(mytable, 2, trimws)
    mytable[std_res > 2 & std_res != "NaN"] <- paste(mytable[std_res > 2 & std_res != "NaN"], intToUtf8(8593))
    mytable[std_res < -2 & std_res != "NaN"] <- paste(mytable[std_res < -2 & std_res != "NaN"], intToUtf8(8595))
    
    mytable <- as.data.frame(mytable)
    
    colnames(mytable) <- paste(cols[i], colnames(mytable), sep = '.')
    
    mytable <- rownames_to_column(mytable, "Column %")
    
    if(i == 1){
      table <- mytable
    }else{
      table <- merge(mytable, table, by = "Column %", all = T)
    }
    
  }
  
  nettable <- as.data.frame(prop.table(table(data[[rowData]])))
  colnames(nettable) <- c("Column %", "NET")
  
  nettable$NET <- round(nettable$NET, 2)
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

table_printer <- function(sigtable, sheetname, currentrow, rowname){
  
  headerrow <- currentrow
  spanrow <- headerrow + 1
  tableheaderrow <- spanrow + 1
  tablerowsstart <- tableheaderrow + 1
  tablerowsend <- tablerowsstart + nrow(sigtable) - 1
  tablecolsend <- length(sigtable)
    
  ######## ADD HEADER ########
  
  writeData(wb, sheetname, paste(str_to_title(rowname), " by Banner"), startCol = 1, startRow = headerrow)
  
  addStyle(wb, sheetname, cols = 1, rows = headerrow, style = boldstyle)
    
  ######## ADD CONDITIONAL FORMATTING ########
  
  conditionalFormatting(wb, sheetname, type = "contains", rule = "↑", cols = 1:1000, rows = 1:1000, style = up_arrow_format)
  conditionalFormatting(wb, sheetname, type = "contains", rule = "↓", cols = 1:1000, rows = 1:1000, style = down_arrow_format) 
    
  ######## ADD SPANS ########
  
  spans <- names(sigtable)[-c(1,2)]
  spans <- sub("\\..*", "", spans)
  currentspan <- ""
  
  for (j in 1:length(spans)) {
    if(j == 1){ currentspan <- spans[j]}
    else{
      if(spans[j] == currentspan){spans[j] <- ""}
      else{currentspan <- spans[j]}
    }
  }
  for (j in 1:length(spans)) {
    writeData(wb, sheetname, str_to_title(spans[j]), startCol = j+2, startRow = spanrow)
  }

  ####### ADD TABLE ########

  writeData(wb, sheetname, sigtable, startRow = tableheaderrow)

  colnames <- sub(".*\\.", "", colnames(sigtable))
  for (j in 1:length(colnames)) {
    writeData(wb, sheetname, colnames[[j]], startCol = j, startRow = tableheaderrow)
  }

  writeData(wb, sheetname, "Column %", startCol = 1, startRow = tableheaderrow)
  
  ####### ADD STYLES ########
  
  addStyle(wb, sheetname, cols = 1:tablecolsend, rows = tableheaderrow, style = boldstyle)
  addStyle(wb, sheetname, cols = 1:tablecolsend, rows = spanrow, style = boldstyle)
  addStyle(wb, sheetname, cols = 1, rows = tablerowsstart:tablerowsend, style = boldstyle)

  for (i in seq(from = tablerowsstart, to = tablerowsend, by = 2)) {
    addStyle(wb, sheetname, cols = 1:tablecolsend, rows = i, style = shadestyle)
    addStyle(wb, sheetname, cols = 1, rows = i, style = boldandshadestyle)
  }
    
}


#--------------------------

tester <- tabledata_maker("gear",  mtcars, list("vs", "am", "carb"))

sheets_maker(sheets)

saveWorkbook(wb, filename, overwrite = TRUE)


# 
# 
# 
# 
# xtabsTest <- xtabs(~relig + partyid ,data=gss_cat)
# # xtabsTest <- xtabs(~gear + vs ,data=mtcars)
# 
# chisqTest <- chisq.test(xtabsTest)
# 
# xtabsTest <- prop.table(xtabsTest, margin = 2)
# 
# std_resTest <- chisqTest$stdres
# 
# xtabsTest <- apply(xtabsTest, 2, round, digits = 2)
# 
# xtabsTest <- xtabsTest * 100
# 
# #consider only applying the arrow if the cell n is larger than some amount?
# 
# mytable <- format(xtabsTest, justify = "right")
# mytable <- apply(mytable, 2, trimws)
# 
# # mytable <- as.data.frame(mytable)
# # std_resTest <- as.data.frame.matrix(std_resTest)
# 
# 
# mytable[std_resTest > 2 & std_resTest != "NaN"] <- paste(mytable[std_resTest > 2 & std_resTest != "NaN"], intToUtf8(8593))
# mytable[std_resTest < -2& std_resTest != "NaN"] <- paste(mytable[std_resTest < -2 & std_resTest != "NaN"], intToUtf8(8595))
# 
# mytable <- as.data.frame(mytable)
# 
# colnames(mytable) <- paste(cols[i], colnames(mytable), sep = '.')
# 
# mytable <- rownames_to_column(mytable, "Column %")




