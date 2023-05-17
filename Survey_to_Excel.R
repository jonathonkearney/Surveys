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

addWorksheet(wb, sheetName = "Table of Contents")
writeData(wb, sheet =  "Table of Contents", x = "Table of Contents", startCol = 1, startRow = 1)
addStyle(wb, "Table of Contents", cols = 1, rows = 1, style = createStyle(fontSize = 16))

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
  toc_link_maker(sheetList)
}

#--------------------------

tables_maker <- function(sheet){
  
  currentrow <- 1
  
  for (i in 1:length(sheet[[1]])) {
    
    tabledata <- tabledata_maker(sheet[[1]][[i]], sheet[[2]], sheet[[3]])
    table_printer(tabledata, sheet[[4]], currentrow, sheet[[1]][[i]])
    currentrow <- currentrow + nrow(tabledata) + 5
  }
  
}

#--------------------------

tabledata_maker <- function(rowData, data, cols){

  #would be good to have an overall and a net column
  #The below calculation doesn't do False Deiscovery Rate Correction, 
  #which is when the significance level is adjusted when you do multiple comparisons 
  
  # tablelist <- list()
  table <- data.frame()
  
  for (i in 1:length(cols)) {
    
    xtabs <- xtabs(~data[[rowData]]+data[[cols[[i]]]],data=data)
    
    #If 20% or more of the cells are <5, and there are no 0 rows/margins - (simulate p doesn't like rows or columns of 0)
    #then simulate.p.value, because the sample is too small for normal chisq
    if(any(margin.table(xtabs, 1) == 0) | any(margin.table(xtabs, 2) == 0)){
      chisq <- chisq.test(xtabs)
    }else if(sum(xtabs < 5)/length(xtabs) > .20){
      chisq <- chisq.test(xtabs, simulate.p.value = TRUE)
    }else{
      chisq <- chisq.test(xtabs)
    }
    
    # chisq <- chisq.test(xtabs)
    
    std_res <- chisq$stdres
    
    xtabs <- prop.table(xtabs, margin = 2)
    
    xtabs <- apply(xtabs, 2, round, digits = 2)
    
    xtabs <- xtabs * 100
    
    #consider only applying the arrow if the cell n is larger than some amount?
    
    #Bonferroni correction on the cutoff
    bon_cutoff <- qnorm(p=1-((0.05/2)/(nrow(xtabs)*ncol(xtabs))))
    
    mytable <- format(xtabs, justify = "right")
    mytable <- apply(mytable, 2, trimws)
    mytable[std_res > bon_cutoff & std_res != "NaN"] <- paste(mytable[std_res > bon_cutoff & std_res != "NaN"], intToUtf8(8593))
    mytable[std_res < -bon_cutoff & std_res != "NaN"] <- paste(mytable[std_res < -bon_cutoff & std_res != "NaN"], intToUtf8(8595))
    
    mytable <- as.data.frame(mytable)
    
    colnames(mytable) <- paste(cols[i], colnames(mytable), sep = '.')
    
    mytable <- rownames_to_column(mytable, "Column %")
    
    if(i == 1){
      table <- mytable
    }else{
      table <- merge(mytable, table, by = "Column %", all = T)
    }
    
  }
  
  overalltable <- as.data.frame(prop.table(table(data[[rowData]])))
  colnames(overalltable) <- c("Column %", "Overall")
  
  overalltable$Overall <- round(overalltable$Overall, 2)
  overalltable$Overall <- overalltable$Overall * 100

  table <- full_join(table,overalltable, by = "Column %")
  
  table <- table %>% select(Overall, everything())

  table <- table %>% relocate(Overall)
  table <- table %>% relocate("Column %")

  table$Overall <- format(table$Overall, justify = "right")
  
  for (i in 2:length(table)) { #starts at 2 to skip first column
    table[,i] <- str_trim(table[,i])
    table[,i] <- gsub("(\\d)(?!.*\\d)", "\\1%", table[,i], perl = TRUE)
  }
  
  return(table)
}

#--------------------------

table_printer <- function(sigtable, sheetname, currentrow, rowname){
  
  tocrow <- currentrow
  headerrow <- tocrow + 1
  spanrow <- headerrow + 1
  tableheaderrow <- spanrow + 1
  tablerowsstart <- tableheaderrow + 1
  tablerowsend <- tablerowsstart + nrow(sigtable) - 1
  tablecolsend <- length(sigtable)
   
  ######## ADD TOC HYPERLINK ######## 
  
  writeFormula(wb, sheet = sheetname, startCol = 1, startRow = tocrow,
               x = makeHyperlinkString(sheet = "Table of Contents", col = 1, row = 1, text = "Back to ToC"))
  
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
toc_link_maker <- function(sheets){
  
  currentrow <- 2
  
  for (i in 1:length(sheets)) {
    sheet <- sheets[i]
    sheetname <- sheets[[i]][[4]]
    
    for (j in 1:length(sheet[[1]][[1]])) {
      writeData(wb, "Table of Contents", x = sheetname, startRow = currentrow, startCol = 1)
      
      writeFormula(wb, sheet = "Table of Contents", startCol = 2, startRow = currentrow,
                    x = makeHyperlinkString(sheet = sheetname, col = 1, row = 1, text = "Test"))
      
      currentrow <- currentrow + 1
    }
  }
  
}


#--------------------------

tester <- tabledata_maker("gear",  mtcars, list("vs", "am", "carb"))

sheets_maker(sheets)

saveWorkbook(wb, filename, overwrite = TRUE)

sheet <- sheets[1]
sheetname <- sheet[[1]][[4]]
tablename <- sheet[[1]][[1]]

