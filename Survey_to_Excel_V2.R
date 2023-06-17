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
setClass("Sheet",
         slots = list(
           data = "data.frame",
           rowVars = "list",
           colVars = "list"
         )
)

#--------------------------
sheets <- list(
  new("Sheet", data = mtcars, rowVars = list("gear"), colVars = list("vs", "am", "carb"))
)
#--------------------------
Sheet_Maker <- function(sheet){
  
  for (i in sheet@rowVars) {
    Table_Maker(sheet, i)
  }
  
}
#--------------------------
Table_Maker <- function(sheet, current_rowVar){
  #have the first two columns be Overall and Net
  
  table <- data.frame()
  
  for (i in sheet@colVars) {
    
    subtable <- Subtable_Maker(sheet, current_rowVar, i)

    if(i == sheet@colVars[1]){
      table <- subtable
    }else{
      table <- merge(table, subtable, by = "Column %", all = T)
    }
  }
  
  print(table)
  
}
#--------------------------
Subtable_Maker <- function(sheet, current_rowVar, current_colVar){ 
  
  data <- sheet@data

  subtable <- xtabs(~data[[current_rowVar]]+data[[current_colVar]],data=data)

  subtable_mean <- mean(subtable)
  
  chisq <- chisq.test(subtable)
  std_res <- chisq$stdres
  subtable <- prop.table(subtable, margin = 2)
  subtable <- apply(subtable, 2, round, digits = 2)
  subtable <- subtable * 100
  
  subtable <- format(subtable, justify = "right")
  subtable <- apply(subtable, 2, trimws)
  
  
  if(subtable_mean >= 4){ #4 is 80% o 5, and supposedly you need 80% of cells to be 5 or more
    subtable[std_res > 1.96 & std_res != "NaN"] <- paste(subtable[std_res > 1.96 & std_res != "NaN"], intToUtf8(8593))
    subtable[std_res < -1.96 & std_res != "NaN"] <- paste(subtable[std_res < -1.96 & std_res != "NaN"], intToUtf8(8595))
  }
  
  subtable <- as.data.frame(subtable)
  
  colnames(subtable) <- paste(current_colVar, colnames(subtable), sep = '.')
  
  subtable <- rownames_to_column(subtable, "Column %")
  
  return(subtable)
}
#--------------------------
#--------------------------

for (sheet in sheets) {
  Sheet_Maker(sheet)
}






