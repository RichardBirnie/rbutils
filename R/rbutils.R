# A collection of utility functions

factorToCharacter = function(df){
  #takes a data frame and converts all factor variables to characters
  df[sapply(df, is.factor)] <- lapply(df[sapply(df, is.factor)],
                                      as.character)
  df
}

capwords <- function(s, strict = FALSE) {
  cap <- function(s) paste(toupper(substring(s, 1, 1)),
{s <- substring(s, 2); if(strict) tolower(s) else s},
sep = "", collapse = " " )
sapply(strsplit(s, split = " "), cap, USE.NAMES = !is.null(names(s)))
}

saveXLSX = function(df, file, sheetName = "Sheet1", col.names = TRUE, row.names = TRUE,
                    append = FALSE, showNA = TRUE, overwrite=TRUE){
  #This function is a simple wrapper around write.xlsx from the xlsx package
  #This version automatically overwrites any existing results
  if(file.exists(file)){
    wb = loadWorkbook(file)
    sheets = getSheets(wb)
    sheetExists = sheetName %in% names(sheets)
    #browser()
    if(sheetExists & overwrite){
      removeSheet(wb, sheetName=sheetName)
      newsheet = createSheet(wb, sheetName=sheetName)
      addDataFrame(x=df, sheet=newsheet, col.names=col.names, row.names=row.names, showNA=showNA)
      saveWorkbook(wb, file)
    } else {
      write.xlsx(as.data.frame(df), file=file, sheetName=sheetName, col.names=col.names, row.names=row.names, showNA=showNA, append=append)
    }
  } else {
    write.xlsx(as.data.frame(df), file=file, sheetName=sheetName, col.names=col.names, row.names=row.names, showNA=showNA, append=append)
  }
}
