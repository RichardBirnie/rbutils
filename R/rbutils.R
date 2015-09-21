# A collection of utility functions

#' Take a data frame and convert all factor variables to character variables
#'
#' @param df A data frame with at least one character column
#' @return The same data frame but with factors converted to character
factorToCharacter = function(df){
  df[sapply(df, is.factor)] <- lapply(df[sapply(df, is.factor)],
                                      as.character)
  df
}

#' Capitalise the first letter of each word in a string
#'
#' @param s A vector of one or more character strings
#' @return A vector of one or more character strings with the first letter of each word capitalised
#' @examples
#' # A single character string with two words
#' capwords('one two')#'
#' # Three character strings. The first has two words the last two
#'  have one word each
#' capwords(c('one two', 'three', 'four'))
capwords <- function(s, strict = FALSE) {
  cap <- function(s) paste(toupper(substring(s, 1, 1)),
                           {s <- substring(s, 2); if(strict) tolower(s) else s},
                           sep = "", collapse = " " )
  sapply(strsplit(s, split = " "), cap, USE.NAMES = !is.null(names(s)))
}

#' A simple wrapper around \code{write.xlsx} from the \code{xlsx} package
#'
#' @param x A \code{data.frame} to be saved as a single sheet in an excel file
#' @param file the path to the output file. This should include the '.xlsx' file extension
#' @param sheetName a character string with the sheet name.
#' @param col.names a logical value indicating if the column names of \code{x} are to be written along with \code{x} to the file.
#' @param row.names a logical value indicating whether the row names of \code{x} are to be written along with \code{x} to the file.
#' @param append a logical value indicating if \code{x} should be appended to an existing file. If \code{TRUE} the file is read from disk.
#' @param showNA a logical value. If set to FALSE, NA values will be left as empty cells
#' @param overwrite a logical value indicating if \code{x} should overwrite the contents of an existing sheet in the workbook
#'
#' @details Add a more verbose explanation here
#'
#' @examples Add some examples. Use some existing R data sets
#'
#' @seealso \code{\link[xlsx]{write.xlsx}}
saveXLSX = function(x, file, sheetName = "Sheet1", col.names = TRUE, row.names = TRUE,
                    append = FALSE, showNA = FALSE, overwrite=TRUE){
  if(file.exists(file)){
    wb = loadWorkbook(file)
    sheets = getSheets(wb)
    sheetExists = sheetName %in% names(sheets)
    if(sheetExists & overwrite){
      removeSheet(wb, sheetName=sheetName)
      newsheet = createSheet(wb, sheetName=sheetName)
      addDataFrame(x=x, sheet=newsheet, col.names=col.names, row.names=row.names, showNA=showNA)
      saveWorkbook(wb, file)
    } else {
      write.xlsx(as.data.frame(x), file=file, sheetName=sheetName, col.names=col.names, row.names=row.names, showNA=showNA, append=append)
    }
  } else {
    write.xlsx(as.data.frame(x), file=file, sheetName=sheetName, col.names=col.names, row.names=row.names, showNA=showNA, append=append)
  }
}
