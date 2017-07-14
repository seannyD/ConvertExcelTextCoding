# Load an excel file and convert a column to a category based on text colour.

library(xlsx)
# Hint:  after making a CellStyle object called cellStyle, run names(cellStyle) to get a list of functions that can be run to get different properties of the font/background colour

filename = "~/Desktop/TextColours.xlsx"
outfilename = "~/Desktop/TextColours_withCategories.csv"
columnToConvert = 2

getCellFontColour = function(cell){
  # Font colours can either be indexed or full hex values
  cellStyle = getCellStyle(cell)
  fontColour = cellStyle$getFont()
  rgb <- tryCatch(fontColour$getRgb(), error = function(e) NULL)
  if(!is.null(rgb)){
    return(rgb)
  } else{
    return(cellStyle$getFont()$getThemeColor())
  }
}



wb <- xlsx::loadWorkbook(filename)
sheets <- getSheets(wb)
sheet <- sheets[[1]]
rows <- getRows(sheet)
cells <- getCells(rows, colIndex =columnToConvert)

colours = sapply(cells,getCellFontColour)
# cut out the first row (header)
colours = colours[2:length(colours)]
colour.categories = as.integer(as.factor(colours))

# Read in the data as a data frame, 
wb2 = xlsx::read.xlsx(filename, sheetIndex=1, header=TRUE)
wb2 = cbind(wb2, colour.categories)
write.csv(wb2, file=outfilename)
