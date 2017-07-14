# Load an excel file and convert a column to a category based on text colour.

library(xlsx)
# Hint:  after making a CellStyle object called cellStyle, run names(cellStyle) to get a list of functions that can be run to get different properties of the font/background colour

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

# the file to convert
filename = "~/Desktop/TextColours.xlsx"
# the file to write the results to
outfilename = "~/Desktop/TextColours_withCategories.csv"
#  the column you want to convert in the file.
columnToConvert = 2


# Load the workbook (we need to do it this way to get at the formatting options)
wb <- xlsx::loadWorkbook(filename)
# Get the sheets in the workbook
sheets <- getSheets(wb)
# Get the first sheet
sheet <- sheets[[1]]
# get rows in the sheet (the data for all rows)
rows <- getRows(sheet)
# get the data in the column to be converted
cells <- getCells(rows, colIndex =columnToConvert)

# Get the text colours in the column
colours = sapply(cells,getCellFontColour)
# cut out the first row (header)
colours = colours[2:length(colours)]
# Convert the colours to an integer 
colour.categories = as.integer(as.factor(colours))

# Read in the data (again, but as a data frame)
wb2 = xlsx::read.xlsx(filename, sheetIndex=1, header=TRUE)
# Add the extra column of colours converted to an integer
wb2 = cbind(wb2, colour.categories)
# Write out the data to a new file
write.csv(wb2, file=outfilename)
