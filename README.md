# ConvertExcelTextCoding

Load an excel file and convert a column to a category based on text colour.

Hint: In xlsx library,  after making a CellStyle object called cellStyle, run names(cellStyle) to get a list of functions that can be run to get different properties of the font/background colour.  Then you'll see things like:

```
cellStyle$getFont()
```

And then:

```
fontColour$getRgb()
```


There are some more hints for different methods here:

https://www.r-bloggers.com/when-life-gives-you-coloured-cells-make-categories/

https://www.extendoffice.com/documents/excel/1418-excel-count-sum-by-font-color.html

http://www.cpearson.com/excel/colors.aspx