Sub OrganizeDeckboxCSV()
'
' OrganizeDeckboxCSV Macro
' Converts .csv from deckbox.org to table
'

' Issues: Must have "Price" column
'

' Expected to be used on .csv file from deckbox.org with any number of
' additional columns.
' Requires the "Price" column to calculated total value.
'
    ' Declare variables:
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    lCol2 = lCol + 1
    
    ' Create "Total" column (Price * Count)
    Range("A1").Offset(, lCol).Select
    ActiveCell.FormulaR1C1 = "Total"
    
    ' Create a table with the data
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, _
        lCol2)), , xlYes).Name = "Table1"
    
    ' Set the formula for calculating Total column data:
    ' TODO: skip if no "Price" column exists; if you run this
    ' and the excel sheet lacks Price column, you'll get an error
    Range("A1").Offset(1, lCol).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=[@Count]*[@Price]"
      
    ' Change format of Total column from General to Currency
    ' Format the header:
    Range("A1").Offset(, lCol).Select
    Selection.NumberFormat = "General"
    ' Format the rest of the cells in the column:
    Range(Cells(2, lCol2), Cells(lRow, lCol2)).Select
    Selection.NumberFormat = "$#,##0.00"
    
    ' Create totals row
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").ShowTotals = True
    Range("Table1[[#Totals],[Count]]").Select
    ActiveSheet.ListObjects("Table1").ListColumns("Count"). _
        TotalsCalculation = xlTotalsCalculationSum
    ActiveCell.FormulaR1C1 = "=CONCATENATE(SUBTOTAL(109,[Count]),"" cards ("",SUBTOTAL(2,[Count]),"" unique)"")"
    
    ' Create totals row value for Total column, and highlight green
    ' Next line is having some difficulty since I updated this, but
    ' that should only be because I've failed to create a Total row
    Range("Table1[[#Totals],[Total]]").Select
    ActiveSheet.ListObjects("Table1").ListColumns("Total").TotalsCalculation = _
        xlTotalsCalculationSum
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Autofit columns - is there a better way to express this?
    ' Gonna go all the way to V regardless of num of columns -> any issue?
    Columns("A:V").EntireColumn.AutoFit
    
    Range("A1").Select
    
End Sub
