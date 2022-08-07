Sub FormatFuelSmartData()
'
' FormatFuelSmartData Macro
'
' Created by: RJ Tocci
' Last updated: 07-18-22
'
' This macro will automatically format data exported from Fuel Smart.
' It will select and remove the redundant/unnecessary columns, delete
' duplicate data entries, and convert the remaining data to a table
' named "Table1".
'
' Note that the exported data is .XLS format and won't have tables.
'
' I might want to put all this on a separate sheet in the future and leave
' the main sheet with the exported data untouched.
'

    ' Remove duplicate entries:
    ActiveSheet.Range("$A:$AQ").RemoveDuplicates Columns:=2, Header:= _
        xlYes
        
    ' Hide/delete extra columns:
    Range("AL:AQ,AE:AJ,K:L,N:AB,F:H,C:C").Select
    Selection.EntireColumn.Hidden = True
    
    ' Convert the data to a table:
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, _
        lCol)), , xlYes).Name = "Table1"
    
    ' Add a subtotal and an invoice count:
    ActiveSheet.ListObjects("Table1").ShowTotals = True
    ActiveSheet.ListObjects("Table1").ListColumns("invoice_amount"). _
        TotalsCalculation = xlTotalsCalculationSum
    ActiveSheet.ListObjects("Table1").ListColumns("bl_no").TotalsCalculation = _
        xlTotalsCalculationNone
    ActiveSheet.ListObjects("Table1").ListColumns("invoice_number").TotalsCalculation = _
        xlTotalsCalculationCount
    
    ' Resize the bl_no column in case you have long BOL numbers:
    Columns("AK:AK").EntireColumn.AutoFit
    
    ' Format the amount values as currency:
    Columns("I:I").NumberFormat = "$#,##0.00"
    
    ' End the macro by selecting the cell A1:
    Range("A1").Select
    
End Sub
Sub FormatFuelSmartTally()
'
' FormatFuelSmartTally Macro
'
' Created by: RJ Tocci
'
' This macro formats my tally sheet before I print it.
'

    ' Initialize variables
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    lTotal = Application.WorksheetFunction.Sum(Selection)
    lRow = Selection.Rows.Count
    lCount = 1
    
    ' Initiate loop to create additional columns if needed:
    Do Until lRow < 46
        ' Count extra data beyond row 46
        ActiveSheet.Cells(47, lCount).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Cut
        ' AutoFit the column width:
        Columns(lCount).EntireColumn.AutoFit
        ' Move that data to the next column
        ActiveSheet.Cells(1, lCount + 1).Select
        ActiveSheet.Paste
        ' Update variables before repeating
        lCount = lCount + 1
        lRow = lRow - 46
    Loop
        
    ' Create a totals cell at the end of the last column.
    ActiveSheet.Cells(lRow + 1, lCount).Select
    Selection.Value = lTotal
    
    ' AutoFit the width of the final column:
    Columns(lCount).EntireColumn.AutoFit
    
    ' Color-code the total yellow:
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Format the total as currency
    Selection.NumberFormat = "$#,##0.00"
    
End Sub

Sub FormatNewUnrec()
'
' FormatNewUnrec Macro
'
' Created by: RJ Tocci
'
' Work-in-progress, but mostly done
' TODO:
' 1. Load notes from previous unrec to this one (WIP)
' 2. Add note "taxes due at end of month" for certain Metroplex entries (DONE)
'
' Pre-requisites:
' 1. Copy the 2nd and 3rd sheets from the unrec to the new data
' 2. Sheets need to be labelled correctly prior to using the macro
'
' Creating a macro to format a new unrec report from exported data from Fuel Smart
' Macro will need to create new columns, new sheets, and delete redundant rows that
' fit certain criteria.
'

    ' Initialize variables:
    Range("A1").Select

    ' Delete the dropped_timestamp column:
    Range("D:D").Select
    Range("D1").Activate
    Selection.Delete Shift:=xlToLeft
   
    ' Add columns "Terms", "Due Date", "Days Past Due", and "Notes":
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Terms"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "Due Date"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Days Past Due"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Notes"
    
    ' Rename Sheets 1, 2, and 3:
    ' The first sheet name may need to be modified based on what gets spit out by Fuel Smart
    Sheets("unrec").Select
    Sheets("unrec").Name = "Unreconciled - Suppliers"
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Unreconciled - Carriers"
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Terms"
    
    '' Create Suppliers range for the gas suppliers:
    Sheets("Terms").Select
    Columns("A:C").Select
    ' Still researching this... probably a better way to create a named range, but this works
    ActiveWorkbook.Names.Add Name:="Suppliers", RefersToR1C1:="=Terms!C1:C3"
    Range("A1").Select
    
    ' Initialize variables for row/column:
    Sheets("Unreconciled - Suppliers").Select
    Range("A1").Select
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Sort by due date before creating a table:
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").Sort.SortFields.Add2 Key:=Range _
        (Cells(2, 3), Cells(lRow, 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Unreconciled - Suppliers").Sort
    ' Make sure dimensions are correct with the added columns
        .SetRange Range(Cells(2, 1), Cells(lRow, 10))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Convert data ranges to a table:
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, _
        lCol)), , xlYes).Name = "Table1"
    
    ' Add formula for "Terms" column:
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],Suppliers,3,0)"
    
    ' Add formula for "Due Date" column:
    Range("H2").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUNDDOWN(RC[-5]+RC[-1],0)+IF(WEEKDAY(ROUNDDOWN(RC[-5]+RC[-1],0))=7,2,IF(WEEKDAY(ROUNDDOWN(RC[-5]+RC[-1],0))=1,1,IF(ROUNDDOWN(RC[-5]+RC[-1],0)=Terms!R4C6,1,0)))"
    
    ' Add formula for "Days Past Due" column:
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=TODAY()+IF(OR(WEEKDAY(TODAY())=5,WEEKDAY(TODAY())=6),MAX(4,Terms!R4C9),MAX(2,Terms!R4C9))-RC[-1]"
    
    ' Format columns pulled_timestamp, bl_adj_amt, Days Past Due:
    Columns("C:C").NumberFormat = "m/d/yy"
    Columns("E:E").NumberFormat = "#,##0.00"
    Columns("H:H").NumberFormat = "m/d/yy"
    Columns("I:I").NumberFormat = "#,##0"
    
    ' Remove entries with bl_adj_amount <1000:
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:= _
        "<1000", Operator:=xlAnd
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Reset filters and reset lRow value
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Set Notes for Metroplex invoices bl_adj_amt < 3500 = "taxes due at end of month"
    ' Set up the filters (Metroplex, blank notes, <3500 bl_adj_amt)
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2, Criteria1:= _
        "Metroplex Energy Inc"
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=10, Criteria1:= _
        "="
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:= _
        "<3500", Operator:=xlAnd
    ' Iterate through the blank cells and add "taxes due at end of month"
    Range(Cells(2, 10), Cells(lRow, 10)).SpecialCells(xlCellTypeVisible).Select
    With Selection
        Selection.Value = "taxes due at end of month"
    End With
    ' Clear the filters
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=10 'Blank notes
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5  '<3500 bl_adj_amt
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2  'Metroplex
    
    '' Create formula to pull data from previous unrec report
    '' Need to create a string to refer to the previous report, or
    '' ask the user to input a filename/file path; not sure how to do that
    '' I do know how to input the formula into the necessary cells though
    'ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=10, Criteria1:= _
    '    "="
    'Range(Cells(2, 10), Cells(lRow, 10)).SpecialCells(xlCellTypeVisible).Select
    'With Selection
    '    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-5],Suppliers,3,0)"
    'End With
    'ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=10 'Blank notes
    
    ' Resize the columns:
    Columns("A:J").EntireColumn.AutoFit
    
    ' End the macro by selecting cell A1:
    Range("A1").Select

End Sub

Sub FormatACH()
'
' Created by: RJ Tocci
'
' TODO:
' 1. Update macro to work with open invoices as well as closed
'
' SAP layout: GAS RECONCILIATION LAYOUT - ROBERT K
'
' Macro that will automatically format export ACH data from SAP
' Note that this macro will only work using Robert K. format from SAP first, and
' only for invoices that have already closed; the format is different for invoices
' that are still open, and it won't work.
'
'
    ' Initialize variables to count the row and column length:
    Range("A1").Select
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    lRowStart = lRow - 2
    lRowEnd = lRow + 1

    ' Delete the redundant last 4 rows:
    Range(Cells(lRowStart, 1), Cells(lRowEnd, 1)).Select
    Selection.EntireRow.Delete
    
    ' Add a new row after the first row and reset lRow variable:
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Formula to find the total and format it in F2:
    Range("F2").Select
    ''Formula sums the total for the column up to row 1000, which in some
    ''circumstances may be too small of a value, or may slow down the macro if
    ''too high of a value.  Probably a better way to do this.
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[1000]C)*-1"
    Selection.Style = "Currency"
    
    ' Store invoice numbers and BL numbers as numbers instead of text
    ' Need to use a loop; can't pass more than one argument to Value method
    Range(Cells(3, 1), Cells(lRow, 3)).Select
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With
    
    ' Rename column headers DocumentNo and Amount; add ACH DETAIL:
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "DocumentNo"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "ACH DETAIL"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=R[1]C"

    ' Add borders and right-alignment:
    ' Not sure how much of this is redundant, so I left all the code intact.
    Range(Cells(1, 1), Cells(lRow, lCol)).Select
    Range("A1").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Adjust column width:
    Columns("A:F").EntireColumn.AutoFit
    
    ' Color the first two rows:
    Range("A1:F2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    '' Delete extra sheets (only if payments match):
    'Sheets("Sheet2").Select
    'ActiveWindow.SelectedSheets.Delete
    'Sheets("Sheet3").Select
    'ActiveWindow.SelectedSheets.Delete
    
    ' End macro by selecting A1:
    Range("A1").Select

End Sub

