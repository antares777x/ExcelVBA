Sub FormatFuelSmartData()
'
' FormatFuelSmartData Macro
'
' Created by: RJ Tocci
'
' Last Update: 09-12-22
'
' This macro will automatically format data exported from Fuel Smart.
' It will select and hide the redundant/unnecessary columns, delete
' duplicate data entries, and convert the remaining data to a table
' named "Table1".
'
' Note that the exported data is .XLS format and won't have Excel tables
' when you open the file if you close it.
'
' Macro has been modified to keep an original copy of the exported data from
' Fuel Smart on the second sheet labeled "raw."
'
' TODO: none
'
'

    ' Create a new sheet for the data
    ActiveSheet.Copy Before:=Sheets(1)
    
    ' Rename sheet1 "formatted" and sheet2 "raw"
    Sheets(1).Name = "formatted"
    Sheets(2).Name = "raw"
    
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
    
    ' Store ap_vendor data as number instead of text:
    Range(Cells(2, 1), Cells(lRow, 1)).Select
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With
    
    ' End the macro by selecting the cell A1:
    Range("A1").Select
    
End Sub
Sub FormatFuelSmartTally()
'
' FormatFuelSmartTally Macro
'
' Created by: RJ Tocci
'
' This macro formats my tally sheet before I print it so that
' it fits on a single page.
'
' TODO:
' [DONE] Add a tally below the total
' [DONE] Organize the tally so that it's inline with data entries
' [DONE] Add keyer name to tally
' [DONE] Move tally, total, and keyer name to beginning of document
' [DONE] Fix final column width when only one entry exists in that column
' [DONE] Move column width code to outside the Do Until Loop for efficiency
'

    ' Initialize variables
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    lTotal = Application.WorksheetFunction.Sum(Selection)
    lNum = Application.WorksheetFunction.Count(Selection)
    lRow = Selection.Rows.Count
    lCount = 1
    
    ' Initiate loop to create additional columns if needed:
    Do Until lRow < 46
        ' Count extra data beyond row 46
        ActiveSheet.Cells(47, lCount).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Cut
        ' Move that data to the next column
        ActiveSheet.Cells(1, lCount + 1).Select
        ActiveSheet.Paste
        ' Update variables before repeating
        lCount = lCount + 1
        lRow = lRow - 46
    Loop
    
    ' Create new first row for tally, total, and keyer name
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    
    ' Display the total in the first cell
    Range("A1").Value = lTotal
    Selection.NumberFormat = "$#,##0.00"      ' format as currency
    
    ' Display the number of invoices keyed in the cell below
    Range("A2").Value = "=CONCATENATE(""" & lNum & ""","" invoices"")"
    
    ' Display the name of the keyer:
    Range("A3").Value = "RJ Tocci"
    
    ' Color-code the total, tally, and keyer yellow:
    Range("A1:A3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' AutoFit the columns:
    Columns("A:J").EntireColumn.AutoFit     ' using J as arbitrary max column
    
    ' Select first cell before ending the macro
    Range("A1").Select
            
End Sub

Sub FormatNewUnrec()
'
' FormatNewUnrec Macro
'
' Created by: RJ Tocci
'
' TODO:
' [DONE] Fix the sorting so that it is identical to the initial report sent by SOPS
' each morning.
' [XXXX] Load notes from previous unrec to this one (WIP; currently doing this manually
' with a VLOOKUP formula and the previous report).
' [XXXX] Use macro to open unrec template, copy relevant data, close it, and paste it so
' that you won't need to open "unrec template" manually before running the macro - WIP. Macro
' currently can't open/close the "unrec template" workbook, but does copy the info properly.
'
' IMPORTANT PLEASE READ:
' Pre-requisites:
' 1. Open workbook named "unrec template"
'
' MACRO WILL FAIL IF "unrec temple.xlsx" IS NOT ALREADY OPEN
'
' Creating a macro to format a new unrec report from exported data from Fuel Smart.
' Macro will need to create new columns, new sheets, and delete redundant rows that
' fit certain criteria.
'

    ' Initialize variables:
    '' TODO: replace Select with variables
    Sheets("unrec").Range("A1").Select

    ' Delete the dropped_timestamp column:
    '' TODO: replace Select with variables:
    Range("D:D").Select
    Range("D1").Activate
    Selection.Delete Shift:=xlToLeft
   
    ' Add columns "Terms", "Due Date", "Days Past Due", and "Notes":
    Range("G1").Value = "Terms"
    Range("H1").Value = "Due Date"
    Range("I1").Value = "Days Past Due"
    Range("J1").Value = "Notes"
        
    ' Create the "Terms" sheet (preq: open "unrec template")
    ' Activate the "unrec template" workbook (should already be open)
    Windows("unrec template").Activate
    ' Copy "Terms" sheet data:
    Sheets("Terms").Select
    Range("A1:J125").Select
    Selection.Copy
    ' Careful with this so you don't overwrite the unrec file you're using!
    ' Other workbook should be named "unrec" for this to work:
    Windows("unrec").Activate
    ' Create a new sheet called "Terms" and paste the data:
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Name = "Terms"
    Sheets("Terms").Select
    ActiveSheet.Paste ' TODO: replace this with adding a new sheet and labeling it "Terms"
    ' Autofit column width and select A1:
    Columns("A:J").EntireColumn.AutoFit
    Range("A1").Select
    
    ' Rename Sheets 1, 2, and 3:
    ' The first sheet name may need to be modified based on what gets spit out by Fuel Smart
    Sheets("unrec").Name = "Unreconciled - Suppliers"
    '' Commenting below lines out while I work on a macro that creates the "Terms" sheet
    '' Also appears that the "Unreconciled - Carriers" data is not needed
    'Sheets("Sheet1").Name = "Unreconciled - Carriers"
    'Sheets("Sheet2").Name = "Terms"
    
    ' Create Suppliers range for the gas suppliers:
    '' TODO: replace Select with variables:
    Sheets("Terms").Select
    Columns("A:C").Select
    ' Still researching this... probably a better way to create a named range, but this works
    ' Create named range "Suppliers" for formula simplification
    ' Note that this range includes the headers and actually includes the entire columns,
    ' so it should be somewhat future-proof when new suppliers are added (make sure to update the
    ' template if that happens so that you include the new supplier in the named range).
    ActiveWorkbook.Names.Add Name:="Suppliers", RefersToR1C1:="=Terms!C1:C3"
    Range("A1").Select
    
    ' Initialize variables for row/column:
    '' TODO: replace Select with variables:
    Sheets("Unreconciled - Suppliers").Select
    Range("A1").Select
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
        
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
    
    ' To match official unrec, sort by pulled_timestamp, then supplier, then Days Past Due
    ' First, sort by pulled_timestamp:
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1"). _
        Sort.SortFields.Clear
    ' Make sure dimensions are correct with the added columns
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range(Cells(2, 3), Cells(lRow, 3)), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Sort by supplier_name:
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range(Cells(2, 2), Cells(lRow, 2)), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Sort by "Days Past Due" in descending order:
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range(Cells(2, 9), Cells(lRow, 9)), SortOn:=xlSortOnValues, Order _
        :=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Unreconciled - Suppliers").ListObjects("Table1" _
        ).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Store supplier_id numbers and as numbers instead of text
    '' TODO: replace Select with variables
    Range(Cells(2, 1), Cells(lRow, 1)).Select
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With
    
    ' Format columns supplier_id, pulled_timestamp, bl_adj_amt, Days Past Due:
    Columns("C:C").NumberFormat = "m/d/yy"   ' pulled_timestamp
    Columns("E:E").NumberFormat = "#,##0.00" ' bl_adj_amt
    Columns("H:H").NumberFormat = "m/d/yy"   ' Due Date
    Columns("I:I").NumberFormat = "#,##0"    ' Days Past Due
    
    ' Remove entries with bl_adj_amount <1000:
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:= _
        "<1000", Operator:=xlAnd
    '' TODO: replace Select with variables
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Reset filters and reset lRow value
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Set Notes for Metroplex invoices bl_adj_amt < 3500 = "taxes due at end of month"
    ' Set up the filters (Metroplex, blank notes, <3500 bl_adj_amt)
    ActiveSheet.ListObjects("Table1").Range.AutoFilter _
        Field:=2, _
        Criteria1:="Metroplex Energy Inc"
    ActiveSheet.ListObjects("Table1").Range.AutoFilter _
        Field:=10, _
        Criteria1:="="
    ActiveSheet.ListObjects("Table1").Range.AutoFilter _
        Field:=5, _
        Criteria1:="<3500", _
        Operator:=xlAnd
    ' Iterate through the blank cells and add "taxes due at end of month"
    '' TODO replace Select with variables
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
' [XXXX] Update macro to work with open invoices as well as closed
' [XXXX] Delete extra sheets if payments match
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
    Range("B1").Value = "DocumentNo"
    Range("F1").Value = "Amount"
    Range("D2").Value = "ACH DETAIL"
    Range("E2").Value = "=R[1]C"

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
