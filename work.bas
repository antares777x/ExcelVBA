Sub FormatFuelSmartData()
'
' FormatFuelSmartData Macro
'
' Created by: RJ Tocci
'
' Module Last Updated: 03-13-23
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
' TODO:
'
' [XXXX] Include a new column for the carrier name for the freight
' invoices that you key. EDIT turns out this is impossible--there
' is no column for carrier name in the exported audit data.  I could
' add one in, but should I?
' [XXXX] Optional: Use a range as a table/dictionary for comparing id number to
' Carrier/Supplier name, that way you could open/reference that info and copy it
' into a new column in the exported data if desired.  Supplier short name is already
' present, so this would really only effect carriers.
'

    ' Create a new sheet for the formatted data:
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
' [XXXX] Set the font and font size to a specific value--shouldn't need to do
' this since I've never changed those settings, but if I share the macro it
' might be wise to add that to this subroutine.
'

    ' Initialize variables
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    lTotal = Application.WorksheetFunction.Sum(Selection)   ' total sum of keyed invoices
    lNum = Application.WorksheetFunction.Count(Selection)   ' total count of keyed invoices
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
' [DONE] Load notes from previous unrec to this one
' [DONE] Enter "taxes due at end of month" to Metroplex BOLs that fit the criteria.
' Very close to solving this, only issue is I get
' an error when I use method .SpecialCells(xlCellTypeVisible).Select and it fails to find any cells
' in the table that fit that criteria.
' HA--solved this by excluding one of the filters entirely!  Took me a while before I realized
' that solution would work.
' [DONE] Switch from using the "unrec template" workbook to just using the prevUnrec workbook
' so that I won't need two open files before running the macro, only one.
' [XXXX] Use macro to open previous unrec, copy relevant data, close it, and paste it so
' that you won't need to open the previous unrec manually before running the macro - WIP. Macro
' currently can't open/close the workbook, but does copy the info properly.  In most cases, the
' previous unrec report will already be open when I receive the new one and go to run this macro.
' [XXXX] Test whether or not the previous unrec needs the filters off for the macro to pull the
' data properly into the new unrec.  By default, the unrec reports are sent to me with some of
' the suppliers filtered out based on the notes.
' [DONE] Replace any of the data pulled into "Notes" that ="0".  NOTE--be careful not to
' replace ALL zeros in any Notes cells, only replace 0 when by itself.
'
' IMPORTANT PLEASE READ:
' Pre-requisites:
' 1. Open previous unrec report
' 2. Previous unrec report and new one must be in the same directory (unconfirmed)
'
' How to export a new unrec report from Fuel Smart:
' 1. Open Fuel Smart and navigate to Research Reports in Fuel Payable
' 2. Select "Unreconciled Liability"
' 3. In Report Options window, select "Supplier" and "Detail"
' 4. Press "Enter" or select "Preview"
' 5. File > Save-As an Excel document named "unrec.xlsx" -> this macro fails if the file is
' saved under a different name (working on improving that).
' 6. DONE
'
' MACRO WILL FAIL IF PREVIOUS UNREC IS NOT ALREADY OPEN
' This file does not need to be in the same directory (as far as I can tell so far).
'
' Creating a macro to format a new unrec report from exported data from Fuel Smart.
' Macro will need to create new columns, new sheets, sort the BOLs to match the
' sorting from previous reports, and delete redundant rows that fit certain criteria.
'
    ' Ask user to input name of previous unrec report so that you can use VLOOKUP to
    ' match all of the notes from the previous report:
    Dim prevUnrec As String
    Dim newUnrec As String
    newUnrec = ActiveWorkbook.Name
    
    ' Set value for prevUnrec while testing updates to the macro (comment out otherwise):
    'prevUnrec = "Unreconciled 03-08-23 local.xlsx"
    ' Comment-out until STOP if prevUnrec already given a value (i.e. above statement)
TryAgain:
On Error GoTo Err1
    prevUnrec = InputBox(prompt:="Enter previous unrec file name:")
    If prevUnrec = "" Then
        Exit Sub
    Else
    ' STOP

    ' Initialize variables:
    '' TODO: replace Select with variables
    Sheets(1).Range("A1").Select

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
        
    ' Create the "Terms" sheet based off the previous unrec report (needs to be open before macro is run)
    Windows(prevUnrec).Activate
    Sheets("Terms").Select
    
    ' Test replacing next two lines with single line?
    'Range("A1:J125").Copy
    Range("A1:J125").Select
    Selection.Copy
    
    Range("A1").Select
    Windows(newUnrec).Activate
    ' Create a new sheet called "Terms" and paste the data:
    Sheets.Add After:=ActiveSheet
    Sheets(2).Name = "Terms"
    Sheets("Terms").Select
    ActiveSheet.Paste
    ' Autofit column width and select A1:
    Columns("A:J").EntireColumn.AutoFit
    Range("A1").Select
    
    ' TODO--copy the unrec carriers sheet in addition to terms?
    ' As of now, it looks like the exported data for carrier/supplier are separated
    
    ' Rename Sheets 1, 2, and 3:
    ' The first sheet name may need to be modified based on what gets spit out by Fuel Smart
    Sheets(1).Name = "Unreconciled - Suppliers"
    
    ' Create Suppliers range:
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
    ' TODO: replace Select with variables
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    ' Reset filters and reset lRow value
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Select the cells in the Notes column and pull the old data from prevUnrec:
    Range(Cells(2, 10), Cells(lRow, 10)).Select
    ' Arbitrarily using row=1200 as the max row for the prevUnrec workbook -> TODO set variables for this?
    With Selection
        Selection.Value = "=IFNA(VLOOKUP(RC[-6],'[" & prevUnrec & "]Unreconciled - Suppliers'!R2C4:R1200C10,7,0),"""")"
    End With
    
    '' Try this to select the right answer and filter down to get the VLOOKUP function in each row
    Range("J2").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=10

    ' Copy & paste the values to remove references to prevUnrec:
    Range(Cells(2, 10), Cells(lRow, 10)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Replace "0" values with blank cell:
    Selection.Replace 0, "", xlWhole        ' xlWhole tells .Replace to look at the whole string
    
    ' Add filters to locate Metroplex BOLs that are only open because of taxes
    ActiveSheet.ListObjects("Table1").Range.AutoFilter _
        Field:=2, Criteria1:="Metroplex Energy Inc", _
        Field:=5, Criteria1:="<3500", Operator:=xlAnd

    ' NOTE--if the above filters fail to find any cells, this next line will cause an error
    ' Add "taxes due at end of month" for each relevant cell in Notes column:
    Range(Cells(2, 10), Cells(lRow, 10)).SpecialCells(xlCellTypeVisible).Select
    With Selection
        Selection.Value = "taxes due at end of month"
    End With

    ' Clear the filters:
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5  '<3500 bl_adj_amt
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2  'Metroplex
        
    ' Resize the columns:
    Columns("A:J").EntireColumn.AutoFit
    
    ' Attempt to find a non-existant string to reset xlWhole to xlPart:
    Dim DNErange As Range
    Set DNErange = Columns(1).Find("blablabla", , xlValues, xlPart, xlByRows, xlNext)
    If DNErange Is Nothing Then Exit Sub
    
    ' End the macro by selecting cell A1:
    Range("A1").Select
    Application.CutCopyMode = False     ' this will de-select the highlighted cells
    
    ' Comment out until STOP if value for prevUnrec predetermined:
    Exit Sub
    End If
Err1:
    MsgBox "File name error. Leave blank to exit."
    GoTo TryAgain
    ' STOP

End Sub

Sub FormatACH()
'
' Created by: RJ Tocci
'
' TODO:
' [XXXX] Update macro to work with open invoices as well as closed, depending
' on whichever option you're working with.
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
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C:R[1000]C)*-1"   ' row 1000 used as arbitrary max row
    Selection.Style = "Currency"
    
    ' Store invoice numbers and BL numbers as numbers instead of text:
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
    
    '' WIP
    '' Delete extra sheets (only if payments match):
    'Sheets("Sheet2").Select
    'ActiveWindow.SelectedSheets.Delete
    'Sheets("Sheet3").Select
    'ActiveWindow.SelectedSheets.Delete
    
    ' End macro by selecting A1:
    Range("A1").Select

End Sub

Sub FormatCarrierUnrec()
'
' Created by: RJ Tocci
'
' TODO:
' [DONE] Copy the carrier sheet so that you leave an unmodified version
' [DONE] Reformat the carrier sheet by removing the subtotals
' [DONE] Copy the VLOOKUP formula to read the previous unrec and update cells
' [DONE] Add error statements to prevent user input from breaking the macro
' [XXXX] Update cells w/highlighting from previous unrec report -> check notes from
' 3/6/23 since I found a good source from someone online who had the same question:
' https://stackoverflow.com/questions/22151426/vlookup-to-copy-color-of-a-cell
' Macro27() also contains a copy of this code for reference.
' [DONE] Change order of functions so that entering "" in the user prompt ends
' the macro BEFORE any of the other functions are called. Use InputBox function
' at the beginning of the macro instead of right before you need it.
' [DONE] Fix error replacing ALL "0" with "" so that it only replaces when val=0 AND len=1
' [DONE] Update the macro so that after you replace all "0" cells in Notes with "", Find/Replace
' will go back to using xlPart as default instead of xlWhole.
' [DONE] Create new sheet for carrier term length. Keep separate (temporarily) from the original Terms
' sheet until I can reconcile both of them with accurate information based on Fuel Smart. Currently both
' vendor names and term lengths from the Terms sheet DO NOT match data in Fuel Smart.
'
' Copies and reformats the Carrier sheet of the Unrec report such that
' all of the entries are organized into a single table.  Also copies the notes
' from previous unrec report (as specified by the user) and adds then to the
' appropriate Notes column to the current unrec report.
'
' WARNING
' MACRO WILL FAIL IF PREVIOUS UNREC REPORT (AS SPECIFIED BY USER) IS NOT ALREADY OPEN
' AND macro will also fail if both unrec reports are not saved in the same directory
'
' Enter "" in the user prompt to end the macro.  Otherwise, it will loop until the user input
' fits the criteria.
'

    ' Initialize variables for current workbook/unrec and previous unrec
    Dim newUnrec As String
    Dim prevUnrec As String
    newUnrec = ActiveWorkbook.Name ' this is the new unrec (currently open)

    ' TryAgain and Err1 being used for error catching
    ' If user input causes an error, loop until input is valid
    ' If input "", end the function early without changing anything
    ' That way if I go to run the macro and forgot to open the file,
    ' I can enter "" and not risk losing any data or overwriting an unsaved file
TryAgain:
On Error GoTo Err1
    ' User could still enter a valid string that's not a valid file name, in
    ' which case, the macro will work, but the cells in the notes columns will
    ' all appear as #REF errors
    prevUnrec = InputBox(prompt:="Enter previous unrec file name:")
    If prevUnrec = "" Then
        Exit Sub
    Else

        ' Copy the carrier sheet so that you leave an unmodified version
        ' After this exectues, the newly created copy will be the active worksheet
        Sheets("Unreconciled - Carriers").Select
        ActiveSheet.Copy Before:=Sheets(1)
    
        ' Rename worksheets:
        Sheets(1).Name = "Carriers formatted"
        Sheets(3).Name = "Carriers raw"
    
        ' Delete the subtotal lines from the new carrier sheet
        Range("A1").Select
        Selection.RemoveSubtotal
    
        ' Rename some of the column headers to shrink width
        ' Need to do this before creating a table to ensure proper size
        Range("A1").Value = "ID"
        Range("C1").Value = "Pull"
        Range("D1").Value = "Drop"
        Range("E1").Value = "BL"
        Range("F1").Value = "Amt"
        Range("G1").Value = "Gallons"
        Range("H1").Value = "Due"
        Range("I1").Value = "DPD"          ' Days Past Due
        Range("J1").Value = "RJ Notes"     ' AP notes
        Range("K1").Value = "Ted Notes"    ' SOPS notes
    
        ' Initialize variables to count row and column length
        lRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
        lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
        ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, _
            lCol)), , xlYes).Name = "Table1"
        
        ' Add totals row with sum and count:
        ActiveSheet.ListObjects("Table1").ShowTotals = True
        ActiveSheet.ListObjects("Table1").ListColumns("Amt"). _
            TotalsCalculation = xlTotalsCalculationSum
        ActiveSheet.ListObjects("Table1").ListColumns("Ted Notes").TotalsCalculation = _
            xlTotalsCalculationNone
        ActiveSheet.ListObjects("Table1").ListColumns("Gallons").TotalsCalculation = _
            xlTotalsCalculationCount
        
        ' Reformat the date columns (Pull, Drop, Due):
        Range(Cells(2, 3), Cells(lRow, 4)).Select
        With Selection
            Selection.NumberFormat = "mm/dd/yy"
            .Value = .Value
        End With
        Columns("H:H").NumberFormat = "mm/dd/yy"
        
        ' Create a new sheet for Carrier Terms
        Sheets.Add Before:=Sheets(4)
        Sheets(4).Name = "Carrier Terms"     ' the new sheet should now be the active sheet
        
        Windows(prevUnrec).Activate
        Sheets("Carrier Terms").Select
        Range("Table2[#All]").Select
        Selection.Copy
        
        Windows(newUnrec).Activate
        Sheets("Carrier Terms").Select
        Range("A1").Select
        ActiveSheet.Paste
        Columns("A:C").EntireColumn.AutoFit
        Range("A1").Select
        
        ' Update the formula for Due Dates based on the new carrier info:
        Sheets("Carriers formatted").Select
        Range("H2").Select
        ' Formula simply replaces the "15" with "VLOOKUP([@[carrier_name]],Table2,3,0)" so that it
        ' uses the actual term dates as they appear in Fuel Smart instead of 15 days for all vendors
        ActiveCell.FormulaR1C1 = _
            "=ROUNDDOWN(C3+VLOOKUP([@[carrier_name]],Table2,3,0),0)+IF(WEEKDAY(ROUNDDOWN(C3+VLOOKUP([@[carrier_name]],Table2,3,0),0))=7,2,IF(WEEKDAY(ROUNDDOWN(C3+VLOOKUP([@[carrier_name]],Table2,3,0),0))=1,1,0))"
        Range("H2").Select
        Selection.AutoFill Destination:=Range("Table1[Due]")
        Range("Table1[Due]").Select
        
        ' Create formula to pull notes into the Carriers formatted sheet:
        Range("J2").Select
        ' Set the VLOOKUP formula to reference prevUnrec for Notes from AP:
        ActiveCell.FormulaR1C1 = _
            "=IFNA(VLOOKUP([@BL],'" & prevUnrec & "'!Table1[[BL]:[Ted Notes]],6,0),"""")"
        Range("J2").Select
        ' Copy the formula for each row in the table:
        Selection.AutoFill Destination:=Range("Table1[RJ Notes]")
        Range("Table1[RJ Notes]").Select
        ' Repeat for the next column for notes from SOPS:
        Range("K2").Select
        ActiveCell.FormulaR1C1 = _
            "=IFNA(VLOOKUP([@BL],'" & prevUnrec & "'!Table1[[BL]:[Ted Notes]],7,0),"""")"
        Range("K2").Select
        ' Copy the formula for each row in the table:
        Selection.AutoFill Destination:=Range("Table1[Ted Notes]")
        Range("Table1[Ted Notes]").Select
                
        ' Need to copy data from two final columns and paste without referencing
        ' the previous unrec report -> same as Copy&Paste values only
        Sheets("Carriers formatted").Select
        Range(Cells(2, 10), Cells(lRow, 11)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
        ' Replace "0" values with blank cell:
        Selection.Replace 0, "", xlWhole        ' xlWhole tells .Replace to look at the whole string
        
        ' Autofit column width:
        Columns("A:K").EntireColumn.AutoFit
        
        ' Hide Drop column:
        Range("D:D").Select
        Selection.EntireColumn.Hidden = True
    
        ' Clear the clipboard:
        Application.CutCopyMode = False

        ' Finish Macro by selecting cell A1
        Range("A1").Select
        
        ' Attempt to find a non-existant string to reset xlWhole to xlPart:
        Dim DNErange As Range
        Set DNErange = Columns(1).Find("blablabla", , xlValues, xlPart, xlByRows, xlNext)
        If DNErange Is Nothing Then Exit Sub
        
        Exit Sub
        
    End If
    
' This is for error catching with the user input
' Enter "" to exit the macro
' All other errors will loop back to the user input prompt
' Entering a proper string name that doesn't match an open unrec file name will
' cause all of the Notes to read "#REF" since the VLOOKUP will fail.
Err1:
    MsgBox "File name error. Leave blank to exit."
    GoTo TryAgain

End Sub

Sub FormatSuspenseReport()
'
' Created by: RJ Tocci
'
' WIP MACRO
'
' Macro formats the suspense report data exported from Fuel Smart.  To export,
' go to Fuel Smart > Fuel Payable > Reports > Auto-Reconcile Invoice Summary
' select "Open" and choose "SH Open Invoice Report" then select "Preview."
' Save this data as any filename and the macro should work properly.
'
' The main use of this macro would be to get the data on suspended invoices in Fuel Smart
' and compare it to BOLs that are still open on the unrec report; this can be done with a VLOOKUP
' formula once the BLs are converted from text to number (otherwise it will appear as #N/A).  This
' can be another tool used to track down invoices that we have received but not been able to key.
'
' TODO
' [DONE] Hide irrelevant columns, fit to width.
' [XXXX] Add appropriate cell formatting for each column.
' [DONE] Create a "Notes" column for any user-input that might be relevant.
' [DONE] Create a column to track the number of times an invoice/BOL appears in the data. Sometimes
' invoices (and BOLs? researching) can appear more than once, and removing duplicates would be a mistake
' because the rest of the rows differ depending on circumstances.
' [XXXX] Organize above info about the process to get the data from Fuel Smart into a more clear/concise, numbered list
' that's easier to read.

    ' Copy the sheet with the original exported data to maintain an unmodified version
    ' After this exectues, the newly created copy will be the active worksheet
    Sheets(1).Select
    ActiveSheet.Copy Before:=Sheets(1)
    
    ' Rename worksheets -> this method doesn't need me to know the worksheet names; should apply this info past macros
    Sheets(1).Name = "modified"
    Sheets(2).Name = "raw"
    
    ' Rename column headers
    Range("A1").Value = "ID"
    Range("B1").Value = "Vendor"
    Range("C1").Value = "Invoice"
    Range("E1").Value = "Date"
    Range("F1").Value = "Due"
    Range("I1").Value = "Club"         ' destination
    Range("M1").Value = "Gallons"      ' invoiced gallons
    Range("O1").Value = "BL"
    Range("P1").Value = "Gross"        ' gross gallons
    Range("Q1").Value = "Net"          ' net gallons
    Range("R1").Value = "Amt"          ' invoiced amount
    ' Adding additional columns:
    Range("S1").Value = "Count"        ' counts the frequency of the invoice number; finds duplicates
    Range("T1").Value = "Notes"        ' Notes column for user input
    
    ' Reorganize data into an Excel table
    ' Initialize variables to count row and column length
    lRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    lCol = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(lRow, _
        lCol)), , xlYes).Name = "Table1"
    
    ' Add subtotal for total amount - WIP
    ' Add subtotal for invoice count - WIP
    ' Add subtotal for UNIQUE invoice count - WIP
    ' Sort table by invoice date
    ' OPTIONAL - create a table to verify some of the ap_vendor numbers--some look invalid/outdated
    
    ' Add formula for Count column:
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(Table1[Invoice],RC[-16])"
    
    ' Change formatting of date columns, gallons, etc
    'Range(Cells(2, 5), Cells(lRow, 6)).Select
    'Selection.NumberFormat = "m/d/yy"
    Columns("E:E").NumberFormat = "mm/dd/yy"   ' invoice date
    Columns("F:F").NumberFormat = "mm/dd/yy"   ' due date
    ' Change the BL column to store values as numbers instead of text
    Range(Cells(2, 15), Cells(lRow, 15)).Select
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With
    
    ' Format column width
    Columns("A:T").EntireColumn.AutoFit
    
    ' Hide irrelevant columns - D(?), G, H, J, L, N
    Range("D:D,G:H,J:J,L:L,N:N").Select
    Selection.EntireColumn.Hidden = True
    
    ' End macro by selecting A1:
    Range("A1").Select

End Sub
