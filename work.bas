Sub FormatFuelSmartData()
'
' FormatFuelSmartData Macro
'
' Created by: RJ Tocci
'
' Module Last Updated: 03-06-23
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
' is no column for carrier name in the exported audit data.
' [XXXX] Optional: Use a range as a dictionary for comparing id number to
' Carrier/Supplier name, that way you could open/reference that info and copy it
' into a new column in the exported data.
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
' [DONE] Load notes from previous unrec to this one
' [DONE] Enter "taxes due at end of month" to Metroplex BOLs that fit the criteria.
' Very close to solving this, only issue is I get
' an error when I use method .SpecialCells(xlCellTypeVisible).Select and it fails to find any cells
' in the table that fit that criteria.
' HA--solved this by excluding one of the filters entirely!
' [XXXX] Use macro to open unrec template, copy relevant data, close it, and paste it so
' that you won't need to open "unrec template" manually before running the macro - WIP. Macro
' currently can't open/close the "unrec template" workbook, but does copy the info properly.
' [XXXX] Switch from using the "unrec template" workbook to just using the prevUnrec workbook
' so that I won't need two open files before running the macro, only one.
'
' IMPORTANT PLEASE READ:
' Pre-requisites:
' 1. Open workbook named "unrec template"
' 2. Open previous unrec report
' 3. Previous unrec report and new one must be in the same directory (I think...)
'
' How to export a new unrec report from Fuel Smart:
' 1. Open Fuel Smart and navigate to Research Reports in Fuel Payable
' 2. Select "Unreconciled Liability"
' 3. In Report Options window, select "Supplier" and "Detail"
' 4. Press "Enter" or select "Preview"
' 5. File > Save-As an Excel document named "unrec.xlsx" -> might have errors if different name
' 6. DONE
'
' MACRO WILL FAIL IF "unrec temple.xlsx" IS NOT ALREADY OPEN
' This file does not need to be in the same directory (as far as I can tell so far).
'
' Creating a macro to format a new unrec report from exported data from Fuel Smart.
' Macro will need to create new columns, new sheets, and delete redundant rows that
' fit certain criteria.
'
    ' Ask user to input name of previous unrec report so that you can use VLOOKUP to
    ' match all of the notes from the previous report:
    Dim prevUnrec As String
    
    ' Set value for prevUnrec while testing updates to the macro (comment out otherwise):
    'prevUnrec = "Unreconciled 03-08-23 local.xlsx"
    ' Comment-out until STOP if prevUnrec already given a value
TryAgain:
On Error GoTo Err1
    prevUnrec = InputBox(prompt:="Enter previous unrec file name:")
    If prevUnrec = "" Then
        Exit Sub
    Else
    ' STOP

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
    ' WIP--replace this code with references to "prevUnrec" instead of "unrec template"
        ' DONE
    Windows(prevUnrec).Activate
    ' Copy "Terms" sheet data:
    Sheets("Terms").Select
    Range("A1:J125").Select
    Selection.Copy
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
    ' TODO--copy the unrec carriers sheet in addition to terms?
    
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

    Dim prevUnrec As String
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
        
        ' Reformat the date columns (Pull & Drop):
        Range(Cells(2, 3), Cells(lRow, 4)).Select
        With Selection
            Selection.NumberFormat = "mm/dd/yy"
            .Value = .Value
        End With
    
        ' Start by asking user for name of previous unrec file
        ' NOTE: closing the box or entering wrong info causes an error
        ' potential errors -> loop InputBox function until input is valid or ""
            ' case sensitivity of file name - unknown if that would throw an error, maybe #REF?
            ' missing/wrong file extension in file name (#REF error for notes cells)
            ' file not open when macro runs (#REF error for notes cells)
            ' closing the input box when prompted for the file name
            ' #REF error if both unrec files are not in the same directory--hadn't thought of this one
        
        ' Try copying the formatting--could do this by creating a range for BLs in prevUnrec and
        ' a range for BLs in the new unrec, then copy and paste the font color and cell highlighting
        ' I don't think this works... untested
        'Dim srcRange As Range
        'srcLRow = prevUnrec.ActiveSheet.Cells(Rows.Count, "E").End(xlUp).Row
        'srcLCol = prevUnrec.ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
        'srcRange = prevUnrec.ActiveSheet.Cells(srcLRow, srcLCol)
        
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
        
        '' WIP--Refer to prevUnrec.Range of BLs and iterate through them to copy the relevant info?
        'Windows(prevUnrec).Activate
    
        ' Need to copy data from two final columns and paste without referencing
        ' the previous unrec report -> same as Copy&Paste values only
        Range(Cells(2, 10), Cells(lRow, 11)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
        ' Replace "0" values with blank cell:
        ' Had to modify this so that only cells with len=1 and val=0 would be replaced
        ' I think including xlWhole changes Ctrl+F in Excel and doesn't change it back after? Confirmed
        ' Should use xlPart to switch it back, not sure best way how since I don't need to use Find or Replace again
        'Selection.Replace What:="0", Replacement:=""      ' This code replaces ALL 0
        Selection.Replace 0, "", xlWhole                   ' xlWhole tells .Replace to look at the whole string
        
        ' So what I could do is run Find that fails and include an error-catch that exits the sub to reset
        ' xlWhole back to xlPart so I don't have to remember to do so manually after running this macro
        ' DONE--I put this code right before exiting the sub
    
        ' Autofit column width:
        Columns("A:K").EntireColumn.AutoFit
    
        ' Hide Drop column:
        Range("D:D").Select
        Selection.EntireColumn.Hidden = True
    
        ' NOTE: Notes columns are still copied to clipboard and surrounded by dashed lines
        ' I believe this code will clear it:
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
