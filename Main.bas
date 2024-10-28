Attribute VB_Name = "Module1"
Sub Analysis()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim wsOverview As Worksheet
    Set wsOverview = wb.Sheets("Overview")
    Dim wsAnal As Worksheet
    Set wsAnal = wb.Sheets("DNA Data")
    Dim wsWater As Worksheet
    Set wsWater = wb.Sheets("NTC Data")
    Dim wsFormat As Worksheet
    Set wsFormat = wb.Sheets("Formats")
    Dim isRetest As VbMsgBoxResult
    Dim retestNum As Variant
    Dim isValid As Boolean
    isValid = False
    Dim retestForm As RetestUserForm
    Set retestForm = New RetestUserForm
    Dim foundRetest As Range
    
    wsOverview.Unprotect Password:="Op3narray"
    
    isRetest = MsgBox("Is this a Retest?", vbYesNo + vbQuestion, "Retest Confirmation")
    If isRetest = vbYes Then
        Do
        ' Display an input box asking for the retest number
        retestNum = InputBox("What retest number is this? Please enter a numeric value", "Enter Retest Number")
        Set foundRetest = wsOverview.Range("A1:Z200").Find(What:="Retest #" & retestNum, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not foundRetest Is Nothing Then
            MsgBox "Retest #" & retestNum & " already exists.  If you need to reimport the data please remove the table in the overview tab and associating analysis sheets."
            Exit Sub
        End If
        ' Check if the input is numeric and not empty
        If IsNumeric(retestNum) And retestNum <> "" Then
            isValid = True
        ElseIf retestNum = "" Then
            ' If the user cancels or leaves it empty, exit the loop
            MsgBox "No input provided. Exiting."
            Exit Sub
        Else
            MsgBox "Please enter a valid numeric value.", vbExclamation
        End If
    Loop Until isValid
    retestForm.Show
    
    ' After the form is hidden, process the results
    If retestForm.chkFunctional.Value = True And retestForm.chkNTC.Value = True Then
        Retest retestNum, True, True
    ElseIf retestForm.chkNTC.Value = True Then
        Retest retestNum, False, True
    ElseIf retestForm.chkFunctional.Value = True Then
        Retest retestNum, True, False
    End If
    
    ' Unload the form from memory
    Unload retestForm
    
    Else
    MsgBox "Please select the Taqman Genotyper export file for inital analysis"
    filePath2 = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select an Excel or CSV file")
    If filePath2 = "False" Then
        MsgBox "No File selected"
        Exit Sub
    End If
    MsgBox "Please select the NTC water plate file for initial analysis"
    filePath3 = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select an Excel or CSV file")
    If filePath3 = "False" Then
        MsgBox "No File selected"
        Exit Sub
    End If
    
    Dim wbExport As Workbook
    Set wbExport = Workbooks.Open(filePath2, UpdateLinks:=False)
    Dim ws1 As Worksheet
    Set ws1 = wbExport.Sheets(1)
    ExportData ws1, wsAnal, wsWater, wsOverview
    wbExport.Close SaveChanges:=False
    
    Dim wbWater As Workbook
    Set wbWater = Workbooks.Open(filePath3, UpdateLinks:=False)
    Dim ws2 As Worksheet
    Set ws2 = wbWater.Sheets(1)
    WaterData ws2, wsWater, wsOverview, "A3", retestNum
    wbWater.Close SaveChanges:=False
    
    wsOverview.Activate
    End If
    wsOverview.Protect Password:="Op3narray"
End Sub
    
Sub Retest(ByVal retestNum As Variant, ByRef func As Boolean, ByRef water As Boolean)
    Dim wb As Workbook
    Dim wsOverview As Worksheet, wsFormat As Worksheet, wsAnal As Worksheet, wsWater As Worksheet
    Dim filePath2 As Variant, filePath3 As Variant

    ' Assign workbook and worksheet references
    Set wb = ThisWorkbook
    Set wsOverview = wb.Sheets("Overview")
    Set wsFormat = wb.Sheets("Formats")
    Set wsAnal = wb.Sheets("DNA Data")
    Set wsWater = wb.Sheets("NTC Data")
    
    wsOverview.Unprotect Password:="Op3narray"

    ' Check for function and water test conditions
    If func = True And water = True Then
        ' Prompt for the Taqman Genotyper export file
        MsgBox "Please select the Taqman Genotyper export file for Retest #" & retestNum
        filePath2 = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select an Excel or CSV file")
        If filePath2 = "False" Then
            MsgBox "No file selected."
            Exit Sub
        End If

        ' Prompt for the NTC water plate file
        MsgBox "Please select the NTC water plate file for Retest #" & retestNum
        filePath3 = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select an Excel or CSV file")
        If filePath3 = "False" Then
            MsgBox "No file selected."
            Exit Sub
        End If
        Createnewtables wsOverview, wsFormat, retestNum, True, True
        Createnewsheets wsAnal, wsWater, retestNum, True, True
        Dim wsFuncRetest As Worksheet, wsWaterRetest As Worksheet
        Set wsFuncRetest = wb.Sheets("DNA Data Retest #" & retestNum)
        Set wsWaterRetest = wb.Sheets("NTC Data Retest #" & retestNum)
        
        Dim wbExport As Workbook
        Set wbExport = Workbooks.Open(filePath2, UpdateLinks:=False)
        Dim ws1 As Worksheet
        Set ws1 = wbExport.Sheets(1)
        ExportData ws1, wsFuncRetest, wsWaterRetest, wsOverview
        wbExport.Close SaveChanges:=False
        
        Dim wbWater As Workbook
        Set wbWater = Workbooks.Open(filePath3, UpdateLinks:=False)
        Dim ws2 As Worksheet
        Set ws2 = wbWater.Sheets(1)
        WaterData ws2, wsWaterRetest, wsOverview, "A3", retestNum
        wbWater.Close SaveChanges:=False
        
        wsOverview.Activate

    ElseIf func = True And water = False Then
        ' Prompt for the Taqman Genotyper export file if only func is true
        MsgBox "Please select the Taqman Genotyper export file for Retest #" & retestNum
        filePath2 = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select an Excel or CSV file")
        If filePath2 = "False" Then
            MsgBox "No file selected."
            Exit Sub
        End If
        Createnewtables wsOverview, wsFormat, retestNum, True, False
        Createnewsheets wsAnal, wsWater, retestNum, True, True
        Set wsFuncRetest = wb.Sheets("DNA Data Retest #" & retestNum)
        Set wsWaterRetest = wb.Sheets("NTC Data Retest #" & retestNum)
        
        Set wbExport = Workbooks.Open(filePath2, UpdateLinks:=False)
        Set ws1 = wbExport.Sheets(1)
        ExportData ws1, wsFuncRetest, wsWaterRetest, wsOverview
        wbExport.Close SaveChanges:=False
        
        wsOverview.Activate

    ElseIf water = True And func = False Then
        ' Prompt for the NTC water plate file if only water is true
        MsgBox "Please select the NTC water plate file for Retest #" & retestNum
        filePath3 = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls; *.csv), *.xlsx; *.xls; *.csv", , "Select an Excel or CSV file")
        If filePath3 = "False" Then
            MsgBox "No file selected."
            Exit Sub
        End If
        Createnewtables wsOverview, wsFormat, retestNum, False, True
        Createnewsheets wsAnal, wsWater, retestNum, False, True
        Set wsWaterRetest = wb.Sheets("NTC Data Retest #" & retestNum)
        
        Set wbWater = Workbooks.Open(filePath3, UpdateLinks:=False)
        Set ws2 = wbWater.Sheets(1)
        WaterData ws2, wsWaterRetest, wsOverview, "A3", retestNum
        wbWater.Close SaveChanges:=False
        
        wsOverview.Activate
        
    ElseIf func = False And water = False Then
        Exit Sub
    End If
    wsOverview.Activate
    wsOverview.Protect Password:="Op3narray"
End Sub

Sub Createnewtables(ByVal wsOverview As Worksheet, ByVal wsFormat As Worksheet, ByVal retestNum As Variant, ByVal func As Boolean, ByVal water As Boolean)
    Dim targetCell As Range
    Dim foundCell As Range
    Dim copyRange1 As Range
    Dim copyRange2 As Range
    Dim copyRange3 As Range
    Dim copyRange4 As Range
    Dim PRnum As Variant
    Dim foundRetest As Range
    
    ' Define the range to copy
    Set copyRange1 = wsFormat.Range("D2")
    Dim originalMergedRange As Range
    Set originalMergedRange = copyRange1.MergeArea
    originalMergedRange.Copy
    Set copyRange2 = wsFormat.Range("D3:J10")
    Set copyRange3 = wsFormat.Range("D9:J10")
    Set copyRange4 = wsFormat.Range("D3:J7")
    
    ' Check the value of retestNum to determine action
    If retestNum = 1 Then
        ' Copy data for the first test
        Set targetCell = wsOverview.Range("J6")
        With targetCell
            .PasteSpecial Paste:=xlPasteAllUsingSourceTheme  ' Retains source theme formatting
            .MergeCells = True  ' Merge the target cell
            .Value = originalMergedRange.Value  ' Set the value of the merged cell
        End With
        Application.CutCopyMode = False
        wsOverview.Cells(6, 10).Value = "Retest #1"
        wsOverview.Cells(6, 10).Font.Bold = True
        
        Set targetCell = targetCell.Offset(1, 0)
        
        If func = True And water = True Then
            copyRange2.Copy
            targetCell.PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False
        ElseIf func = True And water = False Then
            copyRange4.Copy
            targetCell.PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False
        ElseIf func = False And water = True Then
            copyRange3.Copy
            targetCell.PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False
        End If
    Else
        ' Prompt the user for the PR number
        PRnum = InputBox("Please enter PR number", "Input PR Number")
        
        ' Check if user input is valid (not empty and numeric)
        If PRnum = "" Or Not IsNumeric(PRnum) Then
            MsgBox "Invalid input. Please enter a numeric PR number."
            Exit Sub
        End If

        ' Find the first non-empty cell in column B or J depending on retestNum
        If retestNum Mod 2 = 0 Then
        Set foundCell = wsOverview.Range("C1:C200").Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
        If Not foundCell Is Nothing Then
            Set targetCell = foundCell.Offset(2, -1) ' Set the target to the next row
            
            ' Copy the original merged cell if it exists
            If originalMergedRange.MergeCells Then
                originalMergedRange.Copy
                
                With targetCell
                    .PasteSpecial Paste:=xlPasteAllUsingSourceTheme  ' Retains source theme formatting
                    .MergeCells = True  ' Merge the target cell
                    .Value = "Retest #" & retestNum & " PR#" & PRnum ' Set the retest and PR number value
                    .Font.Bold = True  ' Set font to bold
                End With
            End If
            Set targetCell = targetCell.Offset(1, 0)
            Application.CutCopyMode = False
                If func = True And water = True Then
                    copyRange2.Copy
                    targetCell.PasteSpecial Paste:=xlPasteAll
                ElseIf func = True And water = False Then
                    copyRange4.Copy
                    targetCell.PasteSpecial Paste:=xlPasteAll
                ElseIf func = False And water = True Then
                    copyRange3.Copy
                    targetCell.PasteSpecial Paste:=xlPasteAll
                End If
                Application.CutCopyMode = False
            End If
            
        ElseIf retestNum Mod 2 <> 0 Then
            Set foundCell = wsOverview.Range("K1:K200").Find(What:="*", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not foundCell Is Nothing Then
                Set targetCell = foundCell.Offset(2, -1) ' Set the target to the next row
                ' Copy the original merged cell if it exists
            If originalMergedRange.MergeCells Then
                originalMergedRange.Copy
                With targetCell
                    .PasteSpecial Paste:=xlPasteAllUsingSourceTheme  ' Retains source theme formatting
                    .MergeCells = True  ' Merge the target cell
                    .Value = "Retest #" & retestNum & " PR#" & PRnum ' Set the retest and PR number value
                    .Font.Bold = True  ' Set font to bold
                End With
            End If
            Set targetCell = targetCell.Offset(1, 0)
            Application.CutCopyMode = False
                If func = True And water = True Then
                    copyRange2.Copy
                    targetCell.PasteSpecial Paste:=xlPasteAll
                ElseIf func = True And water = False Then
                    copyRange4.Copy
                    targetCell.PasteSpecial Paste:=xlPasteAll
                ElseIf func = False And water = True Then
                    copyRange3.Copy
                    targetCell.PasteSpecial Paste:=xlPasteAll
                End If
                Application.CutCopyMode = False
            End If
        End If
    End If
End Sub

Sub Createnewsheets(ByVal wsAnal As Worksheet, ByVal wsWater As Worksheet, ByVal retestNum As Variant, ByVal func As Boolean, ByVal water As Boolean)
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim wsCopy As Worksheet
    Dim newAnalName As String
    Dim newWaterName As String
    Dim existingSheet As Worksheet
    Dim existingSheet2 As Worksheet
    Dim wsFormat As Worksheet
    Set wsFormat = wb.Sheets("Formats")
    wsFormat.Visible = xlSheetVisible
    
    wsAnal.Unprotect Password:="Op3narray"
    wsWater.Unprotect Password:="Op3narray"
    
    ' Define the new sheet name
    newAnalName = "DNA Data Retest #" & retestNum
    newWaterName = "NTC Data Retest #" & retestNum

    ' Check if the sheet with the new name already exists
    On Error Resume Next
    Set existingSheet = wb.Sheets(newAnalName)
    Set existingSheet2 = wb.Sheets(newWaterName)
    On Error GoTo 0

    If Not existingSheet Is Nothing Then
        MsgBox "A sheet with the name '" & newAnalName & "' already exists.", vbExclamation
        Exit Sub
    End If
    
    If Not existingSheet2 Is Nothing Then
        MsgBox "A sheet with the name '" & newWaterName & "' already exists.", vbExclamation
        Exit Sub
    End If
    
    If func = True Then
        wsAnal.Copy After:=wb.Sheets(wb.Sheets.Count)
        Set wsCopy = wb.Sheets(wb.Sheets.Count)
        Call ClearMergedCells(wsCopy, wsCopy.Range("A3:C500"))
        Call ClearMergedCells(wsCopy, wsCopy.Range("E3:F500"))
        wsCopy.Name = newAnalName
    End If
    
    If water = True Then
        wsWater.Copy After:=wb.Sheets(wb.Sheets.Count)
        Set wsCopy = wb.Sheets(wb.Sheets.Count)
        Call ClearMergedCells(wsCopy, wsCopy.Range("A3:D500"))
        Call ClearMergedCells(wsCopy, wsCopy.Range("H3:K500"))
        'wsCopy.Range("H3:K500").ClearContents
        wsCopy.Name = newWaterName
    End If
    wsFormat.Visible = xlSheetHidden
    wsAnal.Protect Password:="Op3narray"
    wsWater.Protect Password:="Op3narray"
End Sub

Sub ClearMergedCells(ws As Worksheet, rng As Range)
    ' Clear the contents of the merged cells in the specified range
    Dim cell As Range
    Dim mergedArea As Range

    ' Loop through each cell in the specified range
    For Each cell In rng
        If cell.MergeCells Then
            Set mergedArea = cell.MergeArea ' Get the entire merged area
            If mergedArea.Address = cell.Address Then ' Check if it's the top-left cell of the merged area
                mergedArea.ClearContents ' Clear the contents of the entire merged area
            End If
        End If
    Next cell
End Sub


Sub ExportData(ByRef ws As Worksheet, ByRef wsAnal As Worksheet, ByRef wsWater As Worksheet, ByRef wsOverview As Worksheet)
    Dim headerCell As Range
    Dim secondheaderCell As Range
    Dim firstAssayCell As Range
    Dim secondAssayCell As Range
    Dim currentCell As Range
    Dim lastNonEmptyCell As Range
    Dim copyRange As Range
    Dim targetCell As Range
    Dim lastRow As Long
    Dim rng As Range
    Dim i As Long
    Dim waterLastRow As Long
    Dim valueCounts As Object ' For hashmap
    Dim currentValue As Variant
    Dim failureFound As Boolean
    failureFound = False
    Dim rowNum As Long
    rowNum = 3
    wsAnal.Unprotect Password:="Op3narray"
    wsWater.Unprotect Password:="Op3narray"
    Dim searchRange As Range
    Set searchRange = wsOverview.Cells
    searchText = "Functional Test Table"
    Dim count1 As Integer
    count1 = 0
    Set foundCell = searchRange.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Dim insertRange As Range
    Set insertRange = wsOverview.Range(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), _
                                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 6))
    
    Set valueCounts = CreateObject("Scripting.Dictionary")

    ' Search for the "Assay ID" header in the worksheet
    Set headerCell = ws.Cells.Find(What:="Assay ID", LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Search for the "Well" header in the worksheet
    Set secondheaderCell = ws.Cells.Find(What:="Well", LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the "Assay ID" and "Well" headers were found
    If Not headerCell Is Nothing And Not secondheaderCell Is Nothing Then
        ' Get the first Assay ID cell under the "Assay ID" header
        Set firstAssayCell = headerCell.Offset(1, 0)
        
        ' Get the first cell under the "Well" header
        Set secondAssayCell = secondheaderCell.Offset(1, 0)
        
        ' Find the last non-empty row by starting from firstAssayCell and going down until a blank cell is found
        lastRow = firstAssayCell.Row
        Do While Not IsEmpty(ws.Cells(lastRow, "A").Value)
            lastRow = lastRow + 1
        Loop
        lastRow = lastRow - 1 ' Adjust lastRow to the last non-empty cell

        ' Define the range to include rows from firstAssayCell to lastNonEmptyCell
        Set rng = ws.Range("A" & firstAssayCell.Row & ":E" & lastRow)
        
        ' Sort the range based on column A
        rng.Sort Key1:=ws.Range("A" & firstAssayCell.Row), Order1:=xlAscending, Header:=xlNo
        
        ' Recalculate the range after sorting
        Set rng = ws.Range("A" & firstAssayCell.Row & ":G" & lastRow)
        
        ' Start from row 3 in wsWater
        waterLastRow = wsWater.Cells(3, 1).End(xlUp).Row
        If waterLastRow < 3 Then waterLastRow = 3
        
        ' Loop through the rows and delete rows with "NTC" in column B
        For i = rng.Rows.Count To 1 Step -1
            If rng.Cells(i, 2).Value = "NTC" Then
                ' Extract values from columns 1, 4, 5, and 6
                wsWater.Cells(waterLastRow, 8).Value = rng.Cells(i, 1).Value  ' Column 1 (A)
                wsWater.Cells(waterLastRow, 9).Value = rng.Cells(i, 4).Value  ' Column 4 (D)
                wsWater.Cells(waterLastRow, 10).Value = rng.Cells(i, 5).Value ' Column 5 (E)
                wsWater.Cells(waterLastRow, 11).Value = rng.Cells(i, 6).Value ' Column 6 (F)
                If wsWater.Cells(waterLastRow, 9).Value > 0.5 And wsWater.Cells(waterLastRow, 9).Value > wsWater.Cells(waterLastRow, 10).Value Then
                    insertRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    wsOverview.Cells(foundCell.Row + 1, foundCell.Column).Copy
                    wsOverview.Range(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).PasteSpecial Paste:=xlPasteFormats
                    Union(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).Locked = False
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 1).Value = wsWater.Cells(waterLastRow, 8).Value
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 2) = "NTC"
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 3) = wsWater.Cells(waterLastRow, 9).Value
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 4) = wsWater.Cells(waterLastRow, 11).Value
                    ' Track occurrences in a hashmap
                    currentValue = wsWater.Cells(waterLastRow, 8)
                    If Not valueCounts.Exists(currentValue) Then
                        valueCounts.Add currentValue, 1 ' Add new entry
                    Else
                        valueCounts(currentValue) = valueCounts(currentValue) + 1 ' Increment existing entry
                    End If
                    count1 = count1 + 1
                    Application.CutCopyMode = False
                ElseIf wsWater.Cells(waterLastRow, 10).Value > 0.5 Then
                    insertRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    wsOverview.Cells(foundCell.Row + 1, foundCell.Column).Copy
                    wsOverview.Range(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).PasteSpecial Paste:=xlPasteFormats
                    Union(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).Locked = False
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 1).Value = wsWater.Cells(waterLastRow, 8).Value
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 2) = "NTC"
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 3) = wsWater.Cells(waterLastRow, 10).Value
                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 4) = wsWater.Cells(waterLastRow, 11).Value
                    ' Track occurrences in a hashmap
                    currentValue = wsWater.Cells(waterLastRow, 8)
                    If Not valueCounts.Exists(currentValue) Then
                        valueCounts.Add currentValue, 1 ' Add new entry
                    Else
                        valueCounts(currentValue) = valueCounts(currentValue) + 1 ' Increment existing entry
                    End If
                    count1 = count1 + 1
                    Application.CutCopyMode = False
                End If
                ' Increment waterLastRow to move to the next row in wsWater
                waterLastRow = waterLastRow + 1

                ' Delete the row
                rng.Rows(i).Delete
            End If
        Next i
        
        ' Recalculate lastRow after deletion
        lastRow = firstAssayCell.Row
        Do While Not IsEmpty(ws.Cells(lastRow, "A").Value)
            lastRow = lastRow + 1
        Loop
        lastRow = lastRow - 1 ' Adjust lastRow to the last non-empty cell after deletion

        ' Offset lastNonEmptyCell by 4 columns to get the range ending in column D
        Set lastNonEmptyCell = ws.Cells(lastRow, "A").Offset(0, 2)
        
        ' Define the range to copy from firstAssayCell to lastNonEmptyCell
        Set copyRange = ws.Range(firstAssayCell, lastNonEmptyCell)
        
        ' Define the destination cell in wsAnal (A3)
        Set targetCell = wsAnal.Range("A3")
        
        ' Copy the range and paste it to the destination cell
        copyRange.Copy
        targetCell.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        ' Offset lastNonEmptyCell by 4 columns to get the range ending in column E
        Set lastNonEmptyCell = ws.Cells(lastRow, "F").Offset(0, 1)
        
        ' Define the range to copy from secondAssayCell to lastNonEmptyCell
        Set copyRange = ws.Range(secondAssayCell, lastNonEmptyCell)
        
        ' Define the destination cell in wsAnal (E3)
        Set targetCell = wsAnal.Range("E3")
        
        ' Copy the range and paste it to the destination cell
        copyRange.Copy
        targetCell.PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        
        Do While wsAnal.Cells(rowNum, 8).Value <> "" ' 8 refers to column H
        ' Check if the cell equals "FAIL"
        If wsAnal.Cells(rowNum, 8).Value = "FAIL" Then
            insertRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            wsOverview.Cells(foundCell.Row + 1, foundCell.Column).Copy
            wsOverview.Range(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).PasteSpecial Paste:=xlPasteFormats
            Union(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).Locked = False
            wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 1).Value = wsAnal.Cells(rowNum, 1).Value
            wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 2).Value = wsAnal.Cells(rowNum, 2).Value
            wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 3).Value = wsAnal.Cells(rowNum, 3).Value & " Expected:" & wsAnal.Cells(rowNum, 4).Value
            wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 4).Value = wsAnal.Cells(rowNum, 5).Value
            ' Track occurrences in a hashmap
                currentValue = wsWater.Cells(waterLastRow, 8)
                If Not valueCounts.Exists(currentValue) Then
                    valueCounts.Add currentValue, 1 ' Add new entry
                Else
                    valueCounts(currentValue) = valueCounts(currentValue) + 1 ' Increment existing entry
                End If
                count1 = count1 + 1
                Application.CutCopyMode = False
        End If
        rowNum = rowNum + 1
        Loop
        For Each Key In valueCounts.Keys
            If valueCounts(Key) >= 3 Then
                failureFound = True
                Exit For
            End If
            Next Key
        
            If count1 >= 6 Then
                failureFound = True
            End If
        
            If failureFound Then
                wsOverview.Cells(foundCell.Row + 1, foundCell.Column + 6).Value = "Fail"
            Else
                wsOverview.Cells(foundCell.Row + 1, foundCell.Column + 6).Value = "Pass"
            End If
    End If
    wsOverview.Cells(foundCell.Row - 2, foundCell.Column + 1).Value = wsAnal.Cells(3, 9).Value
    wsOverview.Cells(foundCell.Row - 2, foundCell.Column + 2).Value = wsAnal.Cells(3, 10).Value
    wsOverview.Cells(foundCell.Row - 2, foundCell.Column + 3).Value = wsAnal.Cells(3, 11).Value
    wsOverview.Cells(foundCell.Row - 2, foundCell.Column + 5).Value = wsAnal.Cells(3, 12).Value
    wsAnal.Protect Password:="Op3narray"
    wsWater.Protect Password:="Op3narray"
End Sub


Sub WaterData(ByRef ws As Worksheet, ByRef wsWater As Worksheet, ByRef wsOverview As Worksheet, target As String, retestNum As Variant)
    Dim headerCell As Range
    Dim firstAssayCell As Range
    Dim lastRow As Long
    Dim rng As Range
    Dim i As Long
    Dim start As Integer
    Dim searchText As String
    Dim foundCell As Range
    Dim count1 As Integer
    count1 = 0
    Dim copyRange As Range
    Dim targetCell As Range
    Dim valueCounts As Object ' For hashmap
    Dim currentValue As Variant
    Dim failureFound As Boolean
    failureFound = False
    wsWater.Unprotect Password:="Op3narray"
    If retestNum = "" Then
        Set firstCell = wsOverview.Cells.Find(What:="Initial Test", LookIn:=xlValues, LookAt:=xlPart)
    Else
        Set firstCell = wsOverview.Cells.Find(What:="Retest #" & retestNum, LookIn:=xlValues, LookAt:=xlPart)
    End If
    Dim searchRange As Range
    Set searchRange = wsOverview.Range(wsOverview.Cells(firstCell.Row, firstCell.Column), _
                                    wsOverview.Cells(firstCell.Row + 500, firstCell.Column + 6))
    searchText = "Water NTC Test Table"
    Set foundCell = searchRange.Find(What:=searchText, LookIn:=xlValues, LookAt:=xlPart)
    Dim insertRange As Range
    Set insertRange = wsOverview.Range(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), _
                                    wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 6))
                            
    Set valueCounts = CreateObject("Scripting.Dictionary")

    ' Search for the "Assay ID" header in the worksheet
    Set headerCell = ws.Cells.Find(What:="Assay ID", LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the "Assay ID" header was found
    If Not headerCell Is Nothing Then
        ' Get the first Assay ID cell under the header
        Set firstAssayCell = headerCell.Offset(1, 0)
        
        ' Initialize currentCell as the firstAssayCell
        Set currentCell = firstAssayCell
        
        ' Find the last non-empty row by starting from firstAssayCell and going down until a blank cell is found
        lastRow = firstAssayCell.Row
        Do While Not IsEmpty(ws.Cells(lastRow, "A").Value)
            lastRow = lastRow + 1
        Loop
        lastRow = lastRow - 1 ' Adjust lastRow to the last non-empty cell

        ' Define the range to include rows from firstAssayCell to lastNonEmptyCell
        Set rng = ws.Range("A" & firstAssayCell.Row & ":D" & lastRow)
        
        ' Sort the range based on column A
        rng.Sort Key1:=ws.Range("A" & firstAssayCell.Row), Order1:=xlAscending, Header:=xlNo
        
        ' Define the range to include rows from firstAssayCell to lastNonEmptyCell after sorting
        Set rng = ws.Range("A" & firstAssayCell.Row & ":D" & lastRow)
        start = 3
        
        For i = rng.Rows.Count To 1 Step -1
            wsWater.Cells(start, 1).Value = rng.Cells(i, 1).Value
            wsWater.Cells(start, 2).Value = rng.Cells(i, 2).Value
            wsWater.Cells(start, 3).Value = rng.Cells(i, 3).Value
            wsWater.Cells(start, 4).Value = rng.Cells(i, 4).Value
            
            If wsWater.Cells(start, 2).Value > 0.5 Or wsWater.Cells(start, 3).Value > 0.5 Then
                insertRange.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                wsOverview.Cells(foundCell.Row + 1, foundCell.Column).Copy
                wsOverview.Range(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).PasteSpecial Paste:=xlPasteFormats
                Union(wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column), wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 5)).Locked = False
                wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 1).Value = rng.Cells(i, 1).Value
                wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 2) = rng.Cells(i, 2).Value
                wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 3) = rng.Cells(i, 3).Value
                wsOverview.Cells(foundCell.Row + 2 + count1, foundCell.Column + 4) = rng.Cells(i, 4).Value
                
                ' Track occurrences in a hashmap
                currentValue = rng.Cells(i, 1).Value
                If Not valueCounts.Exists(currentValue) Then
                    valueCounts.Add currentValue, 1 ' Add new entry
                Else
                    valueCounts(currentValue) = valueCounts(currentValue) + 1 ' Increment existing entry
                End If
                
                count1 = count1 + 1
                Application.CutCopyMode = False
            End If
            start = start + 1
        Next i
        
        For Each Key In valueCounts.Keys
            If valueCounts(Key) >= 3 Then
                failureFound = True
                Exit For
            End If
        Next Key
        
        If count1 >= 6 Then
            failureFound = True
        End If
        
        If failureFound Then
            wsOverview.Cells(foundCell.Row + 1, foundCell.Column + 6).Value = "Fail"
        Else
            wsOverview.Cells(foundCell.Row + 1, foundCell.Column + 6).Value = "Pass"
        End If
    End If
    wsWater.Protect Password:="Op3narray"
End Sub

