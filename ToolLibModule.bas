Attribute VB_Name = "ToolLibModule"
' Row and column usage on the Main page (LibList)
Const colLibName = 1
Const colPath = 2
Const colBrowse = 3
Const colDate = 4
Const colAskLoad = 5
Const colSourceOption = 6
Const colUpdate = 7
Const colLoad = 8
Const colSource = 9
Const rowFirstLib = 5
' Row and column usage on Rules page
Const colFirstRule = 2
Const rowTitle = 1
Const rowIsTsv = 2
Const rowHeaderRow = 3
Const rowMapStart = 4
Const sheetFirstLib = 2
' Messages
Const msgTitle = "Tool Library"
Const msgNoSource = "You must select a library to act as the source for the update."
Const msgSameSource = "You must choose a different source for the update."
Const msgNoPath = "You must enter or browse for the path to the library file."

Sub BuildLibLine(name As String, rowNum As Integer)
    Dim btn As Button, chk As CheckBox, opt As OptionButton
    
    Set cell = LibList.Cells(rowNum, colLibName)
    If cell.Value <> name Then
        cell.Value = name
        LibList.Cells(rowNum, colPath).Value = Empty
        LibList.Cells(rowNum, colDate).Value = Empty
    End If
    'Create browse button
    Set cell = LibList.Cells(rowNum, colBrowse)
    With cell
        Set btn = LibList.Buttons.Add(.Left, .Top, .Width, .Height)
        .Value = " "
    End With
    With btn
        .Caption = "..."
        .OnAction = "'Browse " & rowNum & "'"
    End With
    'Create load checkbox
    Set cell = LibList.Cells(rowNum, colAskLoad)
    With cell
        Set chk = LibList.CheckBoxes.Add(.Left, .Top, .Width, .Height)
    End With
    With chk
        .Caption = ""
        .LinkedCell = LibList.Cells(rowNum, colLoad).Address
    End With
    'Create source option button
    Set cell = LibList.Cells(rowNum, colSourceOption)
    With cell
        Set opt = LibList.OptionButtons.Add(.Left, .Top, .Width, .Height)
    End With
    With opt
        .Caption = ""
        .LinkedCell = LibList.Cells(rowFirstLib, colSource).Address
    End With
    'Create update button
    Set cell = LibList.Cells(rowNum, colUpdate)
    With cell
        Set btn = LibList.Buttons.Add(.Left, .Top, .Width, .Height)
    End With
    With btn
        .Caption = "Update Library"
        .OnAction = "'Update " & rowNum & "'"
    End With
End Sub

Sub BuildList()
    Dim lib As String, colCur As Integer, rowCur As Integer
    LibList.Buttons.Delete
    LibList.CheckBoxes.Delete
    LibList.OptionButtons.Delete
    colCur = colFirstRule
    rowCur = rowFirstLib
    Do
        lib = Rules.Cells(rowTitle, colCur).Value
        If IsEmpty(lib) Or lib = "" Then Exit Do
        BuildLibLine lib, rowCur
        colCur = colCur + 1
        rowCur = rowCur + 1
    Loop
End Sub

Sub Browse(rowNum As Integer)
    Dim name As String, title As String, filter As String
    name = LibList.Cells(rowNum, colLibName)
    title = "Select " + name + " tool library"
    If IsTsv(name) Then
        filter = "TSV Files (*.tsv),*.tsv"
    Else
        filter = "CSV Files (*.csv),*.csv"
    End If
    
    fName = Application.GetOpenFilename(FileFilter:=filter, title:=title)
    If (fName <> False) Then LibList.Cells(rowNum, colPath).Value = fName
End Sub

Sub LoadSelected()
    Dim rowCur As Integer, name As String
    rowCur = rowFirstLib
    Do
        fUse = LibList.Cells(rowCur, colLibName).Value
        If IsEmpty(fUse) Or fUse = "" Then Exit Do
        fUse = LibList.Cells(rowCur, colLoad).Value
        If fUse = True Then
            name = LibList.Cells(rowCur, colLibName).Value
            fName = LibList.Cells(rowCur, colPath).Value
            If IsEmpty(fName) Or fName = "" Then
                MsgBox msgNoPath, vbOKOnly Or vbCritical, msgTitle
                Exit Do
            End If
            LoadLibrary name, fName
            LibList.Cells(rowCur, colDate).Value = FileDateTime(fName)
        End If
        rowCur = rowCur + 1
    Loop
    LibList.Activate
End Sub

Sub LoadLibrary(name As String, fName As Variant)
    Dim newBook As Workbook, tsv As Boolean
    i = Workbooks.Count
    tsv = IsTsv(name)
    Workbooks.OpenText fName, xlWindows, DataType:=xlDelimited, Tab:=tsv, Comma:=Not tsv
    Set newBook = Workbooks.Item(i + 1)
    ReplaceSheet name, newBook
End Sub

Function GetSource() As Integer
    srcIndex = LibList.Cells(rowFirstLib, colSource)
    If IsEmpty(srcIndex) Or srcIndex = "" Or srcIndex = 0 Then
        MsgBox msgNoSource, vbOKOnly Or vbCritical, msgTitle
        End
    End If
    GetSource = rowFirstLib + srcIndex - 1
End Function

Sub Update(rowNum As Integer)
    Dim srcRow As Integer
    srcRow = GetSource()
    If srcRow = rowNum Then
        MsgBox msgSameSource, vbOKOnly Or vbCritical, msgTitle
        End
    End If
    UpdateLibrary srcRow, rowNum
    LibList.Activate
End Sub

Sub UpdateSelected()
    Dim destRow As Integer, srcRow As Integer
    srcRow = GetSource()
    destRow = rowFirstLib
    Do
        fUse = LibList.Cells(destRow, colLibName).Value
        If IsEmpty(fUse) Or fUse = "" Then Exit Do
        fUse = LibList.Cells(destRow, colLoad).Value
        If fUse = True And srcRow <> destRow Then
            UpdateLibrary srcRow, destRow
        End If
        destRow = destRow + 1
    Loop
    LibList.Activate
End Sub

Sub UpdateLibrary(srcRow As Integer, destRow As Integer)
    Dim src As String, dest As String
    src = LibList.Cells(srcRow, colLibName).Value
    dest = LibList.Cells(destRow, colLibName).Value
    BuildTable src, dest
    fName = LibList.Cells(destRow, colPath).Value
    SaveLibrary dest, fName
    LibList.Cells(destRow, colDate).Value = FileDateTime(fName)
End Sub

Sub SaveLibrary(name As String, fName As Variant)
    Set sheet = Application.ThisWorkbook.Worksheets(name)
    ' copy to new workbook, so we don't rename our own file
    sheet.Copy
    Set sheet = Application.ActiveWorkbook
    Application.DisplayAlerts = False
    sheet.SaveAs fName, xlCSV
    sheet.Close
End Sub

Function IsTsv(name As String) As Boolean
    colRules = Rules.Rows(rowTitle).Find(name).Column
    IsTsv = Rules.Cells(rowIsTsv, colRules).Value
End Function

Function SaveColumnWidths(sheet As Worksheet)
    cnt = sheet.UsedRange.Columns.Count
    ReDim arWidths(cnt) As Double
    For i = 1 To cnt
        arWidths(i - 1) = sheet.Columns(i).ColumnWidth
    Next
    SaveColumnWidths = arWidths
End Function

Sub RestoreColumnWidths(sheet As Worksheet, arWidths)
    If IsEmpty(arWidths) Then Exit Sub
    For i = 1 To UBound(arWidths) + 1
        sheet.Columns(i).ColumnWidth = arWidths(i - 1)
    Next
End Sub

Sub SortLibrary(sheet As Worksheet)
    col = Rules.Rows(rowTitle).Find(sheet.name).Column
    head = Rules.Cells(rowHeaderRow, col).Value
    col = Rules.Cells(rowMapStart, col).Value   'name of column
    col = sheet.Rows(head).Find(col, Lookat:=xlWhole).Column 'column number
    rowLast = sheet.UsedRange.Rows.Count
    colLast = sheet.UsedRange.Columns.Count
    Set area = sheet.Range(sheet.Cells(head + 1, 1), sheet.Cells(rowLast, colLast))
    area.Sort sheet.Cells(head + 1, col)
End Sub

Sub ReplaceSheet(sheetName As String, newBook As Workbook)
    Dim old As Worksheet, sheet As Worksheet
    Set curBook = Application.ThisWorkbook
    Set sheet = newBook.ActiveSheet
    On Error Resume Next
    Set old = curBook.Worksheets(sheetName)
    On Error GoTo 0
    If old Is Nothing Then
        Set old = curBook.Worksheets(sheetFirstLib)
        i = sheetFirstLib
    Else
        i = old.Index
        arWidths = SaveColumnWidths(old)
    End If
    sheet.Copy Before:=old
    newBook.Close False
    Application.DisplayAlerts = False
    If old.name = sheetName Then old.Delete
    Set sheet = curBook.Worksheets(i)
    sheet.name = sheetName
    RestoreColumnWidths sheet, arWidths
    SortLibrary sheet
End Sub

Function BuildMap(sheet As Worksheet, colRules As Integer) As Collection
    Dim map As New Collection
    Dim label As String
    Set rowHeader = sheet.Rows(Rules.Cells(rowHeaderRow, colRules).Value)
    rowCur = rowMapStart
    Do
        label = Rules.Cells(rowCur, colRules).Value
        If IsEmpty(label) Or label = "" Then Exit Do
        Set pos = rowHeader.Find(label, Lookat:=xlWhole)
        If pos Is Nothing Then
            map.Add 0
        Else
            map.Add pos.Column
        End If
        rowCur = rowCur + 1
    Loop
    Set BuildMap = map
End Function

Sub BuildTable(src As String, dst As String)
    Dim srcSheet As Worksheet, dstSheet As Worksheet
    Dim srcMap As Collection, dstMap As Collection
    Dim srcRuleCol As Integer, dstRuleCol As Integer
    Set curBook = Application.ThisWorkbook
    Set Rng = Rules.Rows(rowTitle)
    srcRuleCol = Rng.Find(src).Column
    Set srcSheet = curBook.Worksheets(src)
    srcEndRow = srcSheet.UsedRange.Rows.Count
    Set srcMap = BuildMap(srcSheet, srcRuleCol)
    
    Set dstSheet = curBook.Worksheets(dst)
    dstRuleCol = Rng.Find(dst).Column
    dstStartRow = Rules.Cells(rowHeaderRow, dstRuleCol).Value + 1
    dstEndRow = dstSheet.UsedRange.Rows.Count
    Set dstMap = BuildMap(dstSheet, dstRuleCol)
    Set dstSlotCol = dstSheet.Range(dstSheet.Cells(dstStartRow, dstMap(1)), dstSheet.Cells(dstEndRow, dstMap(1)))
    
    ' Visit each source row
    srcCurRow = Rules.Cells(rowHeaderRow, srcRuleCol).Value
    Do While srcCurRow < srcEndRow
        srcCurRow = srcCurRow + 1
        slot = srcSheet.Cells(srcCurRow, srcMap(1)).Value
        Set dstSlotRow = dstSlotCol.Find(slot)
        If dstSlotRow Is Nothing Then
            ' not found, add to end
            dstEndRow = dstEndRow + 1
            dstSlotRow = dstEndRow
            dstSheet.Cells(dstSlotRow, dstMap(1)).Value = slot
            isNewRow = True
        Else
            dstSlotRow = dstSlotRow.row
            isNewRow = False
        End If
        
        ' Visit each remaining map entry and copy data
        isRowValid = False
        For iMap = 2 To dstMap.Count
            If iMap > srcMap.Count Then Exit For
            srcCol = srcMap(iMap)
            dstCol = dstMap(iMap)
            If srcCol <> 0 And dstCol <> 0 Then
                dat = srcSheet.Cells(srcCurRow, srcCol).Value
                If Not (IsEmpty(dat) Or dat = "" Or dat = 0) Then isRowValid = True
                dstSheet.Cells(dstSlotRow, dstCol).Value = dat
            End If
        Next
        ' Don't add empty row
        If isNewRow And Not isRowValid Then
            dstSheet.Rows(dstSlotRow).EntireRow.Delete
            dstEndRow = dstEndRow - 1
        End If
    Loop
    SortLibrary dstSheet
End Sub
