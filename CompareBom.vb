Sub CompareBom(fullFileName As String)
        'this program matches the item number and compares the records values for the defined columns for comparison
        'if the revision has changed, this program writes into the changed column and previous value column
        'source: old bom file
        'destination: new bom file
    
        'set the source and destination workbooks and sheet
        Dim wsDest As Worksheet
        Set wsDest = Application.ActiveSheet
    
        Application.ScreenUpdating = False 'stop excel screen updating to prevent flashing of screen while code is running
        Dim fileName As String
        fileName = Right(fullFileName, Len(fullFileName) - InStrRev(fullFileName, "\")) 'get filename from filepath
        Dim wbSource As Workbook
        Dim wsSource As Worksheet
        Dim check As Boolean
        check = IsWorkbookOpen(fileName) 'check if workbook is open, function located in another module
        If check = False Then
            On Error GoTo 10
            Set wbSource = Workbooks.Open(fullFileName)
            On Error GoTo 0
        Else
            Set wbSource = Workbooks(fileName)
        End If
        Set wsSource = wbSource.Worksheets(1)
    
        'error handling for opening workbook
        GoTo 11
10      MsgBox "File does not exist. Please browse to an existing file.", , "CompareBom"
        End
11
    
        'define primary column names, item number and category, exit if doesn't exist
        Dim itemNoSource As Range
        Set itemNoSource = wsSource.UsedRange.Find("Item Number", , xlValues, xlWhole)
        If itemNoSource Is Nothing Then
            MsgBox "Item Number column Not found in Old BOM", vbCritical, "CompareBOM"
            wbSource.Close
            Exit Sub
        End If
        
        Dim itemNoDest As Range
        Set itemNoDest = wsDest.UsedRange.Find("Item Number", , xlValues, xlWhole)
        If itemNoDest Is Nothing Then
            MsgBox "Item Number column Not found in New BOM", vbCritical, "CompareBOM"
            wbSource.Close
            Exit Sub
        End If
        'Add Machine Index to match
        Dim MIndexSource As Range
        Set MIndexSource = wsSource.UsedRange.Find("Machine Index", , xlValues, xlWhole)
        If MIndexSource Is Nothing Then
            MsgBox "Machine Index column Not found in Old BOM", vbCritical, "CompareBOM"
            wbSource.Close
            Exit Sub
        End If
        
        Dim MIndexDest As Range
        Set MIndexDest = wsDest.UsedRange.Find("Machine Index", , xlValues, xlWhole)
        If MIndexDest Is Nothing Then
            MsgBox "Machine Index column Not found in New BOM", vbCritical, "CompareBOM"
            wbSource.Close
            Exit Sub
        End If
        
        Dim itemCategorySource As Range
        Set itemCategorySource = wsSource.UsedRange.Find("Item Category", , xlValues, xlWhole)
        If itemCategorySource Is Nothing Then
            MsgBox "Item Category column Not found in Old BOM", vbCritical, "CompareBOM"
            wbSource.Close
            Exit Sub
        End If
        
        Dim itemCategoryDest As Range
        Set itemCategoryDest = wsDest.UsedRange.Find("Item Category", , xlValues, xlWhole)
        If itemCategoryDest Is Nothing Then
            MsgBox "Item Category column Not found in New BOM", vbCritical, "CompareBOM"
            wbSource.Close
            Exit Sub
        End If
    
        'define output columns, changed and previous value, create if doesn't exist
        Dim changed As Range
        Set changed = wsDest.UsedRange.Find("Changed", , xlValues, xlWhole)
        If changed Is Nothing Then
            Set changed = wsDest.Cells(itemNoDest.Row, wsDest.UsedRange.Columns.Count + 1)
            changed.Value = "Changed"
        End If
    
        Dim previousValue As Range
        Set previousValue = wsDest.UsedRange.Find("Previous Value", , xlValues, xlWhole)
        If previousValue Is Nothing Then
            Set previousValue = wsDest.Cells(itemNoDest.Row, wsDest.UsedRange.Columns.Count + 1)
            previousValue.Value = "Previous Value"
        End If
    
        'define column names to compare. Please add this if required
        Dim columnNames As ArrayList 'need to add reference to mscorlib.dll
        Set columnNames = New ArrayList 'columnNames used to store the column header strings
        columnNames.Add ("Drawing Rev")
        columnNames.Add ("QTYxParent")
        columnNames.Add ("Rating")
        columnNames.Add ("Model")
        columnNames.Add ("Brand")
        
        'copy the previous line and tweak the string to add more column headers
    
        Dim columnNumbersSource As ArrayList 'to store the column numbers for source worksheet
        Set columnNumbersSource = New ArrayList
        Dim columnNumbersDest As ArrayList 'to store the column numbers for destination worksheet
        Set columnNumbersDest = New ArrayList
        
        'loop through all column names to get the corresponding column numbers
        For Each ColumnName In columnNames
            Dim columnHeaderSource As Range
            Set columnHeaderSource = wsSource.UsedRange.Find(ColumnName, , xlValues, xlWhole)
            If columnHeaderSource Is Nothing Then
                MsgBox ColumnName & " column Not found in Source BoM", vbCritical, "CompareBOM"
                wbSource.Close
                Exit Sub
            End If
            columnNumbersSource.Add (columnHeaderSource.Column)
            
            Dim columnHeaderDest As Range
            Set columnHeaderDest = wsDest.UsedRange.Find(ColumnName, , xlValues, xlWhole)
            If columnHeaderSource Is Nothing Then
                MsgBox ColumnName & " column Not found in Destination BoM", vbCritical, "CompareBOM"
                wbSource.Close
                Exit Sub
            End If
            columnNumbersDest.Add (columnHeaderDest.Column)
        Next
    
        'Compare Dest againt Source to find new and/or changed records
        For i = itemNoDest.Row + 1 To wsDest.UsedRange.Rows.Count
            If wsDest.Cells(i, itemCategoryDest.Column).Value = "R" Then GoTo 20 'skip this row if category is R
                
            Dim itemExist As Boolean
            itemExist = False 'to check if item is new
            For j = itemNoSource.Row + 1 To wsSource.UsedRange.Rows.Count
                If wsSource.Cells(j, itemCategorySource.Column).Value = "R" Then GoTo 30 'skip this row if category is R
                'Start compare item number and machine index
                If wsDest.Cells(i, itemNoDest.Column) = wsSource.Cells(j, itemNoSource.Column) And wsDest.Cells(i, MIndexDest.Column) = wsSource.Cells(j, MIndexSource.Column) Then
                    itemExist = True
    
                    Dim previousValueString As String
                    previousValueString = ""
                    For k = 0 To columnNames.Count - 1
                        If wsDest.Cells(i, columnNumbersDest(k)).Value <> wsSource.Cells(j, columnNumbersSource(k)).Value Then
                            wsDest.Cells(i, columnNumbersDest(k)).Interior.Color = vbYellow 'highlight changed values
                            'store the values of the column to the previous value string
                            previousValueString = previousValueString + columnNames(k) + ":" + CStr(wsSource.Cells(j, columnNumbersSource(k)).Value) + ", "
                        Else
                            wsDest.Cells(i, columnNumbersDest(k)).Interior.Color = xlNone 'unhighlight cell
                        End If
                    Next k
                                    
                    If previousValueString <> "" Then 'if there were any differences, the previous value string will not be empty
                        wsDest.Cells(i, changed.Column).Value = ChrW(&H2713) 'add a tick mark to indicate changed
                        previousValueString = Left(previousValueString, Len(previousValueString) - 2) 'remove final comma and space
                    Else
                        wsDest.Cells(i, changed.Column).Value = ""
                    End If
                    wsDest.Cells(i, previousValue.Column).Value = previousValueString
                    
                    Exit For 'Exit loop of Source if a match is already found and process. no need to loop through Source further.
                End If
30
            Next j
            
            'if item does not exist, it must be a new item, mark as changed and set previousValue as new
            If Not itemExist Then
                wsDest.Cells(i, changed.Column).Value = ChrW(&H2713)
                wsDest.Cells(i, previousValue.Column).Value = "New"
            End If
20
        Next i
    
        'Compare Source against Dest to find removed records
        Call CreateBomPath(wsDest) 'create BOM Path Column. function located in another module
        Call CreateBomPath(wsSource)
        Dim bomPathSource As Range
        Set bomPathSource = wsSource.UsedRange.Find("BOM Path", , xlValues, xlWhole)
        Dim bomPathDest As Range
        Set bomPathDest = wsDest.UsedRange.Find("BOM Path", , xlValues, xlWhole)
        Dim bomPath As String
        For j = itemNoSource.Row + 1 To wsSource.UsedRange.Rows.Count 'loop through source rows
            bomPath = wsSource.Cells(j, bomPathSource.Column).Value
            'look for the BOM Path in Destination worksheet BOM Path column
            If wsDest.Columns(bomPathDest.Column).Find(bomPath, , xlValues, xlWhole) Is Nothing Then
                wsSource.Cells(j, 1).EntireRow.Copy 'copy the whole row
                
                Dim parentBomPath As Range
                Set parentBomPath = wsDest.UsedRange.Find(parentLevel(bomPath), , xlValues, xlWhole) 'find the item's parent based on its BOM path
                wsDest.Cells(parentBomPath.Row + 1, 1).Insert 'paste in Destination worksheet under the parent
                wsDest.Cells(parentBomPath.Row + 1, 1).EntireRow.Font.Strikethrough = True 'mark as strikethrough to indicate a removed record
            End If
        Next j
        
        wsDest.UsedRange.Find("BOM Path", , xlValues, xlWhole).EntireColumn.Delete 'delete BOM path column
        wbSource.Close SaveChanges:=False 'close source excel file without saving
        Application.ScreenUpdating = True
        MsgBox "Done", , "Compare BOM"
    End Sub
    
    Function parentLevel(bomPath As String)
        If InStr(bomPath, ".") > 0 Then
            parentLevel = Left(bomPath, InStrRev(bomPath, ".") - 1)
        Else
            parentLevel = ""
        End If
    End Function
