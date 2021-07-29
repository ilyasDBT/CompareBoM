Sub CompareBom(fullFileName As String)
    'this program matches the item number and compares the revision
    'if the revision has changed, this program inserts a check mark in the changed column
    'source: old bom file
    'destination: new bom file
    'to add: conditional statement for sheet metal
    Application.ScreenUpdating = False
    Dim fileName As String
    fileName = Right(fullFileName, Len(fullFileName) - InStrRev(fullFileName, "\"))
    
    'set the source and destination workbooks and sheet
    Dim wsDest As Worksheet
    Set wsDest = Application.ActiveSheet

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim check As Boolean
    check = IsWorkbookOpen(fileName)
    If check = False Then
        On Error GoTo 10
        Set wbSource = Workbooks.Open(fullFileName)
        On Error GoTo 0
    Else
        Set wbSource = Workbooks(fileName)
    End If
    Set wsSource = wbSource.Worksheets(1)
    
    'find the property columns
    
    'item number in both old and new BOM
    Dim oldItemNo As Range
    Set oldItemNo = wsSource.UsedRange.Find("Item Number", , xlValues, xlWhole)
    If oldItemNo Is Nothing Then
        MsgBox "Item Number column Not found in Old BOM", vbCritical, "CompareBOM"
        wbSource.Close
        Exit Sub
    End If
    
    
    Dim itemno As Range
    Set itemno = wsDest.UsedRange.Find("Item Number", , xlValues, xlWhole)
    If itemno Is Nothing Then
        MsgBox "Item Number column Not found in New BOM", vbCritical, "CompareBOM"
        wbSource.Close
        Exit Sub
    End If
        
    'Drawing Rev in both old and new BOM
    Dim oldDrawingRev As Range
    Set oldDrawingRev = wsSource.UsedRange.Find("Drawing Rev", , xlValues, xlWhole)
    If oldDrawingRev Is Nothing Then
        MsgBox "Drawing Rev column Not found in Old BOM", vbCritical, "CompareBOM"
        wbSource.Close
        Exit Sub
    End If
    
    
    Dim drawingRev As Range
    Set drawingRev = wsDest.UsedRange.Find("Drawing Rev", , xlValues, xlWhole)
    If drawingRev Is Nothing Then
        MsgBox "Drawing Rev column Not found in New BOM", vbCritical, "CompareBOM"
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
    
    'Changed column
    Dim changed As Range
    Set changed = wsDest.UsedRange.Find("Changed", , xlValues, xlWhole)
    If changed Is Nothing Then
        Set changed = wsDest.Cells(itemno.Row, wsDest.UsedRange.Columns.Count + 1)
        changed.Value = "Changed"
    End If
    'qtyxparent
        Dim qtyxparent As Range
    Set qtyxparent = wsSource.UsedRange.Find("QTYxParent", , xlValues, xlWhole)
    If qtyxparent Is Nothing Then
        MsgBox "QTYxParent column Not found in Old BOM", vbCritical, "CompareBOM"
        wbSource.Close
        Exit Sub
    End If
    'Original Value Qty column
    Dim OriginalValue As Range
    Set OriginalValue = wsDest.UsedRange.Find("Original Value", , xlValues, xlWhole)
    If OriginalValue Is Nothing Then
        Set OriginalValue = wsDest.Cells(itemno.Row, wsDest.UsedRange.Columns.Count + 1)
        OriginalValue.Value = "Original Value"
    End If
    'start compare and copy properties
    Dim newitem As Boolean
    For j = itemno.Row + 1 To wsDest.UsedRange.Rows.Count
        newitem = True
        For i = oldItemNo.Row + 1 To wsSource.UsedRange.Rows.Count
            If wsDest.Cells(j, itemno.Column) = wsSource.Cells(i, oldItemNo.Column) Then
                If CStr(wsDest.Cells(j, drawingRev.Column)) <> CStr(wsSource.Cells(i, oldDrawingRev.Column)) Then
                    wsDest.Cells(j, changed.Column).Value = ChrW(&H2713)
                    newitem = False
                    If wsDest.Cells(j, itemCategoryDest.Column).Value <> "R" And wsSource.Cells(i, itemCategorySource.Column).Value <> "R" Then
                        wsDest.Cells(j, OriginalValue.Column).Value = wsSource.Cells(i, qtyxparent.Column).Value
                    End If
                Else
                    wsDest.Cells(j, changed.Column).Value = ""
                    newitem = False
                    If wsDest.Cells(j, itemCategoryDest.Column).Value <> "R" And wsSource.Cells(i, itemCategorySource.Column).Value <> "R" Then
                        wsDest.Cells(j, OriginalValue.Column).Value = ""
                    End If
                End If
            End If
        Next i
        If newitem Then
            wsDest.Cells(j, changed.Column).Value = ChrW(&H2713)
            If wsDest.Cells(j, itemCategoryDest.Column).Value <> "R" And wsSource.Cells(i, itemCategorySource.Column).Value <> "R" Then
                wsDest.Cells(j, OriginalValue.Column).Value = 0
            End If
        End If
    Next j

    Call CreateBomPath(wsDest)
    Call CreateBomPath(wsSource)
    Dim bompathheader As Range
    Set bompathheader = wsSource.UsedRange.Find("BOM Path", , xlValues, xlWhole)
    Dim parentbompath As Range
    Dim bompath As String
    For i = oldItemNo.Row + 1 To wsSource.UsedRange.Rows.Count
        bompath = wsSource.Cells(i, bompathheader.Column).Value
        If wsDest.UsedRange.Find(bompath, , xlValues, xlWhole) Is Nothing Then
            wsSource.Cells(i, 1).EntireRow.Copy
            Set parentbompath = wsDest.UsedRange.Find(parentlevel(bompath), , xlValues, xlWhole)
            wsDest.Cells(parentbompath.Row + 1, 1).Insert
            wsDest.Cells(parentbompath.Row + 1, 1).EntireRow.Font.Strikethrough = True
        End If
    Next i
    wsDest.UsedRange.Find("BOM Path", , xlValues, xlWhole).EntireColumn.Delete
    wbSource.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Done", , "Compare BOM"
    
    GoTo 11
    
10  MsgBox "File does not exist. Please browse to an existing file.", , "CompareBom"
    End
    
11
End Sub

Function parentlevel(bompath As String)

    If InStr(bompath, ".") > 0 Then
        parentlevel = Left(bompath, InStrRev(bompath, ".") - 1)
    Else
        parentlevel = ""
    End If
    
End Function
