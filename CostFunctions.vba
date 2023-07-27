Function FindTotalDelta(sheetName As String)
    Dim ws As Worksheet
    Dim columnRange As Range
    Dim cell As Range
    Dim stopCell As Range
    Dim hValue As Variant
    Dim totalCount As Single
    totalCount = 0
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set columnRange = ws.Range("H1:H" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
    
    Dim callingRow As Long
    callingRow = Application.Caller.Row
    Set stopCell = ws.Cells(callingRow + 1, "H")
    
    'Set stopCell = ws.Range("H12") 'For debug purposes only
    
    
    'get the cost code of current row
    hValue = stopCell.Offset(-1, 0).Value
    On Error Resume Next
    
    ' Loop through each cell in the column
    For Each cell In columnRange
        If cell.Address = stopCell.Address Then
            Exit For ' Exit the loop when the stop cell is reached
        End If
        If IsEmpty(cell) Then
            ' Skip the instructions inside the loop and move to the next iteration
            GoTo ContinueLoop
        End If
        If Err.Number <> 0 Then
            Err.Clear
            Resume Next
        End If
        If cell.Value = hValue Then
            totalCount = totalCount + ws.Cells(cell.Row, "E")
        End If
        
ContinueLoop:
    Next cell
    
    FindTotalDelta = totalCount
    
    
End Function

Function TotalSpent(COST_CODE As String) As Double
    Dim myList() As Variant
    
    myList = Array("Mechanical", "Electrical", "Comms", "Track", "Traction Power", "Signals", "CMS")
    
    Dim ws As Worksheet
    Dim columnRange As Range
    Dim cell As Range
    totalCount = 0
    On Error Resume Next
    
    For Each sheetName In myList
        
        Set ws = ThisWorkbook.Worksheets(sheetName)
        If ws Is Nothing Then
            
            ' If the sheet does not exist, skip to the next iteration
            GoTo ContinueLoop
        End If
        
        Set columnRange = ws.Range("H1:H" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
        
        For Each cell In columnRange
            'MsgBox sheetName & " " & cell.Value & " " & totalCount
            If IsEmpty(cell) Then
                ' Skip the instructions inside the loop and move to the next iteration
                GoTo ContinueCellLoop
            End If
            If Err.Number <> 0 Then
                Err.Clear
                Resume Next
            End If
            If cell.Value = COST_CODE Then
                totalCount = totalCount + ws.Cells(cell.Row, "E")
            End If
            
ContinueCellLoop:
        Next cell
        
ContinueLoop:
    Next sheetName
    
    TotalSpent = totalCount
End Function

Function ProcurementPercentage(sheetName As String)
    
    Dim ws As Worksheet
    Dim columnRange As Range
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set columnRange = ws.Range("N3:N" & lastRow)
    Dim ans As Double
    Dim totalProcurement As Single
    Dim notProcured As Single
    notProcured = 0
    totalProcurement = 0
    
    
    For Each cell In columnRange
        If Err.Number <> 0 Then
            Err.Clear
            Resume Next
        End If
        If IsEmpty(cell) Then
            'Count the empty cells to subtract from total
            notProcured = notProcured + 1
        End If
        
ContinueLoop:
    Next cell
    
    ans = (lastRow - 2 - notProcured) / (lastRow - 2)
    ProcurementPercentage = ans
End Function

Function DeliveryPercentage(sheetName As String)
    
    Dim ws As Worksheet
    Dim columnRange As Range
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set columnRange = ws.Range("R3:R" & lastRow)
    Dim ans As Double
    Dim totalProcurement As Single
    Dim notProcured As Single
    notProcured = 0
    totalProcurement = 0
    
    
    For Each cell In columnRange
        If Err.Number <> 0 Then
            Err.Clear
            Resume Next
        End If
        If IsEmpty(cell) Then
            'Count the empty cells to subtract from total
            notProcured = notProcured + 1
        End If
        
ContinueLoop:
    Next cell
    
    ans = (lastRow - 2 - notProcured) / (lastRow - 2)
    DeliveryPercentage = ans
End Function

Function BECalc(Etype As String)
    Dim myList() As Variant
    
    myList = Array("Mechanical", "Electrical", "Comms", "Track", "Traction Power", "Signals", "CMS")
    'myList = Array("Mechanical", "Electrical", "Comms", "Track", "Traction Power") 'Debug Tool Only
    
    Dim ws As Worksheet
    Dim columnRange As Range
    Dim row2Range As Range
    Dim ReqCol As Integer
    Dim VendorCol As Integer
    Dim CostCol As Integer
    Dim cell As Range

    TotalM = 0
    On Error Resume Next
    For Each sheetName In myList
        Set ws = ThisWorkbook.Worksheets(sheetName)
        If ws Is Nothing Then
            ' If the sheet does not exist, skip to the next iteration
            GoTo ContinueLoop
        End If
        
        ReqCol = 0
        VendorCol = 0
        CostCol = 0
        
        Set row2Range = ws.Range("A2:Z2")
        'find column for Req #
        For Each cell2 In row2Range
            'MsgBox cell2.Value
            If cell2.Value = "Req #" Then
                ReqCol = cell2.Column ' Assign the column number to ReqCol
            End If
            If cell2.Value = "Vendor/Cert" Then
                VendorCol = cell2.Column
            End If
            If cell2.Value = "Total Cost" Then
                CostCol = cell2.Column
            End If
            If ReqCol > 0 And VendorCol > 0 And CostCol > 0 Then
                Exit For
            End If
        Next cell2
        'MsgBox CostCol
        'ReqCol & VendorCol is set correctly at this point
        
        If ReqCol > 0 And VendorCol > 0 And CostCol > 0 Then
            ' Create column range based on ReqCol
            Set columnRange = ws.Columns(ReqCol)
            'MsgBox columnRange
            ' Find the last populated row in ReqCol
            lastRow = ws.Cells(ws.Rows.Count, ReqCol).End(xlUp).Row
            Set columnRange = ws.Range(ws.Cells(3, ReqCol), ws.Cells(lastRow, ReqCol))
            For Each cell In columnRange
                If Not IsEmpty(cell) Then
                    ' Access the cell on the same row but in the VendorCol column using Offset
                    Dim vendorCell As Range
                    Set vendorCell = cell.Offset(0, VendorCol - ReqCol)
                    'MsgBox sheetName & " " & vendorCell.Value & " " & cell.Value
                    If InStr(1, vendorCell.Value, Etype, vbTextCompare) > 0 Then 'If substring is recognized
                        Dim CostCell As Range
                        Set CostCell = cell.Offset(0, CostCol - ReqCol) 'Retrieve info from Cost cell on same line item
                        TotalM = CostCell.Value + TotalM
                    End If
                End If
            Next cell
        End If
        
ContinueLoop:
    Next sheetName
    
    BECalc = TotalM
End Function
