''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This script holds functions used in the procurement log spreadsheet.
' Anytime code is updated, record the date and contributor name on the last line of this comment block.
'
' Contributors:
' Moinak Das | Moinak.Das@stonybrook.edu
'
' Last update on 1/10/2024 by Moinak Das
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'FindTotalDelta looks at the line it was called on, records the cost code, and then return the sum of all the total costs
'deducted from the money in the cost code up until the line it was called upon.
Function FindTotalDelta(sheetName As String)
    
    ''VARIABLE DECLARATION ------------------------------------------------------------------------------------------------------
    Dim ws As Worksheet           'variable to store worksheet object
    Dim columnRange As Range      'variable to store column range for the for loop to iterate through
    Dim cell As Range             'variable to store current cell in for loop (pointer/cursor equiv.)
    Dim stopCell As Range         'variable to store the cell that the for loop stops at (end of for loop w/ break statement)
    Dim hValue As Variant         'variable to store the cost code (as String)
    Dim totalCount As Single      'variable to store the current calculated value for total delta, updating during for loop
    Dim callingRow As Long        'variable to store the row from which the function is called
    
    Dim CostCol As Integer        'within outer for loop, store the "Cost Code" column number for the current sheet
    Dim TotalCostCol As Integer   'within outer for loop, store the "Total Cost" column number for the current sheet
    
    ''VARIABLE INITIALIZATION ----------------------------------------------------------------------------------------------------
    totalCount = 0                              'set current pointer to zero, this value will update during for loop
    Set ws = ThisWorkbook.Worksheets(sheetName) 'set sheet as the input value, so that the script searches through that sheet
    CostCol = 0                                 'Set "Cost Code" column number as zero (to be updated in first for loop)
    TotalCostCol = 0                            'Set "Total Cost" column number as zero (to be updated in first for loop)
          
    On Error Resume Next
    
    ''FIND COST CODE COLUMN AND TOTAL COST COLUMN IN SPREADSHEET -----------------------------------------------------------------------
    Set row2Range = ws.Range("A2:Z2") 'set the range as the second row from columns A through Z.
    On Error Resume Next
    'The for loop iterates cell by cell along the row
    For Each cell2 In row2Range
        'if the current cell equals the string "Cost Code", save the column index
        If cell2.Value = "Cost Code" Then
            CostCol = cell2.Column
        End If
        'if the current cell equals the string "Total Cost", save the column index
        If cell2.Value = "Total Cost" Then
            TotalCostCol = cell2.Column
        End If
        'if every value has been assigned, exit the loop early
        If CostCol > 0 And TotalCostCol > 0 Then
            Exit For
        End If
    Next cell2
    
    Set columnRange = ws.Range(ws.Cells(3, CostCol), ws.Cells(Application.Caller.Row, CostCol)) 'set range as the cost code column until the last populated cell
    hValue = ws.Cells(Application.Caller.Row, CostCol).Value 'Set hValue as the cost code located on the current row
    
    ''MAIN FOR LOOP ---------------------------------------------------------------------------------------------------------------
    For Each cell In columnRange 'Iterate cell by cell in the cost code column
        
        If IsEmpty(cell) Then  ''''''''''''''''''''''''''''''''''''''''
            GoTo ContinueLoop  'Continue even if current cell is empty
        End If                 ''''''''''''''''''''''''''''''''''''''''
        
        If Err.Number <> 0 Then '''''''''''''''''''''''''''''''''''''
            Err.Clear           'Continue if any error
            Resume Next         '
        End If                  '''''''''''''''''''''''''''''''''''''
        
        If cell.Value = hValue Then                           '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            totalCount = totalCount + ws.Cells(cell.Row, TotalCostCol) 'if the current cell is equivalent to the cost code,
                                                              ' add the total cost of that entry to totalCost
        End If                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
ContinueLoop: 'command to continue loop
    Next cell
    
    ''RETURN VALUE
    FindTotalDelta = totalCount 'return the final calculated valuw of totalDelta
    
    
End Function

'TotalSpent takes in a cost code and then returns the total amount spent in the specified cost code across all sheets
Function TotalSpent(COST_CODE As String) As Double
    ''VARIABLE DECLARATION
    Dim myList() As Variant 'store list of all the sheet names
    Dim ws As Worksheet 'within outer for loop, store current sheet
    Dim columnRange As Range 'store range for inner for loop
    Dim cell As Range 'within inner for loop, store current cell
    
    Dim CostCol As Integer     'within outer for loop, store the "Cost Code" column number for the current sheet
    Dim TotalCostCol As Integer     'within outer for loop, store the "Total Cost" column number for the current sheet
    
    ''VARIABLE INITIALIZATION
    myList = Array("Mechanical", "Electrical", "Comms", "Track", "Traction Power", "Signals", "CMS") 'stores name of all sheets to look through
                                                                                                     'NEEDS TO BE UPDATED WHEN SHEETS ARE ADDED
    totalCount = 0 'totalCount stores the current calculated value of the
    On Error Resume Next
    
    ''OUTER FOR LOOP - iterates sheet by sheet
    For Each sheetName In myList
        
        Set ws = ThisWorkbook.Worksheets(sheetName) 'set current sheet
        
        If ws Is Nothing Then ' If the sheet does not exist, skip to the next iteration
            GoTo ContinueLoop
        End If
        
        CostCol = 0
        TotalCostCol = 0
                
        ''FIND COST CODE COLUMN AND TOTAL COST COLUMN IN SPREADSHEET
        Set row2Range = ws.Range("A2:Z2") 'set the range as the second row from columns A through Z.
        On Error Resume Next
        'The for loop iterates cell by cell along the row
        For Each cell2 In row2Range
            'if the current cell equals the string "Cost Code", save the column index
            If cell2.Value = "Cost Code" Then
                CostCol = cell2.Column
            End If
            'if the current cell equals the string "Total Cost", save the column index
            If cell2.Value = "Total Cost" Then
                TotalCostCol = cell2.Column
            End If
            'if every value has been assigned, exit the loop early
            If CostCol > 0 And TotalCostCol > 0 Then
                Exit For
            End If
        Next cell2
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        Set columnRange = ws.Range(ws.Cells(3, CostCol), ws.Cells(lastRow, CostCol)) 'set range as the cost code column until the last populated cell
        
        ''INNER FOR LOOP - iterates cell by cell and updates totalCount
        For Each cell In columnRange
            'MsgBox sheetName & " " & cell.Value & " " & totalCount 'debug purposes only
            
            If IsEmpty(cell) Then      ''''''''''''''''''''''''''''''''''
                GoTo ContinueCellLoop  'If the cell is empty then skip it
            End If                     ''''''''''''''''''''''''''''''''''
            
            If Err.Number <> 0 Then ''''''''''''''''''''''''''''''''''
                Err.Clear           'if there's an error ignore it
                Resume Next         ''''''''''''''''''''''''''''''''''
            End If
            
            If cell.Value = COST_CODE Then                        'if the specified cost code is found, then
                totalCount = totalCount + ws.Cells(cell.Row, TotalCostCol) 'add the money spent on the line item to totalCost
            End If                                                '''''''''''''''''''''''''''''''''''''''''''''''''''
            
ContinueCellLoop: 'continue inner loop, move to next cell
        Next cell
        
ContinueLoop:     'continue outer loop, move to next sheet
    Next sheetName
    
    TotalSpent = totalCount 'RETURN STATEMENT
End Function

'ProcurementPercentage takes a sheet as input, and then returns the percentage of procured items (procured items/total items)
Function ProcurementPercentage(sheetName As String)
    
    ''VARIABLE DECLARATION
    Dim ws As Worksheet            'store the worksheet object corresponding to the sheetName input
    Dim columnRange As Range       'store the column range for the for loop
    Dim lastRow                    'store the last populated row of the "items" column
    Dim ans As Double              'store the return value (decimal)
    Dim notProcured As Single      'store the number of not procured items (updated through for loop)
    Dim ReqCol As Integer      'within first for loop, store the "Req #" column number for the current sheet

    ''VARIABLE INITIALIZATION
    Set ws = ThisWorkbook.Worksheets(sheetName)           'set worksheet object to specified worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row  'set lastRow equal to last populated row of items column
    notProcured = 0                                       'set number of not procured items to zero
    ReqCol = 0                                            'set Req # column number to zero
    
    ''FIND REQ # COLUMN
    Set row2Range = ws.Range("A2:Z2") 'set the range as the second row from columns A through Z.
    
    On Error Resume Next
    'The for loop iterates cell by cell along the row
    For Each cell2 In row2Range
        
        'if the current cell equals the string "Req #", save the column index
        If cell2.Value = "Req #" Then
            ReqCol = cell2.Column
        End If
        
        'if every value has been assigned, exit the loop early
        If ReqCol > 0 Then
            Exit For
        End If
    Next cell2
    
    Set columnRange = ws.Range(ws.Cells(3, ReqCol), ws.Cells(lastRow, ReqCol))          'set columnRange as "Req #" from top of table to last row
    
    ''FOR LOOP, iterates cell by cell through "Req #" column
    For Each cell In columnRange
    
        If Err.Number <> 0 Then ''''''''''''''''''''''''''''''''''''''
            Err.Clear           'If there's an error, ignore it
            Resume Next         ''''''''''''''''''''''''''''''''''''''
        End If                  ''''''''''''''''''''''''''''''''''''''
        
        If IsEmpty(cell) Then             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            notProcured = notProcured + 1 'If cell is empty, item is not procured, add one to not procured count
        End If                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
ContinueLoop: 'continue the for loop, move to the next cell
    Next cell
    
    ''RETURN STATEMENT
    ans = (lastRow - 2 - notProcured) / (lastRow - 2) 'calculate answer, lastRow stores the row number of the last populated cell in the items column
                                                      'then subtracting two accounts for the empty row and the column title row. lastRow - 2 is equivalent
                                                      'to the total number of items. Subtracting notProcured from this value yields the number of procured
                                                      'items: (lastRow - 2 - notProcured). This is then divided by the total items (lastRow - 2) to provide
                                                      'a percentage and then assign it to the ans variable
                                                      
    ProcurementPercentage = ans
End Function

'DeliveryPercentage takes a sheet as input, and then returns the percentage of delivered items (delivered items/total items)
'Works almost exactly like the ProcurementPercentage function
Function DeliveryPercentage(sheetName As String)
   ''VARIABLE DECLARATION
    Dim ws As Worksheet            'store the worksheet object corresponding to the sheetName input
    Dim columnRange As Range       'store the column range for the for loop
    Dim lastRow                    'store the last populated row of the "items" column
    Dim ans As Double              'store the return value (decimal)
    Dim numRec As Single      'store the number of items with a Req #
    Dim numDel As Single      'store the number of items with a delivery date (delivered items)
    Dim ReqCol As Integer      'within first for loop, store the "Req #" column number for the current sheet
    Dim DelDate As Integer     'within first for loop, store the "Delivery Date" column (or similar) for the current sheet

    ''VARIABLE INITIALIZATION
    Set ws = ThisWorkbook.Worksheets(sheetName)           'set worksheet object to specified worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row  'set lastRow equal to last populated row of items column
    numRec = 0                                            'set number of requested items to zero
    numDel = 0                                            'set number of delivered items to zero
    ReqCol = 0                                            'set Req # column number to zero
    
    
    ''FIND REQ # COLUMN
    Set row2Range = ws.Range("A2:Z2") 'set the range as the second row from columns A through Z.
    On Error Resume Next
    'The for loop iterates cell by cell along the row
    For Each cell2 In row2Range
        'if the current cell equals the string "Req #", save the column index
        If cell2.Value = "Delivery Date # 1" Or cell2.Value = "Delivery Date #1" Or cell2.Value = "Delivery Date" Then
            DelDate = cell2.Column
        End If
        If cell2.Value = "Req #" Then
            ReqCol = cell2.Column
        End If
        'if every value has been assigned, exit the loop early
        If ReqCol > 0 And DelDate > 0 Then
            Exit For
        End If
    Next cell2
    
    
    Set columnRange = ws.Range(ws.Cells(3, ReqCol), ws.Cells(lastRow, ReqCol))          'set columnRange as "Req #" from top of table to last row
    
    ''FOR LOOP, iterates cell by cell through "Req #" column
    For Each cell In columnRange
        'MsgBox cell.Value & " " & notProcured
        If Err.Number <> 0 Then ''''''''''''''''''''''''''''''''''''''
            Err.Clear           'If there's an error, ignore it
            Resume Next         ''''''''''''''''''''''''''''''''''''''
        End If                  ''''''''''''''''''''''''''''''''''''''
        
        If Not (IsEmpty(cell)) Then            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            numRec = numRec + 1                'If cell is not empty, item not procured, add one to requested count
            If Not (IsEmpty(ws.Cells(cell.Row, DelDate))) Then
                numDel = numDel + 1
            End If
        End If
        
ContinueLoop: 'continue the for loop, move to the next cell
    Next cell
    
    
    ''MsgBox sheetName & " " & numRec & " " & numDel
    ''RETURN STATEMENT
    'Note that under the current structure, the if statement below may also be taken with
    'respect to procured items instead of total items (lastrow - 2)
    If lastRow < 2 Then
        DeliveryPercentage = 0
    Else
        ans = numDel / (lastRow - 2) 'calculate answer
        DeliveryPercentage = ans
    End If
End Function

'BECalc takes a disadvantaged business type as input, currently limited to {"MBE","WBE","SDVOB"}, and returns the total money spent on that business type
Function BECalc(Etype As String)

    ''VARIABLE DECLARATION
    Dim myList() As Variant    'store list of sheets in spreadsheet
    Dim ws As Worksheet        'store current worksheet in outer for loop
    Dim columnRange As Range   'store column range for inner for loop to iterate through
    Dim row2Range As Range     '
    Dim ReqCol As Integer      'within outer for loop, store the "Req #" column number for the current sheet
    Dim VendorCol As Integer   'within outer for loop, store the "Vendor/Cert" column number for the current sheet
    Dim CostCol As Integer     'within outer for loop, store the "Total Cost" column number for the current sheet
    Dim cell As Range          'store current cell within inner for loop
    
    ''VARIABLE INITIALIZATION
    myList = Array("Mechanical", "Electrical", "Comms", "Track", "Traction Power", "Signals", "CMS")
    TotalM = 0
    
    On Error Resume Next
    
    ''OUTER FOR LOOP
    For Each sheetName In myList
        Set ws = ThisWorkbook.Worksheets(sheetName) 'set current worksheet object as current sheet
        
        If ws Is Nothing Then '''''''''''''''''''''''''''''''
            GoTo ContinueLoop ' If worksheet is empty skip it
        End If                '''''''''''''''''''''''''''''''
        
        '' initialize the column indicies
        ReqCol = 0
        VendorCol = 0
        CostCol = 0
        
        '' Now, search through the second row of the spreadsheet to find the necessary column indicies
        Set row2Range = ws.Range("A2:Z2") 'set the range as the second row from columns A through Z.
                                          
        'The for loop iterates cell by cell along the row
        For Each cell2 In row2Range
            
            'if the current cell equals the string "Req #", save the column index
            If cell2.Value = "Req #" Then
                ReqCol = cell2.Column
            End If
            
            'if the current cell equals the string "Vendor/Cert", save the column index
            If cell2.Value = "Vendor/Cert" Then
                VendorCol = cell2.Column
            End If
            
            'if the current cell equals the string "Total Cost", save the column index
            If cell2.Value = "Total Cost" Then
                CostCol = cell2.Column
            End If
            
            'if every value has been assigned, exit the loop early
            If ReqCol > 0 And VendorCol > 0 And CostCol > 0 Then
                Exit For
            End If
            
        Next cell2
        'ReqCol & VendorCol is set correctly at this point
        
        
        'Before continuing, verify that every column has been identified. If necessary,
        'add an else statement to throw an error/MsgBox containing error details
        If ReqCol > 0 And VendorCol > 0 And CostCol > 0 Then
        
            Set columnRange = ws.Columns(ReqCol) 'set the column range as the entire "Req #" column
            lastRow = ws.Cells(ws.Rows.Count, ReqCol).End(xlUp).Row ' Find the last populated row in ReqCol
            Set columnRange = ws.Range(ws.Cells(3, ReqCol), ws.Cells(lastRow, ReqCol)) 'set the columnRange as the "Req #" column from row 3 to the last populated row
            
            ''FOR LOOP, iterates cell by cell through "Req #" column
            For Each cell In columnRange
                If Not IsEmpty(cell) Then
                    ' Access the cell on the same row but in the VendorCol column (vendorCell) using Offset
                    Dim vendorCell As Range
                    Set vendorCell = cell.Offset(0, VendorCol - ReqCol)
                    'MsgBox sheetName & " " & vendorCell.Value & " " & cell.Value 'debugging purposes only
                    
                    If InStr(1, vendorCell.Value, Etype, vbTextCompare) > 0 Then 'If substring is recognized
                        Dim CostCell As Range
                        Set CostCell = cell.Offset(0, CostCol - ReqCol) 'Retrieve info from corresponding "Total Cost" cell on same line item (CostCell)
                        TotalM = CostCell.Value + TotalM
                    End If
                End If
            Next cell
        End If
        
ContinueLoop:
    Next sheetName 'continue for loop
    
    ''RETURN STATEMENT
    BECalc = TotalM
End Function
