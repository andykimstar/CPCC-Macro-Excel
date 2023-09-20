Sub List_Of_Users()


'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Users List"

'Set the Columns in the 'Order'
institutionColumn = "D"
userColumn = "C"
cityColumn = "E"
regionColumn = "F"
countryColumn = "H"
affiColumn = "I"

' Start of the Row in the 'List of User' page
startRow = 4

' Dates row
DateStartRow = "J11"
DateEndRow = "J12"

'****************************************************************************************************


'***************************************** Actual Code *********************************************
' Assign Variables
Dim row As Integer
Dim i As Integer
Dim Years As String
Dim Count As Integer

' Enter List_Of_User Sheet to collected enetered years
Sheets(sheetName).Select

' Assigne variables to the date
Dim DateFrom As String
Dim DateTo As String

' Collect the entered From & To Date
DateFrom = Range(DateStartRow).Value
DateTo = Range(DateEndRow).Value

'** Move to the User Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
Count = 0

'** Determine data for the selected Year
' Loop Through to Determine
For row = 3 To No_Of_Rows
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    If cellDate >= DateFrom And cellDate <= DateTo Then Count = Count + 1
Next row

' Msg about the Number of Order
'MsgBox "Number of Order for the fiscal Year: " & Count

'** Collect data for the selected Year
Dim fisical_Order As New Collection
Dim place As String

' Loop Through to collect data for the fisical year
For row = No_Of_Rows To 3 Step -1
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Assigning variables
    Set institution = Range(institutionColumn & row)
    Set user = Range(userColumn & row)
    Set city = Range(cityColumn & row)
    Set region = Range(regionColumn & row)
    Set country = Range(countryColumn & row)
    Set affiliation = Range(affiColumn & row)
    place = city & ", " & region
    
    ' Enter only if its meets the condition of the fisical year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        ' Each Order
        Dim orderList As New Collection
        orderList.Add institution 'First Value
        orderList.Add user 'Second Value
        orderList.Add place 'third Value
        orderList.Add country 'fourth Value
        orderList.Add affiliation 'fifth Value
        fisical_Order.Add orderList
    End If
Next row


'** Move to the User Sheet
Sheets(sheetName).Select


'** Count the number of rows
lastRow = Range("F" & Rows.Count).End(xlUp).row


'** Clear the previous data
If lastRow <> 3 Then

    Range("A" & startRow & ":F" & lastRow).Clear

End If


'** Enter data in the User Sheet
' Loop through to enter data into the User Sheet
For i = 1 To fisical_Order.Count

    ' Declare index of items
    rownum = i + startRow - 1 ' row number
    Item = i * 5 ' item number
    
    'Items
    institution = fisical_Order(1)(Item - 4)
    user = fisical_Order(1)(Item - 3)
    region = fisical_Order(1)(Item - 2)
    country = fisical_Order(1)(Item - 1)
    affiliation = fisical_Order(1)(Item)
    num_request = 1
    
    ' Entering each value to the
    Range("A" & rownum) = institution
    Range("B" & rownum) = user
    Range("C" & rownum) = region
    Range("D" & rownum) = country
    Range("E" & rownum) = affiliation
    Range("F" & rownum) = num_request
Next i
    
    
'** Remove and Add the duplicates
' Count the number of rows
lastRow = Range("A" & Rows.Count).End(xlUp).row

For iCntr = lastRow To startRow Step -1
    
    'if the match index is not equals to current row number, then it is a duplicate value
    If Range("B" & iCntr).Value = "" Then
         Range("A" & iCntr & ":F" & iCntr).EntireRow.Delete
         matchFoundIndex = 0
    Else
        matchFoundIndex = WorksheetFunction.Match(Range("B" & iCntr).Value, Range("B1:B" & lastRow), 0)
        'if the match index is not equals to current row number, then it is a duplicate value
        If iCntr <> matchFoundIndex And Cells(iCntr, 1) = Cells(matchFoundIndex, 1) Then
        
            original = Cells(matchFoundIndex, 2)
            Duplicate = Cells(iCntr, 2)
            
            duplicate_request = Cells(iCntr, 6)
            original_request = Cells(matchFoundIndex, 6) + duplicate_request
            Cells(matchFoundIndex, 6) = original_request
        
            
            'Delete Repetitive data
            Range("A" & iCntr & ":F" & iCntr).Delete shift:=xlUp
        End If
    
    End If
    
Next


'** Count the number of Total
' Find the total number of SUM
finalRow = Range("A" & Rows.Count).End(xlUp).row
TotalRow = Range("A" & Rows.Count).End(xlUp).row + 2



' Add the number of TOTAL
Range("E" & TotalRow) = "Total ="
Range("F" & TotalRow) = Application.WorksheetFunction.Sum(Range("F2:F" & finalRow))

' Center the D to F Columns
Range("D" & startRow & ":F" & TotalRow).HorizontalAlignment = xlCenter 'Center the column
Range("F" & startRow & ":F" & finalRow).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous ' Right Border in Column
Range("A" & finalRow & ":F" & finalRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous ' Bottom Border in Column
Range("E" & TotalRow).Font.Bold = True 'Bold
Range("F" & TotalRow).Font.Bold = True 'Bold

End Sub


