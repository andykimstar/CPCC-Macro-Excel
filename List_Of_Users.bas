Attribute VB_Name = "List_Of_Users"
Sub List_Of_Users()

' Assign Variables
Dim row As Integer
Dim i As Integer
Dim Years As String
Dim Count As Integer

' Enter List_Of_User Sheet to collected enetered years
Sheets("List_Of_Users").Select

' Assigne variables to the date
Dim DateFrom As String
Dim DateTo As String

' Collect the entered From & To Date
DateFrom = Range("I13").Value
DateTo = Range("I14").Value


'** Move to the User Sheet
Sheets("Orders").Select

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
    Set institution = Range("E" & row)
    Set user = Range("D" & row)
    Set city = Range("F" & row)
    Set region = Range("G" & row)
    Set country = Range("I" & row)
    Set affiliation = Range("J" & row)
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
Sheets("List_Of_Users").Select


'** Count the number of rows
lastRow = Range("A" & Rows.Count).End(xlUp).row


'** Clear the previous data
If lastRow <> 1 Then

    Range("A2:F" & lastRow).Clear

End If


'** Enter data in the User Sheet
' Loop through to enter data into the User Sheet
For i = 1 To fisical_Order.Count

    ' Declare index of items
    rownum = i + 1 ' row number
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

For iCntr = lastRow To 2 Step -1
    
    'if the match index is not equals to current row number, then it is a duplicate value
    If Range("B" & iCntr).Value = "" Then
         Rows(iCntr).EntireRow.Delete
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

' Center the D to F Columns
Range("D2" & ":F" & finalRow).HorizontalAlignment = xlCenter 'Center the column

' Add the number of TOTAL
Range("E" & TotalRow) = "Total ="
Range("F" & TotalRow) = Application.WorksheetFunction.Sum(Range("F2:F" & finalRow))


End Sub
