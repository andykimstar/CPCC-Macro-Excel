Sub List_of_Users()


'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Users List"

'Set the Columns in the 'Order'
priUserColumn = "D"
secUserColumn = "E"
institutionColumn = "F"
regionColumn = "G"
countryColumn = "H"
affiColumn = "I"
mergedColumn = "AB" 'Total Cost $CAD

' Start of the Row in the 'List of User' page
StartRow = 4

' Dates row
DateStartRow = "K11"
DateEndRow = "K12"

'****************************************************************************************************


'***************************************** Actual Code *********************************************



'***************************************** Get Dates
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


'***************************************** Data Collection

'** Move to the User Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
Count = 0

'** Determine data for the selected Year
' Loop Through to Determine
For row = 2 To No_Of_Rows
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    If cellDate >= DateFrom And cellDate <= DateTo Then Count = Count + 1
Next row


'** Collect data for the selected Year
Dim list_of_User As New Collection

' Loop Through to collect data for the fisical year
For row = No_Of_Rows To 2 Step -1
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Assigning variables
    Set priUser = Range(priUserColumn & row)
    Set secUser = Range(secUserColumn & row)
    Set institution = Range(institutionColumn & row)
    Set region = Range(regionColumn & row)
    Set country = Range(countryColumn & row)
    Set affiliation = Range(affiColumn & row)
    Set merged = Range(mergedColumn & row)

    
    ' Non-Empty User Info
    'If InStr(institution, ";") > 0 Then
       ' Split the User Info
        'instituion_Array = Split(institution, "; ")
        'institution = institutionFull
        'region = instituion_Array(1)
    
        
        ' Enter only if its meets the condition of the fisical year
        If cellDate >= DateFrom And cellDate <= DateTo Then
        
            ' Each Order
            Dim userInfoList As New Collection
            
            ' Check End-User
            'If endUser <> "" Then
            '    userInfoList.Add endUser 'End-User
            'Else
            '    userInfoList.Add infoUser 'User
            'End If
            userInfoList.Add priUser 'Primary User
            userInfoList.Add secUser 'Secondary User
            userInfoList.Add institution 'Institution
            userInfoList.Add region 'Region
            userInfoList.Add country 'Country
            userInfoList.Add affiliation ' Affiliation
            
            If merged <> "" Then
                userInfoList.Add "new" ' Affiliation
            Else
                userInfoList.Add "merge" ' Affiliation
            End If
            
            list_of_User.Add userInfoList
            
            'MsgBox userInfoList.Count
         End If
         
    'End If
Next row


'MsgBox list_of_User.Count


'***************************************** Data Entry

'** Move to the User Sheet
Sheets(sheetName).Select

'** Count the number of rows
LastRow = Range("G" & Rows.Count).End(xlUp).row

'** Clear the previous data
If LastRow <> 3 Then

    Range("A" & StartRow & ":G" & LastRow).Clear

End If


'** Enter data in the User Sheet
' Loop through to enter data into the User Sheet
For i = 1 To list_of_User.Count

    ' Declare index of items
    rownum = i + StartRow - 1 ' row number
    Item = i * 7 ' item number
    
    'Items
    puser = list_of_User(1)(Item - 6)
    suser = list_of_User(1)(Item - 5)
    institution = list_of_User(1)(Item - 4)
    region = list_of_User(1)(Item - 3)
    country = list_of_User(1)(Item - 2)
    affiliation = list_of_User(1)(Item - 1)
    Merge = list_of_User(1)(Item)
    
    If Merge = "new" Then
         num_request = 1
    Else
         num_request = 0
    End If
    
    ' Entering each value to the
    Range("A" & rownum) = puser
    Range("B" & rownum) = suser
    Range("C" & rownum) = institution
    Range("D" & rownum) = region
    Range("E" & rownum) = country
    Range("F" & rownum) = affiliation
    Range("G" & rownum) = num_request
Next i

    
'***************************************** Remove and Add the duplicates
' Count the number of rows
LastRow = Range("A" & Rows.Count).End(xlUp).row

For iCntr = LastRow To StartRow Step -1
    
    'if the match index is not equals to current row number, then it is a duplicate value
    If Range("A" & iCntr).Value = "" Then
         Range("A" & iCntr & ":G" & iCntr).EntireRow.Delete
         matchFoundIndex = 0
    Else
        matchFoundIndex = WorksheetFunction.Match(Range("A" & iCntr).Value, Range("A1:A" & LastRow), 0)
        'if the match index is not equals to current row number, then it is a duplicate value
        If iCntr <> matchFoundIndex And Cells(iCntr, 1) = Cells(matchFoundIndex, 1) Then
        
            original = Cells(matchFoundIndex, 1)
            Duplicate = Cells(iCntr, 1)
            
            ' Duplicate User
            Duplicate_sUSer = Cells(iCntr, 2)
            Original_sUser = Cells(matchFoundIndex, 2)
            
            ' Condition in adding Duplicate User one by one
            If Duplicate_sUSer <> "-" Then
                For Each itm In Split(Duplicate_sUSer, ", ")
                    If InStr(Original_sUser, itm) <= 0 Then
                    
                        If Original_sUser <> "-" And Duplicate_sUSer <> "-" Then
                             Cells(matchFoundIndex, 2) = Original_sUser + ", " + itm
                        End If
                        
                        If Original_sUser = "-" Then
                            Cells(matchFoundIndex, 2) = itm
                        End If
                        
                    End If
                Next itm
            End If
            ' Duplicate Request
            duplicate_request = Cells(iCntr, 7)
            original_request = Cells(matchFoundIndex, 7) + duplicate_request
            Cells(matchFoundIndex, 7) = original_request
        
            
            'Delete Repetitive data
            Range("A" & iCntr & ":G" & iCntr).Delete shift:=xlUp
        End If
    
    End If
    
Next


'***************************************** Count the number of Secondary User
Dim s As String
CountSecondary = 0

' Count the number of rows
LastRow = Range("A" & Rows.Count).End(xlUp).row

For iCntr = LastRow To StartRow Step -1
    s = Range("B" & iCntr).Value
    If s <> "-" Then
        For Each itm In Split(s, ", ")    'Split cells 'iterate through the array
          CountSecondary = CountSecondary + 1
        Next itm
    End If
Next


' Re-Order all the Clients in Column B
For iCntr = LastRow To StartRow Step -1

    Dim bry() As Variant
    ch = Range("B" & iCntr).Value
    ary = Split(ch, ", ")
    L = LBound(ary)
    U = UBound(ary)
    
    ' Only when there are more than 2 Additional User
    If U - L > 0 Then
        ReDim bry(L To U)
        
        For i = LBound(ary) To UBound(ary)
            bry(i) = ary(i)
        Next i
        
        Call SortFunction(bry)
        
        For i = LBound(bry) To UBound(bry)
            ary(i) = CStr(bry(i))
        Next i
        CellSort = Join(ary, ", ")
        Cells(iCntr, 2) = CellSort
    End If

Next iCntr



'***************************************** Count the number of Total
' Find the total number of SUM
finalRow = Range("A" & Rows.Count).End(xlUp).row
TotalRow = finalRow + 2

' Add the number of TOTAL
Range("N31") = Range("G" & StartRow & ":G" & finalRow).Count
Range("N32") = CountSecondary
Range("N33") = Range("G" & StartRow & ":G" & finalRow).Count + CountSecondary
Range("N34") = Application.WorksheetFunction.Sum(Range("G" & StartRow & ":G" & finalRow))

' Center the D to F Columns
Range("B" & StartRow & ":G" & TotalRow).HorizontalAlignment = xlCenter 'Center the column
Range("A" & StartRow & ":G" & TotalRow).VerticalAlignment = xlCenter 'Center the column
Range("B" & StartRow & ":C" & TotalRow).WrapText = True
Range("G" & StartRow & ":G" & finalRow).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous ' Right Border in Column
Range("A" & finalRow & ":G" & finalRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous ' Bottom Border in Column
Range("F" & TotalRow).Font.Bold = True 'Bold
Range("G" & TotalRow).Font.Bold = True 'Bold

End Sub

Sub SortFunction(arr)
    Dim strTemp As Variant
    Dim i As Long
    Dim j As Long
    Dim lngMin As Long
    Dim lngMax As Long
    lngMin = LBound(arr)
    lngMax = UBound(arr)
    For i = lngMin To lngMax - 1
        For j = i + 1 To lngMax
            If arr(i) > arr(j) Then
                strTemp = arr(i)
                arr(i) = arr(j)
                arr(j) = strTemp
            End If
        Next j
    Next i
End Sub
