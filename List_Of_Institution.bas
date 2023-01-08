Attribute VB_Name = "List_Of_Institution"
Sub List_Of_Institution()

'** Move to the User Sheet
Sheets("List_Of_Users").Select

Dim CA_Collection As New Collection
Dim CC_Collection As New Collection
Dim CG_Collection As New Collection
Dim IA_Collection As New Collection
Dim IC_Collection As New Collection
Dim IG_Collection As New Collection
Dim nameInstitution As String

lastRow = Range("A" & Rows.Count).End(xlUp).row

' Loop Through to collect data for the fisical year
For row = 2 To lastRow
    
    ' Assigning variables
    Set institution = Range("A" & row)
    Set user = Range("B" & row)
    Set region = Range("C" & row)
    Set country = Range("D" & row)
    Set affiliation = Range("E" & row)
    Set Request = Range("F" & row)
    nameInstitution = institution & ", " & region

    
    
    ' Enter only if its meets the condition CA
    If affiliation = "CA" Then
    
        ' Each CA Order
        Dim CAList As New Collection
        CAList.Add nameInstitution 'First Value
        CAList.Add user 'Second Value
        'CAList.Add region 'third Value
        CAList.Add country 'fourth Value
        CAList.Add affiliation 'fifth Value
        CAList.Add "Canadian Academic" 'sixth Value
        CAList.Add Request 'seventh Value
        CA_Collection.Add CAList
    End If
    
    ' Enter only if its meets the condition CC
    If affiliation = "CC" Then
    
        ' Each CC Order
        Dim CCList As New Collection
        CCList.Add nameInstitution 'First Value
        CCList.Add user 'Second Value
        'CAList.Add place 'third Value
        CCList.Add country 'fourth Value
        CCList.Add affiliation 'fifth Value
        CCList.Add "Canadian Commerical" 'sixth Value
        CCList.Add Request 'seventh Value
        CC_Collection.Add CCList
    End If
    
    
    ' Enter only if its meets the condition CG
    If affiliation = "CG" Then
    
        ' Each CC Order
        Dim CGList As New Collection
        CGList.Add nameInstitution 'First Value
        CGList.Add user 'Second Value
        'CAList.Add place 'third Value
        CGList.Add country 'fourth Value
        CGList.Add affiliation 'fifth Value
        CGList.Add "Canadian Government" 'sixth Value
        CGList.Add Request 'seventh Value
        CG_Collection.Add CGList
    End If
    
    ' Enter only if its meets the condition IA
    If affiliation = "IA" Then
    
        ' Each IA Order
        Dim IAList As New Collection
        IAList.Add nameInstitution 'First Value
        IAList.Add user 'Second Value
        'IAList.Add place 'third Value
        IAList.Add country 'fourth Value
        IAList.Add affiliation 'fifth Value
        IAList.Add "International Academic" 'sixth Value
        IAList.Add Request 'seventh Value
        IA_Collection.Add IAList
    End If
    
    
    ' Enter only if its meets the condition IC
    If affiliation = "IC" Then
    
        ' Each IC Order
        Dim ICList As New Collection
        ICList.Add nameInstitution 'First Value
        ICList.Add user 'Second Value
        'ICList.Add place 'third Value
        ICList.Add country 'fourth Value
        ICList.Add affiliation 'fifth Value
        ICList.Add "International Commerical" 'sixth Value
        ICList.Add Request 'seventh Value
        IC_Collection.Add ICList
    End If
    
    
    ' Enter only if its meets the condition IG
    If affiliation = "IG" Then
    
        ' Each IG Order
        Dim IGList As New Collection
        IGList.Add nameInstitution 'First Value
        IGList.Add user 'Second Value
        'IGList.Add place 'third Value
        IGList.Add country 'fourth Value
        IGList.Add affiliation 'fifth Value
        IGList.Add "International Government" 'sixth Value
        IGList.Add Request 'seventh Value
        IG_Collection.Add IGList
    End If
    
Next row


'** Enter the each Institution Collection into a Total Collection
Dim TotalCollection As New Collection
TotalCollection.Add CA_Collection
TotalCollection.Add CC_Collection
TotalCollection.Add CG_Collection
TotalCollection.Add IA_Collection
TotalCollection.Add IC_Collection
TotalCollection.Add IG_Collection


'** Find the total request sum by Institution
Dim TotalRequestSum As Integer
Dim TotalInstitutionSum As Integer

'** Move to the Fisical_Year Sheet
Sheets("Fisical_Institution").Select


'** Count the number of rows
lastRow = Range("A" & Rows.Count).End(xlUp).row


'** Clear the previous data
If lastRow <> 1 Then
    Range("A2:E" & lastRow).Clear
End If


'** Start going through list of Total Collection
For Each collectionItem In TotalCollection

    '** Making sure that the selected Collection is not Empty
    If collectionItem.Count <> 0 Then

    '** Find the starting row
    startRow = Range("A" & Rows.Count).End(xlUp).row + 2
    
    '** Enter data of each Institution Collection
    ' Loop through to enter data into the Fisical_Year Sheet
    For i = 1 To collectionItem.Count
    
        ' Declare index of items
        rownum = i + startRow ' row number
        Index = i * 6 ' item number
        
        'Items
        institution = collectionItem(1)(Index - 5)
        user = collectionItem(1)(Index - 4)
        country = collectionItem(1)(Index - 3)
        affiliation = collectionItem(1)(Index - 2)
        full_Aff = collectionItem(1)(Index - 1)
        num_request = collectionItem(1)(Index)
        
        ' Entering each value to the
        Range("A" & rownum) = institution
        Range("B" & rownum) = user
        Range("C" & rownum) = country
        Range("D" & rownum) = affiliation
        Range("E" & rownum) = num_request
    Next i
    
    ' Loop through to find the duplicates
    'MsgBox startRow
    'MsgBox rownum
    'startRow = startRow + 1
    'MsgBox Range("A" & startRow & ":A" & rownum).Count
    For iCntr = startRow + 1 To rownum
    
        'if the match index is not equals to current row number, then it is a duplicate value
        If Range("B" & iCntr).Value <> "" Then
        
            matchFoundIndex = WorksheetFunction.Match(Range("A" & iCntr).Value, Range("A" & startRow & ":A" & rownum), 0)
               'if the match index is not equals to current row number, then it is a duplicate value
            If iCntr <> matchFoundIndex Then
            
                original = Cells(matchFoundIndex, 2)
                Duplicate = Cells(iCntr, 2)
                
                    'duplicate_request = Cells(iCntr, 6)
                original_request = original & ", " & Duplicate
                Cells(matchFoundIndex, 2) = original_request
            
                
                'Delete Repetitive data
                'Range("A" & iCntr & ":E" & iCntr).Delete Shift:=xlUp
            End If
        End If
        
    Next iCntr
    
    '** Name of the Affiliation
    Range("A" & startRow) = affiliation & " = " & full_Aff
    Range("A" & startRow).Font.Bold = True
    
    '** Highlight the row
    Range("A" & startRow & ":E" & startRow).Select
    Selection.Interior.Color = vbYellow
    
    '** Enter the Sum and Count of the row
    lastRow = Range("A" & Rows.Count).End(xlUp).row
    
    'Find the Sum of each Insitiutions
    InstitutionSum = Range("A" & startRow + 1 & ":A" & lastRow).Count  'Count in number of Institution
    RequestSum = Application.WorksheetFunction.Sum(Range("E" & startRow + 1 & ":E" & lastRow))  'Sum in number of Request
    
    'Enter the Sum of each Insitiutions
    Range("A" & lastRow + 1) = "TOTAL # OF " & UCase(full_Aff) & " INSTITUTION =  " & InstitutionSum
    Range("E" & lastRow + 1) = "TOTAL # OF " & affiliation & " REQUEST =  " & RequestSum
    Range("A" & lastRow + 1).Font.Bold = True
    Range("E" & lastRow + 1).Font.Bold = True
    
    'Find the Total Numbers
    TotalInstitutionSum = TotalInstitutionSum + InstitutionSum
    TotalRequestSum = TotalRequestSum + RequestSum
    
    '** Highlight the row
    Range("A" & lastRow + 1 & ":E" & lastRow + 1).Select
    Selection.Interior.Color = vbYellow
    End If

Next collectionItem


'** Enter the Total Numbers
lastRow = Range("A" & Rows.Count).End(xlUp).row + 2
Range("A" & lastRow) = "TOTAL # OF INSTITUTION =  " & TotalInstitutionSum 'Total Numbers of Institution
Range("E" & lastRow) = "TOTAL # OF REQUEST =  " & TotalRequestSum 'Total Numbers of Requests
Range("A" & lastRow).Font.Bold = True
Range("E" & lastRow).Font.Bold = True

'** Highlight the row
Range("A" & lastRow & ":E" & lastRow).Select
Selection.Interior.Color = vbYellow


End Sub
