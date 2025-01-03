Sub List_Of_Institution()

'***************************************** USER EDITS *********************************************
' Last Edit: 2025-01-02

' Sheet Name
fromsheetName = "Users List"
sheetName = "Institution List"

' Start of the Row in the 'List of Institution' & 'List of User' page
rowStartNumber = 4


' These columns are from the Users Lists
puserVAR = "A"
suserVAR = "B"
institutionVAR = "C"
regionVAR = "D"
countryVAR = "E"
affiliationVAR = "F"
numrequestVAR = "G"



'****************************************************************************************************


'***************************************** Actual Code *********************************************


'***************************************** Data Collection


'** Move to the User Sheet
Sheets(fromsheetName).Select

Dim CA_Collection As New Collection
Dim CC_Collection As New Collection
Dim CG_Collection As New Collection
Dim IA_Collection As New Collection
Dim IC_Collection As New Collection
Dim IG_Collection As New Collection
Dim nameInstitution As String

LastRow = Range("A" & Rows.Count).End(xlUp).row


' Loop Through to collect data from the "User List"
For row = rowStartNumber To LastRow
    
    ' Assigning variables
    Set puser = Range(puserVAR & row)
    Set suser = Range(suserVAR & row)
    Set user = puser
    Set institution = Range(institutionVAR & row)
    Set region = Range(regionVAR & row)
    Set country = Range(countryVAR & row)
    Set affiliation = Range(affiliationVAR & row)
    Set request = Range(numrequestVAR & row)
    regionCountry = region & ", " & country
    
    ' Re-Assign User
    If suser <> "-" Then
        user = puser + ", " + suser
    End If

    
    ' Enter only if its meets the condition CA
    If affiliation = "CA" Then
    
        ' Each CA Order
        Dim CAList As New Collection
        CAList.Add institution 'First Value
        CAList.Add user  'Second Value
        'CAList.Add region 'third Value
        CAList.Add regionCountry 'fourth Value
        CAList.Add affiliation 'fifth Value
        CAList.Add "Canadian Academic" 'sixth Value
        CAList.Add request 'seventh Value
        CA_Collection.Add CAList
    End If
    
    ' Enter only if its meets the condition CC
    If affiliation = "CC" Then
    
        ' Each CC Order
        Dim CCList As New Collection
        CCList.Add institution 'First Value
        CCList.Add user  'Second Value
        'CCList.Add region 'third Value
        CCList.Add regionCountry 'fourth Value
        CCList.Add affiliation 'fifth Value
        CCList.Add "Canadian Commerical" 'sixth Value
        CCList.Add request 'seventh Value
        CC_Collection.Add CCList
    End If
    
    
    ' Enter only if its meets the condition CG
    If affiliation = "CG" Then
    
        ' Each CC Order
        Dim CGList As New Collection
        CGList.Add institution 'First Value
        CGList.Add user  'Second Value
        'CAList.Add place 'third Value
        CGList.Add regionCountry 'fourth Value
        CGList.Add affiliation 'fifth Value
        CGList.Add "Canadian Government" 'sixth Value
        CGList.Add request 'seventh Value
        CG_Collection.Add CGList
    End If
    
    ' Enter only if its meets the condition IA
    If affiliation = "IA" Then
    
        ' Each IA Order
        Dim IAList As New Collection
        IAList.Add institution 'First Value
        IAList.Add user  'Second Value
        'IAList.Add place 'third Value
        IAList.Add regionCountry 'fourth Value
        IAList.Add affiliation 'fifth Value
        IAList.Add "International Academic" 'sixth Value
        IAList.Add request 'seventh Value
        IA_Collection.Add IAList
    End If
    
    
    ' Enter only if its meets the condition IC
    If affiliation = "IC" Then
    
        ' Each IC Order
        Dim ICList As New Collection
        ICList.Add institution 'First Value
        ICList.Add user 'Second Value
        'ICList.Add place 'third Value
        ICList.Add regionCountry 'fourth Value
        ICList.Add affiliation 'fifth Value
        ICList.Add "International Commerical" 'sixth Value
        ICList.Add request 'seventh Value
        IC_Collection.Add ICList
    End If
    
    
    ' Enter only if its meets the condition IG
    If affiliation = "IG" Then
    
        ' Each IG Order
        Dim IGList As New Collection
        IGList.Add institution 'First Value
        IGList.Add user  'Second Value
        'IGList.Add place 'third Value
        IGList.Add regionCountry 'fourth Value
        IGList.Add affiliation 'fifth Value
        IGList.Add "International Government" 'sixth Value
        IGList.Add request 'seventh Value
        IG_Collection.Add IGList
    End If
    
Next row




'***************************************** Data Clear

'** Move to the Fisical_Year Sheet
Sheets(sheetName).Select


'** Count the number of rows
LastRow = Range("A" & Rows.Count).End(xlUp).row


'** Clear the previous data
If LastRow <> 3 Then
    Range("A" & rowStartNumber & ":E" & LastRow).Clear
End If




'***************************************** Data Entry

'** Each Institution Collection
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


'** Start going through list of Total Collection
For Each collectionItem In TotalCollection

    '** Eliminate EMPTY Collections
    If collectionItem.Count <> 0 Then

    '** Create new starting row
    StartRow = Range("A" & Rows.Count).End(xlUp).row + 2
    
    '** Enter data of each Institution Collection
    
    ' Loop through to enter data into the 'Institution List' Sheet
    For i = 1 To collectionItem.Count
    
        ' Declare index of items
        rownum = i + StartRow ' row number
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
    
    
    '***************************************** Remove and Add the duplicates
    ' Count the number of rows
    LastRow = StartRow + 1 + collectionItem.Count
    
    For iCntr = LastRow To StartRow + 1 Step -1
        
        'if the match index is not equals to current row number, then it is a duplicate value
        If Range("A" & iCntr).Value = "" Then
             Range("A" & iCntr & ":E" & iCntr).EntireRow.Delete
             matchFoundIndex = 0
        Else
            matchFoundIndex = WorksheetFunction.Match(Range("A" & iCntr).Value, Range("A1:A" & LastRow), 0)
            
            'if the match index is not equals to current row number, then it is a duplicate value
            If iCntr <> matchFoundIndex And Cells(iCntr, 1) = Cells(matchFoundIndex, 1) And Cells(iCntr, 4) = Cells(matchFoundIndex, 4) Then
            
                original = Cells(matchFoundIndex, 1)
                Duplicate = Cells(iCntr, 1)
                
                 'MsgBox Cells(matchFoundIndex, 1)
                
                ' Duplicate User
                Duplicate_sUSer = Cells(iCntr, 2)
                Original_sUser = Cells(matchFoundIndex, 2)
                Cells(matchFoundIndex, 2) = Original_sUser + ", " + Duplicate_sUSer
                
                ' Duplicate Request
                duplicate_request = Cells(iCntr, 5)
                original_request = Cells(matchFoundIndex, 5) + duplicate_request
                Cells(matchFoundIndex, 5) = original_request
            
                'Delete Repetitive data
                Range("A" & iCntr & ":E" & iCntr).Delete shift:=xlUp
            End If
        
        End If
        
    Next

'***************************************** Data Structure

    '** Name of the Affiliation
    Range("A" & StartRow) = affiliation & " = " & full_Aff
    Range("A" & StartRow).Font.Bold = True
    
    '** Highlight the row
    Range("A" & StartRow & ":E" & StartRow).Select
    Selection.Interior.Color = vbYellow
    
    '** Enter the Sum and Count of the row
    LastRow = Range("A" & Rows.Count).End(xlUp).row
    
    'Find the Sum of each Insitiutions
    InstitutionSum = Range("A" & StartRow + 1 & ":A" & LastRow).Count  'Count in number of Institution
    RequestSum = Application.WorksheetFunction.Sum(Range("E" & StartRow + 1 & ":E" & LastRow))  'Sum in number of Request
    
    'Enter the Sum of each Insitiutions
    Range("A" & LastRow + 1) = "TOTAL # OF " & UCase(full_Aff) & " INSTITUTION =  " & InstitutionSum
    Range("E" & LastRow + 1) = "TOTAL # OF " & affiliation & " REQUEST =  " & RequestSum
    Range("A" & LastRow + 1).Font.Bold = True
    Range("E" & LastRow + 1).Font.Bold = True
    
    'Find the Total Numbers
    TotalInstitutionSum = TotalInstitutionSum + InstitutionSum
    TotalRequestSum = TotalRequestSum + RequestSum
    
    '** Highlight the row
    Range("A" & LastRow + 1 & ":E" & LastRow + 1).Select
    Selection.Interior.Color = vbYellow
    End If

Next collectionItem


'** Enter the Total Numbers
LastRow = Range("A" & Rows.Count).End(xlUp).row + 2
Range("A" & LastRow) = "TOTAL # OF INSTITUTION =  " & TotalInstitutionSum 'Total Numbers of Institution
Range("E" & LastRow) = "TOTAL # OF REQUEST =  " & TotalRequestSum 'Total Numbers of Requests
Range("A" & LastRow).Font.Bold = True
Range("E" & LastRow).Font.Bold = True

'** Highlight the row
Range("A" & LastRow & ":E" & LastRow).Select
Selection.Interior.Color = vbYellow

'** Center & Border
Range("C" & rowStartNumber & ":E" & LastRow).HorizontalAlignment = xlCenter 'Center the column
Range("A" & rowStartNumber & ":E" & LastRow).VerticalAlignment = xlCenter 'Center the column
Range("A" & rowStartNumber & ":E" & LastRow).WrapText = True
Range("E" & rowStartNumber & ":E" & LastRow).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous ' Right Border in Column
Range("A" & LastRow & ":E" & LastRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous ' Bottom Border in Column


End Sub
