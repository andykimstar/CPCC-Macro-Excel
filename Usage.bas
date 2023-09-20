Sub Usage()


'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Usage"

' Dates row
DateStartRow = "R15"
DateEndRow = "R16"

' Set the Columns in the 'Order'
newClientColumn = "K"
culColumn = "L"
strColumn = "M"
mlCulColumn = "N"
mlStrColumn = "O"
mlConColumn = "P"

' Set the Rows in the 'Usage'
requestUsage = 6
newCliUsage = 7
culUsage = 8
strUsage = 9
volCulUsage = 10
volMedUsage = 11

'****************************************************************************************************



'************************************** Usage: Find Years *******************************************

'Dim Year As String
Dim DateFrom As String
Dim DateTo As String

'** Move to the User Sheet
Sheets(sheetName).Select

' Collect the entered From & To Date
DateFrom = Range(DateStartRow).Value
DateTo = Range(DateEndRow).Value

' Find the Months
FromMonth = CInt(Month(DateFrom))
ToMonth = CInt(Month(DateTo))

' Find the Years
FromYear = CStr(Year(DateFrom))
ToYear = CStr(Year(DateTo))

' Find the Month Name
FromString = MonthName(FromMonth, True)
ToString = MonthName(ToMonth, True)



'**************************************** Order Sheet: Data Collection ***********************************************

' Move to the User Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
Count = 0

'** Collect data for the selected Year
Dim numRequests As New Collection
Dim newClientList As New Collection
Dim numCulList As New Collection
Dim numStraList As New Collection
Dim mlCulList As New Collection
Dim mlMedList As New Collection
Dim mlConList As New Collection

' Collection of each requests
Dim order_Media As New Collection
Dim type_Media As New Collection

' Loop Through to collect data for the fisical year
For row = No_Of_Rows To 3 Step -1
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Assigning variables
    Set new_Client = Range(newClientColumn & row)
    Set num_Cultures = Range(culColumn & row)
    Set num_Strain = Range(strColumn & row)
    Set ml_Culture = Range(mlCulColumn & row)
    Set ml_Medium = Range(mlStrColumn & row)
    Set ml_Concentrate = Range(mlConColumn & row)
    
     ' Find the Media
    Set media = Range("R" & row)
    
    ' Enter only if its meets the condition of the fisical year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        ' Each Usage
        If IsDate(cellDate) And Not IsEmpty(cellDate) Then
            numRequests.Add cellDate
        End If
        
        If new_Client = "yes" And Not IsEmpty(new_Client) Then
            newClientList.Add cellDate
        End If
        
        '** # of Cultures
        If IsNumeric(num_Cultures) And Not IsEmpty(num_Cultures) Then
            numCulList.Add cellDate
            numCulList.Add num_Cultures
        End If
        
        '** # of Strain
        If IsNumeric(num_Strain) And Not IsEmpty(num_Strain) Then
            numStraList.Add cellDate
            numStraList.Add num_Strain
        End If
         
        '** mL of Cultures
        If IsNumeric(ml_Culture) And Not IsEmpty(ml_Culture) Then
            mlCulList.Add cellDate
            mlCulList.Add ml_Culture
        End If
         
        '** L of Medium
        If IsNumeric(ml_Medium) And Not IsEmpty(ml_Medium) Then
            mlMedList.Add cellDate
            mlMedList.Add ml_Medium
        End If
        
        '** mL of Concentrate
        If IsNumeric(ml_Concentrate) And Not IsEmpty(ml_Concentrate) Then
            mlConList.Add cellDate
            mlConList.Add ml_Concentrate
        End If
        
        ' Only collect data if its a matching month
        If Not IsEmpty(media) Then
            'Add the country data into the monthly list
            order_Media.Add cellDate
            order_Media.Add media
            'MsgBox media
        End If
        
         ' Only collect data if its a matching month
        If Not IsEmpty(media) Then
            'Add the country data into the monthly list
            type_Media.Add cellDate
            type_Media.Add media
            type_Media.Add mlMedList
            type_Media.Add mlConList
            'MsgBox media
        End If
        
    End If
Next row


'**************************************** Usage: Data Input ***********************************************

Dim requestCount As Integer
Dim newClientCount As Integer
Dim numCulturesCount As Integer
Dim numStrainCount As Integer
Dim mlCulCount As Integer
Dim mlMedCount As Integer
Dim mlConCount As Integer


Sheets(sheetName).Select

'*** Request
'** Request per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' *** Overall Usage
    ' CACount refreshes to zero
   requestCount = 0

    ' Count through the CA Dates
    For i = 1 To numRequests.Count
    
        'Each CA Request Date
        requestDate = numRequests(i)

        ' If the Month matches add
        If Month(DateNext) = Month(requestDate) Then
            requestCount = requestCount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(requestUsage, Index) = requestCount
    
    ' Find the Next month
    DateNext = DateAdd("m", 1, DateNext) ' Find the next month

Next n

'*** New Client
'**New Client per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CACount refreshes to zero
    newClientCount = 0

    ' Count through the CA Dates
    For i = 1 To newClientList.Count
    
        'Each CA Request Date
        NewClientDate = newClientList(i)

        ' If the Month matches add
        If Month(DateNext) = Month(NewClientDate) Then
            newClientCount = newClientCount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(newCliUsage, Index) = newClientCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n

'*** Number Cultures
'**Number of Cultures per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CACount refreshes to zero
    numCultureCount = 0

    ' Count through the CA Dates
    For i = 1 To numCulList.Count Step 2
    
        'Each CA Request Date
        CultureDate = numCulList(i)
        CultureRequest = numCulList(i + 1)

        ' If the Month matches add
        If Month(DateNext) = Month(CultureDate) Then
            numCultureCount = numCultureCount + CultureRequest
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(culUsage, Index) = numCultureCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n


'*** Number Strain
'**Number of Strain per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CACount refreshes to zero
    numStrainCount = 0

    ' Count through the CA Dates
    For i = 1 To numStraList.Count Step 2
    
        'Each CA Request Date
        strainDate = numStraList(i)
        StrainRequest = numStraList(i + 1)

        ' If the Month matches add
        If Month(DateNext) = Month(strainDate) Then
            numStrainCount = numStrainCount + StrainRequest
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(strUsage, Index) = numStrainCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n

'*** Volume Total Culture
'**Volume Total Culture per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CACount refreshes to zero
    volCultureCount = 0

    ' Count through the CA Dates
    For i = 1 To mlCulList.Count Step 2
    
        'Each CA Request Date
        volCultureDate = mlCulList(i)
        volCultureTotal = mlCulList(i + 1)

        ' If the Month matches add
        If Month(DateNext) = Month(volCultureDate) Then
            volCultureCount = volCultureCount + volCultureTotal
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(volCulUsage, Index) = volCultureCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n


'*** vol Total Medium
'**Number of Strain per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' mlMediumCount refreshes to zero
    mlMediumCount = 0

    ' Count through the CA Dates
    For i = 1 To mlMedList.Count Step 2
    
        'Each CA Request Date
        mediumDate = mlMedList(i)
        mediumRequest = mlMedList(i + 1)

        ' If the Month matches add
        If Month(DateNext) = Month(mediumDate) Then
            mlMediumCount = mlMediumCount + mediumRequest
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(volMedUsage, Index) = mlMediumCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n

End Sub

