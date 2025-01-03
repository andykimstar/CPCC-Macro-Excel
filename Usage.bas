Sub Usage()


'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Usage"

' Dates row
DateStartRow = "R14"
DateEndRow = "R15"

' ## Set the Columns in the 'Order' ##
newClientColumn = "J"
strColumn = "L"
mlCulColumn = "M"
culColumn = "N"
mlMedColumn = "P"
mlConColumn = "R"
mergedColumn = "AB" 'Total Cost $CAD

' Set the Rows in the 'Usage'
requestUsage = 6
newCliUsage = 7
culUsage = 8
strUsage = 9
volCulUsage = 10
volMedUsage = 11
volConcUsage = 12

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

' Loop Through to collect data for the fisical Year
For rownum = No_Of_Rows To 3 Step -1

    Set ref = Range("A" & rownum)
    row = ref.row
    
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    'MsgBox cellDate
    
    ' Assigning variables
    Set new_Client = Range(newClientColumn & row)
    Set num_Cultures = Range(culColumn & row)
    Set num_Strain = Range(strColumn & row)
    Set ml_Culture = Range(mlCulColumn & row)
    Set ml_Medium = Range(mlMedColumn & row)
    Set ml_Concentrate = Range(mlConColumn & row)
    Set merged = Range(mergedColumn & row)
    
     ' Find the Media
    Set media = Range("R" & row)
    
    ' Enter only if its meets the condition of the fisical year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        ' Each Usage
        If IsDate(cellDate) And Not IsEmpty(cellDate) Then
            If Not IsEmpty(merged) Then
                numRequests.Add cellDate
            End If
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
        If ml_Culture <> "-" And Not IsEmpty(ml_Culture) Or IsNumeric(ml_Culture) Then
        
            arr = Split(ml_Culture, ", ")
            For Each each_item In arr
                mlCulList.Add cellDate
                mlCulList.Add each_item
            Next
        End If
         
        '** L of Medium
        If ml_Medium <> "-" And Not IsEmpty(ml_Medium) Or IsNumeric(ml_Medium) Then
        
            arr = Split(ml_Medium, ", ")
            For Each each_item In arr
                mlMedList.Add cellDate
                mlMedList.Add each_item
            Next
        End If
        
        '** mL of Concentrate
        If ml_Concentrate <> "-" And Not IsEmpty(ml_Concentrate) Or IsNumeric(ml_Concentrate) Then
        
            arr = Split(ml_Concentrate, ", ")
            For Each each_item In arr
                mlConList.Add cellDate
                mlConList.Add each_item
            Next
        End If
        
        
    End If
Next rownum


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

    ' Count through the Dates
    For i = 1 To newClientList.Count
    
        'Each Request Date
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
        volCultureTotal = mlCulList(i + 1) / 1000

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



'*** vol Total Concentrate
'**Number of Conc per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' mlMediumCount refreshes to zero
    mlConcCount = 0

    ' Count through the CA Dates
    For i = 1 To mlConList.Count Step 2
    
        'Each CA Request Date
        concDate = mlConList(i)
        concRequest = mlConList(i + 1) / 1000

        ' If the Month matches add
        If Month(DateNext) = Month(concDate) Then
            mlConcCount = mlConcCount + concRequest
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(volConcUsage, Index) = mlConcCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n

End Sub
