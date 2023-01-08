Attribute VB_Name = "Usage"
Sub Usage()


' Enter List_Of_User Sheet to collected enetered years
Sheets("Usage").Select

' Assigne variables to the date
Dim DateFrom As String
Dim DateTo As String

' Collect the entered From & To Date
YearEntered = Range("B3").Value
YearSplit = Split(YearEntered, "-")
YearFrom = YearSplit(0)
YearTo = YearSplit(1)

DateFrom = YearFrom & "-05-01"
DateTo = YearTo & "-04-30"


'** Move to the User Sheet
Sheets("Orders").Select

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

' Loop Through to collect data for the fisical year
For row = No_Of_Rows To 3 Step -1
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Assigning variables
    Set new_Client = Range("K" & row)
    Set num_Cultures = Range("L" & row)
    Set num_Strain = Range("M" & row)
    Set ml_Culture = Range("N" & row)
    Set ml_Medium = Range("O" & row)
    Set ml_Concentrate = Range("P" & row)
    
    ' Enter only if its meets the condition of the fisical year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        ' Each Usage
        If IsDate(cellDate) And Not IsEmpty(cellDate) Then
            numRequests.Add cellDate
        End If
        
        If new_Client = "yes" And Not IsEmpty(new_Client) Then
            newClientList.Add cellDate
        End If
        
        If IsNumeric(num_Cultures) And Not IsEmpty(num_Cultures) Then
            numCulList.Add cellDate
            numCulList.Add num_Cultures
        End If
        
        If IsNumeric(num_Strain) And Not IsEmpty(num_Strain) Then
            numStraList.Add cellDate
            numStraList.Add num_Strain
        End If
         
        If IsNumeric(ml_Culture) And Not IsEmpty(ml_Culture) Then
            mlCulList.Add cellDate
            mlCulList.Add ml_Culture
        End If
         
        If IsNumeric(ml_Medium) And Not IsEmpty(ml_Medium) Then
            mlMedList.Add cellDate
            mlMedList.Add ml_Medium
        End If
        
        If IsNumeric(ml_Concentrate) And Not IsEmpty(ml_Concentrate) Then
            mlConList.Add cellDate
            mlConList.Add ml_Concentrate
        End If
        
    End If
Next row

Dim requestCount As Integer
Dim newClientCount As Integer
Dim numCulturesCount As Integer
Dim numStrainCount As Integer
Dim mlCulCount As Integer
Dim mlMedCount As Integer
Dim mlConCount As Integer


Sheets("Usage").Select

'*** Request
'** Request per Month
' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

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
    Cells(6, Index) = requestCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

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
    Cells(7, Index) = newClientCount
    
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
    Cells(8, Index) = numCultureCount
    
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
        StrainDate = numStraList(i)
        StrainRequest = numStraList(i + 1)

        ' If the Month matches add
        If Month(DateNext) = Month(StrainDate) Then
            numStrainCount = numStrainCount + StrainRequest
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 1
    Cells(9, Index) = numStrainCount
    
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
    Cells(10, Index) = volCultureCount
    
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
    Cells(11, Index) = mlMediumCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n


'************************* USAGE Year CPCC
yearSelect = YearFrom & "-" & YearTo

' Find the row
yearCount = YearTo - 1999
yearFinal = 15 + yearCount
'MsgBox yearFinal

' Find the number of users
Sheets("List_Of_Users").Select
userRow = Range("A" & Rows.Count).End(xlUp).row - 1

' Find each of the value
Sheets("Usage").Select
numRequest = Range("N6")
numCultures = Range("N8")
'numUsers = userRow
numUsers = 0
numNewUsers = Range("N7")


' ** Determine Whether the year exists
i = 15
Exist = False
Do While Cells(i, 1).Value <> "Total"
    'your code here
    If Cells(i, 1).Value = yearSelect Then
        Exist = True
        yearIndex = i
    End If
    i = i + 1
Loop

'MsgBox yearIndex
'MsgBox Exist


' ** Add the values into the chart

' Case 1. The year does NOT exists
If Exist = False Then

    ' Set the variables
    yearIndex = i
    totalFinal = yearIndex + 1
    
    ' Shift a row down to add the year
    Range("A" & yearIndex & ":E" & yearIndex).Insert shift:=xlDown
    
    ' Add a new row of year
    Range("A" & yearIndex) = yearSelect
    Range("B" & yearIndex) = numRequest
    Range("C" & yearIndex) = numCultures
    Range("D" & yearIndex) = numUsers
    Range("E" & yearIndex) = numNewUsers
    
    ' Fix the sums in the TOTAL Row
    Range("B" & totalFinal) = "=SUM(B15:B" & yearIndex & ")"
    Range("C" & totalFinal) = "=SUM(C15:C" & yearIndex & ")"
    Range("D" & totalFinal) = "=SUM(D15:D" & yearIndex & ")"
    Range("E" & totalFinal) = "=SUM(E15:E" & yearIndex & ")"
    
    
' Case 2. The year does exists
Else
    Range("B" & yearIndex) = numRequest
    Range("C" & yearIndex) = numCultures
    Range("D" & yearIndex) = numUsers
    Range("E" & yearIndex) = numNewUsers

End If


' Find the most recent year for the Graph Chart
yearFinal = i - 1
UsageYear = Range("A" & yearFinal)
UsageSplit = Split(UsageYear, "-")
UsageTo = UsageSplit(1)
'MsgBox UsageTo


'****************** Create Chart ******************

'** Move to the User Sheet
Sheets("Usage").Select

'** Delete Charts in the Sheet
If Worksheets("Usage").ChartObjects.Count > 0 Then
    Worksheets("Usage").ChartObjects.Delete
End If

'** Create the Chart
Set MyRange = Sheets("Usage").Range("A15:A" & yearFinal & ",B15:C" & yearFinal)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRange
ActiveChart.ChartType = xlLine

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Usage of CPCC 1998" & " - " & UsageTo
End With

' Axis Title
With ActiveChart.Axes(xlValue)
 .HasTitle = True
 With .AxisTitle
 .Caption = "Amount"
 .Font.Name = "Arial"
 .Font.Size = 10
 '.Characters(10, 8).Font.Italic = True
 End With
End With

' Data setup
With ActiveChart
    '.DataLabels.Font.Name = "Arial"
    '.DataLabels.Font.Size = 11
    '.DataLabels.ShowPercentage = True
    '.DataLabels.ShowValue = True
    .SeriesCollection(1).Name = "=""Number of Requests"""
    .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(153, 153, 255)
    .SeriesCollection(2).Name = "=""Number of Cultures"""
    .SeriesCollection(2).Points(1).Format.Fill.ForeColor.RGB = RGB(153, 51, 102)
End With

' Size of the Line
With ActiveChart.PlotArea
    .Width = 245
    .Height = 200
    .Left = 20
    .Top = 100
End With

' Size of the Line Chart
With ActiveChart.Parent
     .Height = 320 ' resize
     .Width = 420  ' resize
     .Top = 160    ' reposition
     .Left = 300   ' reposition
End With

' Select the Source_of_Requests Sheet
Sheets("Usage").Select

End Sub
