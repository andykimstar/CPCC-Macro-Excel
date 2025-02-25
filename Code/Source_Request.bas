Sub Source_Of_Request()


'***************************************** USER EDITS *********************************************
' Last Edit: 2025-01-02


' Sheet Name
fromsheetName = "Orders"
sheetName = "Source Requests"

'Set the Columns in the 'Order'
affColumn = "I"
countryColumn = "H"
mergedColumn = "AB" 'Total Cost $CAD
'Set the Rows in the 'Source Requests'
CAstart = 4
IAstart = 5
CGstart = 6
IGstart = 7
CCstart = 8
ICstart = 9
AffStart = CAstart
AffEnd = ICstart

' Start of the Country Row in the 'Source of Request' page
startCountry = 30

' Dates row
DateStartRow = "V13"
DateEndRow = "V14"

'****************************************************************************************************


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


' Move to the User Sheet
Sheets(fromsheetName).Select


'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row


'*************** Beginning of the Affiliation Request Function ***************

'** Collect data for the selected Year
Dim Affiliation_Requests As New Collection

' Collection of each requests
Dim CA_Request As New Collection
Dim CC_Request As New Collection
Dim CG_Request As New Collection
Dim IA_Request As New Collection
Dim IC_Request As New Collection
Dim IG_Request As New Collection

' Loop Through to collect data for the Affiliation Requests
For row = No_Of_Rows To 2 Step -1

    ' Find the date of each data
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Find the affiliation
    affiliation = Range(affColumn & row).Text
    Set merged = Range(mergedColumn & row)
    
    ' Only collect data within the selected year
    If merged <> "" And cellDate >= DateFrom And cellDate <= DateTo Then
    
        If affiliation = "CA" Then
            CA_Request.Add cellDate
        End If
        
        If affiliation = "CC" Then
            CC_Request.Add cellDate
        End If
        
        If affiliation = "CG" Then
            CG_Request.Add cellDate
        End If
        
        If affiliation = "IA" Then
            IA_Request.Add cellDate
        End If
        
        If affiliation = "IC" Then
            IC_Request.Add cellDate
        End If
        
        If affiliation = "IG" Then
            IG_Request.Add cellDate
        End If
        
    End If
    
Next row

'** Enter the each Institution Collection into a Total Collection
Dim TotalRequests As New Collection
TotalRequests.Add CA_Request
TotalRequests.Add CC_Request
TotalRequests.Add CG_Request
TotalRequests.Add IA_Request
TotalRequests.Add IC_Request
TotalRequests.Add IG_Request


'** Move to the Source of Requests Sheet
Sheets(sheetName).Select

Dim CACount As Integer
Dim CCount As Integer
Dim CGCount As Integer
Dim IACount As Integer
Dim ICCount As Integer
Dim IGCount As Integer

'*** CA
'** CA Request per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CACount refreshes to zero
    CACount = 0

    ' Count through the CA Dates
    For i = 1 To CA_Request.Count
    
        'Each CA Request Date
        CADate = CA_Request(i)

        ' If the Month matches add
        If Month(DateNext) = Month(CADate) Then
            CACount = CACount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(CAstart, Index) = CACount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n


'*** CC
'** CC Request per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CCCount refreshes to zero
    CCCount = 0

    ' Count through the CA Dates
    For i = 1 To CC_Request.Count
    
        'Each CC Request Date
        CCDate = CC_Request(i)

        ' If the Month matches add
        If Month(DateNext) = Month(CCDate) Then
            CCCount = CCCount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(CCstart, Index) = CCCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)

Next n


'*** CG
'** CG Request per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' CGCount refreshes to zero
    CGCount = 0

    ' Count through the CG Dates
    For i = 1 To CG_Request.Count
    
        'Each CG Request Date
        CGDate = CG_Request(i)

        ' If the Month matches add
        If Month(DateNext) = Month(CGDate) Then
            CGCount = CGCount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(CGstart, Index) = CGCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)

Next n


'*** IA
'** IA Request per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' IACount refreshes to zero
    IACount = 0

    ' Count through the IA Dates
    For i = 1 To IA_Request.Count
        
        'Each IA Request Date
        IADate = IA_Request(i)

        ' If the Month matches add
        If Month(DateNext) = Month(IADate) Then
            IACount = IACount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(IAstart, Index) = IACount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)

Next n


'*** IC
'** IC Request per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' ICCount refreshes to zero
    ICCount = 0

    ' Count through the IC Dates
    For i = 1 To IC_Request.Count
    
        'Each IC Request Date
        ICDate = IC_Request(i)

        ' If the Month matches add
        If Month(DateNext) = Month(ICDate) Then
            ICCount = ICCount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(ICstart, Index) = ICCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)

Next n


'*** IG
'** IG Request per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' IGCount refreshes to zero
    IGCount = 0

    ' Count through the IG Dates
    For i = 1 To IG_Request.Count
        
        'Each IG Request Date
        IGDate = IG_Request(i)

        ' If the Month matches add
        If Month(DateNext) = Month(IGDate) Then
            IGCount = IGCount + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(IGstart, Index) = IGCount
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)

Next n



'*************** Beginning of the Counrty Request Function ***************

'** Move to the User Sheet
Sheets(fromsheetName).Select


'** Collect data for the selected Year
Dim Country_Requests As New Collection

' Collection of each requests
Dim year_Request As New Collection
Dim monthly_Request As New Collection


'Dim FirstMonth_Request As String
' Find the Date to start from
DateNext = DateFrom
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row

' Count up the 12 month
For n = 1 To 12

    ' Count through each order data
    For row = No_Of_Rows To 3 Step -1

        ' Find the date of each data
        Set Cell = Range("A" & row)
        cellDate = Format(Cell.Value, "yyyy-mm-dd")
        
        ' Find the Country
        country = Range(countryColumn & row).Text
        Set merged = Range(mergedColumn & row)
        
        ' Only collect data within the selected year
        If merged <> "" And cellDate >= DateFrom And cellDate <= DateTo Then
        
            ' Only collect data if its a matching month
            If Month(DateNext) = Month(cellDate) Then

                'Add the country data into the monthly list
                monthly_Request.Add country
                
            End If
            
        End If
    
    Next row
    
    ' Add the Monthly Request List into the year list
    year_Request.Add monthly_Request
    
    ' Set deafult values
    Set monthly_Request = New Collection ' Reset the Monthly Request List
    DateNext = DateAdd("m", 1, DateNext) ' Find the next month

Next n



'** Move to the User Sheet
Sheets(sheetName).Select

'** Find the List of the Countries

'Declare the Country Collection
Dim CountryList As New Collection

' Find the Number of countries
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
LastRow = No_Of_Rows - 1

' Add the each country to the Country Collection
For n = startCountry To LastRow
    country = Range("A" & n)
    CountryList.Add country
Next n



'**** Begin Counting and entering request of each country per month
Dim Counter As Integer

'***** FirstMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(1).Count

        ' Count if the Country matches
        If country = year_Request(1)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 3) = Counter
    
Next EachCountry

'***** SecondMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(2).Count

        ' Count if the Country matches
        If country = year_Request(2)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 4) = Counter
    
Next EachCountry


'***** ThirdMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(3).Count

        ' Count if the Country matches
        If country = year_Request(3)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 5) = Counter
    
Next EachCountry


'***** FourthMonth_Request
' Loop through the list of Countries

For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(4).Count

        ' Count if the Country matches
        If country = year_Request(4)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 6) = Counter
    
Next EachCountry



'***** FifthMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(5).Count

        ' Count if the Country matches
        If country = year_Request(5)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 7) = Counter
    
Next EachCountry


'***** SixthMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(6).Count

        ' Count if the Country matches
        If country = year_Request(6)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 8) = Counter
    
Next EachCountry


'***** SeventhMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(7).Count

        ' Count if the Country matches
        If country = year_Request(7)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 9) = Counter
    
Next EachCountry


'***** EighthMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(8).Count

        ' Count if the Country matches
        If country = year_Request(8)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 10) = Counter
    
Next EachCountry


'***** NinthMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(9).Count

        ' Count if the Country matches
        If country = year_Request(9)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 11) = Counter
    
Next EachCountry


'***** TenthMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(10).Count

        ' Count if the Country matches
        If country = year_Request(10)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 12) = Counter
    
Next EachCountry


'***** EleventhMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(11).Count

        ' Count if the Country matches
        If country = year_Request(11)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 13) = Counter
    
Next EachCountry


'***** TwelevthMonth_Request
' Loop through the list of Countries
For EachCountry = 1 To CountryList.Count

    'Set default values
    country = CountryList(EachCountry)
    Counter = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To year_Request(12).Count

        ' Count if the Country matches
        If country = year_Request(12)(i) Then
            Counter = Counter + 1
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachCountry + startCountry - 1
    Cells(col, 14) = Counter
    
Next EachCountry



'****************** Create Chart ******************

'** Move to the User Sheet
Sheets(sheetName).Select

'** Delete Charts in the Sheet
If Worksheets(sheetName).ChartObjects.Count > 0 Then
    Worksheets(sheetName).ChartObjects.Delete
End If

'** Create the Chart
Set MyRange = Sheets(sheetName).Range("A" & AffStart & ":A" & AffEnd & ",O" & AffStart & ":P" & AffEnd)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRange
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    '.Legend.Delete
    '.ChartTitle.Delete
    .ChartTitle.Text = "Source of Requests to CPCC  (" & FromYear & " - " & ToYear & ")"
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = 11
With ActiveChart.SeriesCollection(1).DataLabels
    .ShowPercentage = True
    '.Separator = " "
    .Separator = "" & Chr(10) & ""
    .ShowValue = True
End With


    '.ApplyDataLabels Type:=xlDataLabelsShowLabel
    '.DataLabels.LegendKey = True
    '.DataLabels.ShowValue = True
    .Points(1).Format.Fill.ForeColor.RGB = RGB(153, 153, 255)
    .Points(2).Format.Fill.ForeColor.RGB = RGB(153, 51, 102)
    .Points(3).Format.Fill.ForeColor.RGB = RGB(255, 254, 204)
    .Points(4).Format.Fill.ForeColor.RGB = RGB(0, 128, 0)
    .Points(5).Format.Fill.ForeColor.RGB = RGB(255, 128, 128)
    .Points(6).Format.Fill.ForeColor.RGB = RGB(51, 153, 255)
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Width = 160
    .Height = 160
    .Left = 100
    .Top = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 450  ' resize
     .Top = 210    ' reposition
     .Left = 180   ' reposition
End With

' Select the Source of Requests Sheet
Sheets(sheetName).Select

End Sub
