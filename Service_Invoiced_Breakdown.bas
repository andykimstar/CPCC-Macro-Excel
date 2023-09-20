Sub Service_Invoiced_Breakdown()

'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Service Revenue Breakdown"

'Set the Columns in the 'Order'
invoicedDateColumn = "AG"
cultureCostColumn = "U"
'mediaTypeColumn = ""
mediaCostColumn = "V"
categoryTypeColumn = "W"
categoryCostColumn = "X"
shippingCostColumn = "Y"
OutstandingColumn = "AH"

' Sheet Color Design
sideColor = RGB(255, 204, 102) ' Side column color
totalColor = RGB(255, 153, 102) ' Total cell Color

' Start of the Ordered Row in the 'Service Revenue Breakdown' page
invoicedStartRow = 12
rowGap = 3

' Dates row
DateStartRow = "U15"
DateEndRow = "U16"

'****************************************************************************************************
Dim DateFrom As String
Dim DateTo As String

'** Move to the Service Revenue Breakdown Sheet
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



'** Move to the 'Order' Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
MsgBox No_Of_Rows

'*************** Beginning of the Service Revenue Breakdown Function ***************

' Collection of each requests
Dim invoicedItem As New Collection
Dim categoryItem As New Collection
Dim OutstandingList As New Collection
Dim OutstandingItem As New Collection

' Loop Through to collect data for the Service Revenue Breakdown
For row = No_Of_Rows To 3 Step -1

    ' Find the date of each data
    Set Cell = Range(invoicedDateColumn & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Find the invoicedItem
    Set cultureCost = Range(cultureCostColumn & row)
    'Set mediaType = Range(mediaTypeColumn & row)
    Set mediaCost = Range(mediaCostColumn & row)
    Set CategoryType = Range(categoryTypeColumn & row)
    Set categoryTypeCost = Range(categoryCostColumn & row)
    Set shippingCost = Range(shippingCostColumn & row)
    Set Outstanding = Range(OutstandingColumn & row)
    
    'MsgBox serviceFee
    
    ' Only collect data within the selected year
    
    ' Invoiced items not Outstanding
    If Outstanding = "no" And Not IsEmpty(Outstanding) And IsDate(cellDate) And Not IsEmpty(cellDate) And cellDate >= DateFrom And cellDate <= DateTo Then
        
        ' Culture
        If IsNumeric(cultureCost) And Not IsEmpty(cultureCost) Then
            invoicedItem.Add "Cultures"
            invoicedItem.Add cultureCost
            invoicedItem.Add cellDate
        End If
        
        ' Media
        If IsNumeric(mediaCost) And Not IsEmpty(mediaCost) Then
            'invoicedItem.Add mediaType
            invoicedItem.Add "Concentrate"
            invoicedItem.Add mediaCost
            invoicedItem.Add cellDate
        End If
        
        ' Category
        If IsNumeric(categoryTypeCost) And Not IsEmpty(categoryTypeCost) Then
            invoicedItem.Add CategoryType
            invoicedItem.Add categoryTypeCost
            invoicedItem.Add cellDate
            categoryItem.Add CategoryType
        End If
        
        ' Shipping
        If IsNumeric(shippingCost) And Not IsEmpty(shippingCost) Then
            invoicedItem.Add "Shipping"
            invoicedItem.Add shippingCost
            invoicedItem.Add cellDate
        End If
        
    End If
        
    ' Invoiced items Outstanding
    If Outstanding = "yes" And Not IsEmpty(Outstanding) And IsDate(cellDate) And Not IsEmpty(cellDate) And cellDate >= YearFrom And cellDate <= YearTo Then
    
        Set referenceNum = Range("B" & row)
        OutstandingItem.Add referenceNum
        OutstandingItem.Add cellDate
        
        ' Culture
        If IsNumeric(cultureCost) And Not IsEmpty(cultureCost) Then
            OutstandingList.Add "Cultures"
            OutstandingList.Add cultureCost
        End If
        
        ' Media
        If IsNumeric(mediaCost) And Not IsEmpty(mediaCost) Then
            OutstandingList.Add "Concentrate"
            OutstandingList.Add mediaCost
        End If
        
        ' Category
        If IsNumeric(categoryTypeCost) And Not IsEmpty(categoryTypeCost) Then
            OutstandingList.Add CategoryType
            OutstandingList.Add categoryTypeCost
            OutstandingItem.Add CategoryType
        End If
        
        ' Shipping
        If IsNumeric(shippingCost) And Not IsEmpty(shippingCost) Then
            OutstandingList.Add "Shipping"
            OutstandingList.Add shippingCost
        End If
                   
            
    End If
    
Next row




'** Back to the 'Service Revenue Breakdown' Sheet
Sheets(sheetName).Select

MsgBox invoicedStartRow
'MsgBox OutstandingList.Count

'** Enter the Invoiced Items
' Remove the previous list of Invoiced items
lastRow = Range("A" & Rows.Count).End(xlUp).row
Range("A" & invoicedStartRow & ":P" & lastRow).Delete shift:=xlUp


'** Enter the Invoiced Items
' Find the List of the Invoiced Items
Dim BreakdownList As New Collection
Dim IsItem As Boolean

BreakdownList.Add "Cultures"
BreakdownList.Add "Medium"
BreakdownList.Add "Concentrate"

For Each c_Item In categoryItem

    IsItem = True
    
    For Each b_item In BreakdownList
    
    
        If b_item = c_Item Then
            IsItem = False
            Exit For
        End If
        
    Next b_item
    
    If IsItem = True Then
        BreakdownList.Add c_Item
    End If

Next c_Item

BreakdownList.Add "Shipping"
'MsgBox BreakdownList.Count


'*******
' Enter the values in the list of Invoiced Items
For Index = 1 To BreakdownList.Count

    Counter = invoicedStartRow - 1 + Index
    Range("C" & Counter & ":P" & Counter).NumberFormat = "$#,##0.00" ' Make the cells a currency
    Range("A" & Counter) = BreakdownList(Index)
    
Next Index


'**** Begin Counting and entering request of each invoiced items per month
Dim totalSum As Integer

Dim invoicedMonth As New Collection
Dim invoicedYear As New Collection
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12
    For Index = 1 To invoicedItem.Count Step 3
    
    
         itemType = invoicedItem(Index)
         itemCost = invoicedItem(Index + 1)
         ItemInvoiced = invoicedItem(Index + 2)
         
         'Only collect data if its a matching month
         If Month(DateNext) = Month(ItemInvoiced) Then
             
             invoicedMonth.Add itemType
             invoicedMonth.Add itemCost
             
         End If
         
    Next Index
    
    ' Add the Monthly Request List into the year list
    invoicedYear.Add invoicedMonth
      
    ' Set deafult values
    Set invoicedMonth = New Collection ' Reset the Monthly Request List
    DateNext = DateAdd("m", 1, DateNext) ' Find the next month

Next n


'***** Outstanding_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To OutstandingList.Count Step 2

        ' Count if the items matches
        If breakdownItem = OutstandingList(i) Then
            totalSum = totalSum + OutstandingList(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 2) = totalSum
    
Next EachBreakdown


'***** FirstMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(1).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(1)(i) Then
            totalSum = totalSum + invoicedYear(1)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    'MsgBox col
    Cells(col, 3) = totalSum
    
Next EachBreakdown


'***** SecondMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(2).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(2)(i) Then
            totalSum = totalSum + invoicedYear(2)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 4) = totalSum
    
Next EachBreakdown


'***** ThirdMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(3).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(3)(i) Then
            totalSum = totalSum + invoicedYear(3)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 5) = totalSum
    
Next EachBreakdown


'***** FourthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(4).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(4)(i) Then
            totalSum = totalSum + invoicedYear(4)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 6) = totalSum
    
Next EachBreakdown


'***** FifthMonth_Request
' Loop through the list of Countries
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To invoicedYear(5).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(5)(i) Then
            totalSum = totalSum + invoicedYear(5)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 7) = totalSum
    
Next EachBreakdown


'***** SixthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(6).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(6)(i) Then
            totalSum = totalSum + invoicedYear(6)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 8) = totalSum
    
Next EachBreakdown

'***** SeventhMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(7).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(7)(i) Then
            totalSum = totalSum + invoicedYear(7)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 9) = totalSum
    
Next EachBreakdown


'***** EigthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(8).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(8)(i) Then
            totalSum = totalSum + invoicedYear(8)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 10) = totalSum
    
Next EachBreakdown

'***** NinethMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(9).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(9)(i) Then
            totalSum = totalSum + invoicedYear(9)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 11) = totalSum
    
Next EachBreakdown

'***** TenthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(10).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(10)(i) Then
            totalSum = totalSum + invoicedYear(10)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 12) = totalSum
    
Next EachBreakdown

'***** EleventhMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(11).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(11)(i) Then
            totalSum = totalSum + invoicedYear(11)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 13) = totalSum
    
Next EachBreakdown


'***** TwelevthMonth_Request
' Loop through the list of Countries
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To invoicedYear(12).Count Step 2

        ' Count if the items matches
        If breakdownItem = invoicedYear(12)(i) Then
            totalSum = totalSum + invoicedYear(12)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    Cells(col, 14) = totalSum
    
Next EachBreakdown


'***** Sum_of_Invoiced
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count
      
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    SumFormula = "=SUM(B" & col & ":N" & col & ")"
    Cells(col, 15) = SumFormula
    
Next EachBreakdown


'***** Percent_of_Invoiced
finalRow = Range("A" & Rows.Count).End(xlUp).row

' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count
      
    ' Locate the entry of the data
    col = EachBreakdown + invoicedStartRow - 1
    PercentFormula = "=O" & col & "/O" & finalRow + 1
    Cells(col, 16) = PercentFormula
    Range("P" & col).NumberFormat = "0.0%"
    
Next EachBreakdown


'***** Total_of_Invoiced
Range("A" & finalRow + 1) = "Total ="

' Loop through the list of items
For EachTotal = 1 To 15
      
    ' Locate the entry of the data
    rowStart = EachTotal + 1
    
    colName = Split(Cells(1, rowStart).Address(True, False), "$")
    Col_Letter = colName(0)
    sumEnd = finalRow
    
    SumTotal = "=SUM(" & Col_Letter & invoicedStartRow & ":" & Col_Letter & sumEnd & ")"
    Cells(finalRow + 1, rowStart) = SumTotal
    
Next EachTotal

'******************************
'***** Enter the Outstanding Date

' Delete the rows
outstandingStartRow = finalRow + rowGap
'lastOutstanding = Range("B" & Rows.Count).End(xlUp).row
'Range("B" & outstandingStartRow & ":C" & lastOutstanding).Delete shift:=xlUp

' Enter the Setup Value
'Range("B" & finalRow + 2) = "Total Revenue according to UW Finance:"
'Range("C" & finalRow + 2) = Range("O8")
Range("B" & outstandingStartRow) = "Outstanding"
Range("B" & outstandingStartRow + 1) = "Reference"
Range("C" & outstandingStartRow + 1) = "Date"

' Count
outstandingCount = 1

' Loop through the list of items
For EachTotal = 1 To OutstandingItem.Count Step 2
      
    ' Locate the entry of the data
    rowStart = outstandingStartRow + outstandingCount
       
    ' Enter the value
    Range("B" & rowStart) = OutstandingItem(EachTotal)
    Range("C" & rowStart) = OutstandingItem(EachTotal + 1)
    outstandingCount = outstandingCount + 1
   
Next EachTotal



'******************************
'***** Create the Summary CHART & UW Finance
Dim invoicedSummary As New Collection

' Find the variables (Start & End Rows)
totalRevenueStartRow = finalRow + rowGap
summaryStartRow = finalRow + rowGap + 2
lastSummary = Range("M" & Rows.Count).End(xlUp).row
'Range("M" & totalRevenueStartRow & ":O" & lastSummary).Delete shift:=xlUp
'finalRow = Range("A" & Rows.Count).End(xlUp).row

' Set Up Value
Range("N" & totalRevenueStartRow) = "Total Revenue according to UW Finance:"
Range("N" & totalRevenueStartRow).Font.Bold = True  ' Font Bold
Range("O" & totalRevenueStartRow) = Range("O8")
Range("O" & totalRevenueStartRow).NumberFormat = "$#,##0.00"

Range("M" & summaryStartRow) = "Summary"
Range("N" & summaryStartRow) = "$"
Range("O" & summaryStartRow) = "%"
Range("M" & summaryStartRow).Font.Bold = True  ' Font Bold
Range("N" & summaryStartRow).Font.Bold = True  ' Font Bold
Range("O" & summaryStartRow).Font.Bold = True  ' Font Bold




'****************** Create Chart ******************

'** Move to the User Sheet
Sheets(sheetName).Select

'** Delete Charts in the Sheet
If Worksheets(sheetName).ChartObjects.Count > 0 Then
    Worksheets(sheetName).ChartObjects.Delete
End If

'** Find the List of Items
fIndex = invoicedStartRow + BreakdownList.Count - 1 ' final Index
'MsgBox fIndex

'** Create the Chart
Set MyRange = Sheets(sheetName).Range("A" & invoicedStartRow & ":A" & fIndex & ",P" & invoicedStartRow & ":P" & fIndex)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRange
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Services " & YearFrom & " - " & YearTo
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = 11
    .DataLabels.ShowPercentage = True
    '.DataLabels.ShowValue = True
    
    'Color Options
    '.Points(1).Format.Fill.ForeColor.RGB = RGB(153, 153, 255)
    '.Points(2).Format.Fill.ForeColor.RGB = RGB(153, 51, 102)
    '.Points(3).Format.Fill.ForeColor.RGB = RGB(255, 255, 204)
    '.Points(4).Format.Fill.ForeColor.RGB = RGB(0, 128, 0)
    '.Points(5).Format.Fill.ForeColor.RGB = RGB(255, 128, 128)
    '.Points(6).Format.Fill.ForeColor.RGB = RGB(51, 153, 255)
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 275
    .Top = 50
    .Left = 50
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 430  ' resize
     .Top = 270 + (20 * (BreakdownList.Count - 4))  ' reposition
     .Left = 300   ' reposition
End With



' ************ Finishing Design
' Creating a Borders
lastOutstandingRow = Range("B" & Rows.Count).End(xlUp).row
lastSummaryRow = Range("M" & Rows.Count).End(xlUp).row

'If 31 > lastOutstandingRow And 31 > lastOutstandingRow Then
'    finalSheetRow = 31
'ElseIf lastOutstandingRow >= lastSummaryRow Then
'    finalSheetRow = lastOutstandingRow
'Else
'    finalSheetRow = lastSummaryRow
'End If
finalSheetRow = 31

' Invoice Chart (Phase I)
Range("A" & finalRow + 1 & ":P" & finalRow + 1).BorderAround , ColorIndex:=1
Range("A" & 11 & ":P" & finalRow + 1).BorderAround , ColorIndex:=1
Range("A" & invoicedStartRow & ":P" & invoicedStartRow).Borders(xlEdgeBottom).LineStyle = xlNone
Range("A" & invoicedStartRow & ":A" & finalRow).Interior.Color = sideColor
Range("A" & finalRow + 1).Interior.Color = totalColor
Range("A" & invoicedStartRow & ":A" & finalRow + 1).Font.Bold = True
Range("C" & finalRow + 1 & ":O" & finalRow + 1).NumberFormat = "$#,##0.00" ' Make the cells a currency
Range("A" & finalRow + 1 & ":P" & finalRow + 1).Font.Bold = True
Range("B" & invoicedStartRow & ":P" & invoicedStartRow).Font.Bold = False

' Oustanding Charts (Phase II)
'Range("B" & outstandingStartRow & ":C" & lastOutstandingRow).BorderAround , ColorIndex:=1 ' Outstanding
'Range("M" & summaryStartRow & ":O" & lastSummaryRow).BorderAround , ColorIndex:=1 ' Summary

' Border the whole chart
'Range("A" & 2 & ":P" & finalSheetRow).BorderAround , ColorIndex:=1

End Sub





