Attribute VB_Name = "Service_Invoiced_Breakdown"
Sub Service_Invoiced_Breakdown()

Dim Year As String
Dim DateFrom As String
Dim DateTo As String

'** Move to the Service Revenue Breakdown Sheet
Sheets("Service_Revenue_Breakdown").Select

' Collect the entered From & To Date
Year = Range("C2").Value
YearSplit = Split(Year, "-")
YearFrom = YearSplit(0)
YearTo = YearSplit(1)

DateFrom = YearFrom & "-05-01"
DateTo = YearTo & "-04-30"
'MsgBox DateFrom
'MsgBox DateTo


'** Move to the Order Sheet
Sheets("Orders").Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row


'*************** Beginning of the Service Revenue Breakdown Function ***************

' Collection of each requests
Dim invoicedItem As New Collection
Dim categoryItem As New Collection
Dim OutstandingList As New Collection
Dim OutstandingItem As New Collection

' Loop Through to collect data for the Service Revenue Breakdown
For row = No_Of_Rows To 3 Step -1

    ' Find the date of each data
    Set Cell = Range("AH" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Find the invoicedItem
    Set cultureCost = Range("T" & row)
    Set mediaType = Range("U" & row)
    Set mediaCost = Range("V" & row)
    Set CategoryType = Range("W" & row)
    Set categoryTypeCost = Range("X" & row)
    Set shippingCost = Range("Y" & row)
    Set Outstanding = Range("AI" & row)
    
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
            invoicedItem.Add mediaType
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
            OutstandingList.Add mediaType
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


Sheets("Service_Revenue_Breakdown").Select


'** Enter the Invoiced Items
' Remove the previous list of Invoiced items
lastRow = Range("A" & Rows.Count).End(xlUp).row - 1
Range("A" & 13 & ":P" & lastRow).Delete shift:=xlUp


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


' Enter the values in the list of Invoiced Items
For Index = 1 To BreakdownList.Count

    Counter = 12 + Index
    Range("A" & Counter & ":P" & Counter).Insert shift:=xlDown  ' Insert the new row and shift down
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
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
    col = EachBreakdown + 12
    Cells(col, 14) = totalSum
    
Next EachBreakdown


'***** Sum_of_Invoiced
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count
      
    ' Locate the entry of the data
    col = EachBreakdown + 12
    SumFormula = "=SUM(B" & col & ":N" & col & ")"
    Cells(col, 15) = SumFormula
    
Next EachBreakdown


'***** Percent_of_Invoiced
finalRow = Range("A" & Rows.Count).End(xlUp).row

' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count
      
    ' Locate the entry of the data
    col = EachBreakdown + 12
    PercentFormula = "=O" & col & "/O" & finalRow
    Cells(col, 16) = PercentFormula
    
Next EachBreakdown


'***** Total_of_Invoiced
' Loop through the list of items
For EachTotal = 1 To 15
      
    ' Locate the entry of the data
    rowStart = EachTotal + 1
    
    colName = Split(Cells(1, rowStart).Address(True, False), "$")
    Col_Letter = colName(0)
    sumEnd = finalRow - 1
    
    SumTotal = "=SUM(" & Col_Letter & "13:" & Col_Letter & sumEnd & ")"
    Cells(finalRow, rowStart) = SumTotal
    
Next EachTotal


' Create the Invoiced CHART
Range("B13:O" & finalRow).NumberFormat = "$#,##0.00"  'Make the cells in money format
Range("B13:B" & finalRow).Font.Color = vbBlack  ' Make the column cells in Black
Range("P13:P" & finalRow).NumberFormat = "0.00%"  'Make the cells in percent format
Range("C13:P" & finalRow).Font.Bold = True  'Make the cells bold
Range("C12:N12").Interior.Color = RGB(221, 235, 247)  'Color the Month rows BLUE
Range("C13:P" & finalRow).Interior.Color = RGB(255, 255, 255)  'Color the Month rows WHITE
Range("A" & finalRow & ":P" & finalRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous  'Apply the line at the top of the Month
Range("A13:A" & finalRow - 1).Interior.Color = RGB(226, 239, 218)  'Color the Invoiced Item columns GREEN
Range("A13:A" & finalRow).Font.Bold = True  'Color the Invoiced Item columns BOLD
Range("P13:P" & finalRow).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous  'Right Border



'******************************
'***** Enter the Outstanding Date & UW Finance

' Delete the rows
startRow = finalRow + 2
lastOutstanding = Range("O" & Rows.Count).End(xlUp).row - 1
Range("O" & startRow & ":P" & lastOutstanding).Delete shift:=xlUp

' Enter the Setup Value
Range("O" & finalRow + 2) = "Total Revenue according to UW Finance:"
Range("P" & finalRow + 2) = Range("O8")
Range("O" & finalRow + 4) = "Outstanding"
Range("O" & finalRow + 5) = "Reference"
Range("P" & finalRow + 5) = "Date"

' Count
outstandingCount = 1

' Loop through the list of items
For EachTotal = 1 To OutstandingItem.Count Step 2
      
    ' Locate the entry of the data
    rowStart = finalRow + 5 + outstandingCount
       
    ' Enter the value
    Range("O" & rowStart) = OutstandingItem(EachTotal)
    Range("P" & rowStart) = OutstandingItem(EachTotal + 1)
    outstandingCount = outstandingCount + 1
   
Next EachTotal

' The Cell Font Setup
Range("O" & finalRow + 2).Font.Bold = True  ' Font Bold
Range("P" & finalRow + 2).NumberFormat = "$#,##0.00"  ' Font Currency
Range("P" & finalRow + 2).Font.Bold = True  ' Font Bold
Range("O" & finalRow + 4).Font.Bold = True  ' Font Bold
Range("O" & finalRow + 4).Font.Color = vbMagenta  ' Font color Pink
Range("O" & finalRow + 5).Font.Bold = True  ' Font Bold
Range("P" & finalRow + 5).Font.Bold = True  ' Font Bold
  


'******************************
'***** Create the Summary CHART
Dim invoicedSummary As New Collection

' Find the variables (Start & End Rows)
startRow = finalRow + 3
lastSummary = Range("B" & Rows.Count).End(xlUp).row - 1
Range("B" & startRow & ":D" & lastSummary).Delete shift:=xlUp


' Find and enter the Summary Value
For SummaryCount = 13 To finalRow - 2

    ' Declare the Variables
    ItemInvoiced = Range("A" & SummaryCount)
    ItemTotal = Range("O" & SummaryCount)
    ItemPercent = Range("P" & SummaryCount)
    SummaryIndex = SummaryCount + BreakdownList.Count + 3
    
    ' Enter the value
    Range("B" & SummaryIndex) = ItemInvoiced
    Range("C" & SummaryIndex) = ItemTotal
    Range("D" & SummaryIndex) = ItemPercent
    
Next SummaryCount


' Create the chart for Summary
Range("B" & SummaryIndex + 1) = "Total"  'Enter the Total
Range("C" & SummaryIndex + 1) = "=SUM(C" & startRow & ":C" & SummaryIndex & ")"  ' Enter the Summary of the SUMS ($)
Range("D" & SummaryIndex + 1) = "=SUM(D" & startRow & ":D" & SummaryIndex & ")"  ' Enter the Summary of the Percent (%)
Range("D" & startRow & ":D" & SummaryIndex + 1).NumberFormat = "0.00%"  ' Apply the percentage
Range("B" & SummaryIndex + 1 & ":D" & SummaryIndex + 1).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous  'Top Border
Range("B" & SummaryIndex + 1 & ":D" & SummaryIndex + 1).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous  'Bottom Border
Range("B" & startRow & ":B" & SummaryIndex + 1).Borders(xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous  'Left Border
Range("D" & startRow & ":D" & SummaryIndex + 1).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous  'Right Border
'Range("B" & startRow & ":D" & startRow).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous



'****************** Create Chart ******************

'** Move to the User Sheet
Sheets("Service_Revenue_Breakdown").Select

'** Delete Charts in the Sheet
If Worksheets("Service_Revenue_Breakdown").ChartObjects.Count > 0 Then
    Worksheets("Service_Revenue_Breakdown").ChartObjects.Delete
End If

'** Find the List of Items
fIndex = 12 + BreakdownList.Count - 1 ' final Index
'MsgBox fIndex

'** Create the Chart
Set MyRange = Sheets("Service_Revenue_Breakdown").Range("A13:A" & fIndex & ",O13:P" & fIndex)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRange
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    '.Legend.Delete
    '.ChartTitle.Delete
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
    .Width = 225
    .Height = 160
    .Left = 100
    .Top = 100
End With

' Size of the Pie Chart
With ActiveChart.Parent
     .Height = 200 ' resize
     .Width = 380  ' resize
     .Top = 270    ' reposition
     .Left = 360   ' reposition
End With

' Select the Service_Revenue_Breakdown Sheet
Sheets("Service_Revenue_Breakdown").Select

End Sub




