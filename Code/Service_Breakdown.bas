Sub Service_Breakdown()


'***************************************** USER EDITS *********************************************
' Last Edit: 2025-01-02

' Sheet Name
fromsheetName = "Orders"
sheetName = "Service Revenue Breakdown"

'Set the Columns in the 'Order'
orderedDateColumn = "A"
invoicedDateColumn = "AK"
cultureCostColumn = "V"
'mediaTypeColumn = ""
mediaCostColumn = "W"
concentrateCostColumn = "X"
categoryTypeColumn = "S"
categoryCostColumn = "Y"
shippingCostColumn = "AA"

' Sheet Color Design
sideColor = RGB(255, 204, 102) ' Side column color
totalColor = RGB(255, 153, 102) ' Total cell Color

' Start of the Ordered Row in the 'Service Revenue Breakdown' page
orderedStartRow = 6
invoicedStartRow = 21
rowGap = 3

' Dates row
DateStartRow = "U18"
DateEndRow = "U19"

'****************************************************************************************************



'***************************************** Actual Code *********************************************

Dim DateFrom As String
Dim DateTo As String

'************************************** Service Breakdown: Find Years *******************************************

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



'*** Ordered Breakdown
'**************************************** Order Sheet: Data Collection ***********************************************

'** Move to the 'Order' Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
'MsgBox No_Of_Rows

'*************** Beginning of the Service Revenue Breakdown Function ***************

' Collection of each requests
Dim orderedItem As New Collection
Dim categoryItem As New Collection

' Loop Through to collect data for the Service Revenue Breakdown
For row = No_Of_Rows To 2 Step -1

    'MsgBox row

    ' Find the date of each data
    Set Cell = Range(orderedDateColumn & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Find the orderedItem
    Set cultureCost = Range(cultureCostColumn & row)        ' Culture Cost
    Set mediaCost = Range(mediaCostColumn & row)            ' Media Cost
    Set concCost = Range(concentrateCostColumn & row)       ' Concentrate Cost
    Set CategoryType = Range(categoryTypeColumn & row)      ' Category Type
    Set categoryTypeCost = Range(categoryCostColumn & row)  ' Category Cost
    Set shippingCost = Range(shippingCostColumn & row)      ' Shipping Cost
    
    ' Only collect data within the selected year
    ' Ordered items not Outstanding
    If IsDate(cellDate) And Not IsEmpty(cellDate) And cellDate >= DateFrom And cellDate <= DateTo Then
        
        ' Culture
        If IsNumeric(cultureCost) And Not IsEmpty(cultureCost) Then
            orderedItem.Add "Cultures"
            orderedItem.Add cultureCost
            orderedItem.Add cellDate
        End If
        
        ' Media
        If IsNumeric(mediaCost) And Not IsEmpty(mediaCost) Then
            'orderedItem.Add mediaType
            orderedItem.Add "Medium"
            orderedItem.Add mediaCost
            orderedItem.Add cellDate
        End If
        
        ' Concentrate
        If IsNumeric(concCost) And Not IsEmpty(concCost) Then
            'orderedItem.Add mediaType
            orderedItem.Add "Concentrate"
            orderedItem.Add concCost
            orderedItem.Add cellDate
        End If
        
        ' Category
        If IsNumeric(categoryTypeCost) And Not IsEmpty(categoryTypeCost) Then
            orderedItem.Add CategoryType
            orderedItem.Add categoryTypeCost
            orderedItem.Add cellDate
            'categoryItem.Add CategoryType
        End If
        
        ' Shipping
        If IsNumeric(shippingCost) And Not IsEmpty(shippingCost) Then
            orderedItem.Add "Shipping & Handling"
            orderedItem.Add shippingCost
            orderedItem.Add cellDate
        End If
        
    End If
    
Next row


'MsgBox orderedItem.Count

' ******************************* Service Ordered Breakdown Sheet *******************************

'** Back to the 'Service Revenue Breakdown' Sheet
Sheets(sheetName).Select

'MsgBox orderedStartRow
'MsgBox OutstandingList.Count

'** Enter the Ordered Items
' Remove the previous list of Ordered items
LastRow = Range("A" & Rows.Count).End(xlUp).row
'Range("A" & orderedStartRow & ":P" & lastRow).Delete shift:=xlUp


'** Enter the Ordered Items
' Find the List of the Ordered Items
Dim BreakdownList As New Collection

i = orderedStartRow
Do While Cells(i, 1).Value <> "Total ="
    'your code here
     serviceType = Cells(i, 1).Value
    'MsgBox mediaType
    BreakdownList.Add serviceType
    i = i + 1
Loop


'**** Begin Counting and entering request of each ordered items per month
Dim totalSum As Integer

Dim orderedMonth As New Collection
Dim orderedYear As New Collection
DateNext = DateFrom


For Index = 1 To orderedItem.Count Step 3
    
    itemType = orderedItem(Index)
    itemCost = orderedItem(Index + 1)
    ItemOrdered = orderedItem(Index + 2)
    
    'If itemType = "Concentrate" Then
    '    MsgBox itemCost
    'End If
         
Next Index
    
    
' Count up the 12 month
For n = 1 To 12
    For Index = 1 To orderedItem.Count Step 3
    
         itemType = orderedItem(Index)
         itemCost = orderedItem(Index + 1)
         ItemOrdered = orderedItem(Index + 2)
         
         'Only collect data if its a matching month
         If Month(DateNext) = Month(ItemOrdered) Then
             
             orderedMonth.Add itemType
             orderedMonth.Add itemCost
             
         End If
         
    Next Index
    
    ' Add the Monthly Request List into the year list
    orderedYear.Add orderedMonth
      
    ' Set deafult values
    Set orderedMonth = New Collection ' Reset the Monthly Request List
    DateNext = DateAdd("m", 1, DateNext) ' Find the next month

Next n


'***** FirstMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(1).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(1)(i) Then
            totalSum = totalSum + orderedYear(1)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    'MsgBox col
    Cells(col, 2) = totalSum
    
Next EachBreakdown


'***** SecondMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(2).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(2)(i) Then
            totalSum = totalSum + orderedYear(2)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 3) = totalSum
    
Next EachBreakdown


'***** ThirdMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(3).Count Step 2
    
        ' Count if the items matches
        If breakdownItem = orderedYear(3)(i) Then
            totalSum = totalSum + orderedYear(3)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 4) = totalSum
    
Next EachBreakdown


'***** FourthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(4).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(4)(i) Then
            totalSum = totalSum + orderedYear(4)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 5) = totalSum
    
Next EachBreakdown


'***** FifthMonth_Request
' Loop through the list of Countries
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of countries in each given month
    For i = 1 To orderedYear(5).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(5)(i) Then
            totalSum = totalSum + orderedYear(5)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 6) = totalSum
    
Next EachBreakdown


'***** SixthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(6).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(6)(i) Then
            totalSum = totalSum + orderedYear(6)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 7) = totalSum
    
Next EachBreakdown

'***** SeventhMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(7).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(7)(i) Then
            totalSum = totalSum + orderedYear(7)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 8) = totalSum
    
Next EachBreakdown


'***** EigthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(8).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(8)(i) Then
            totalSum = totalSum + orderedYear(8)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 9) = totalSum
    
Next EachBreakdown

'***** NinethMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(9).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(9)(i) Then
            totalSum = totalSum + orderedYear(9)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 10) = totalSum
    
Next EachBreakdown

'***** TenthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(10).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(10)(i) Then
            totalSum = totalSum + orderedYear(10)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 11) = totalSum
    
Next EachBreakdown

'***** EleventhMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(11).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(11)(i) Then
            totalSum = totalSum + orderedYear(11)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 12) = totalSum
    
Next EachBreakdown


'***** TwelevthMonth_Request
' Loop through the list of Countries
For EachBreakdown = 1 To BreakdownList.Count

    'Set default values
    breakdownItem = BreakdownList(EachBreakdown)
    totalSum = 0
    
    ' Loop through the list of items in each given month
    For i = 1 To orderedYear(12).Count Step 2

        ' Count if the items matches
        If breakdownItem = orderedYear(12)(i) Then
            totalSum = totalSum + orderedYear(12)(i + 1)
        End If
    
    Next i
    
    ' Locate the entry of the data
    col = EachBreakdown + orderedStartRow - 1
    Cells(col, 13) = totalSum
    
Next EachBreakdown



'*** Invoiced Breakdown
'**************************************** Order Sheet: Data Collection ***********************************************

'** Move to the 'Order' Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
'MsgBox No_Of_Rows

'*************** Beginning of the Service Revenue Breakdown Function ***************

' Collection of each requests
Dim invoicedItem As New Collection
'Dim categoryItem As New Collection

' Loop Through to collect data for the Service Revenue Breakdown
For row = No_Of_Rows To 2 Step -1

    'MsgBox row

    ' Find the date of each data
    Set Cell = Range(invoicedDateColumn & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Find the invoicedItem
    Set cultureCost = Range(cultureCostColumn & row)        ' Culture Cost
    Set mediaCost = Range(mediaCostColumn & row)            ' Media Cost
    Set concCost = Range(concentrateCostColumn & row)       ' Concentrate Cost
    Set CategoryType = Range(categoryTypeColumn & row)      ' Category Type
    Set categoryTypeCost = Range(categoryCostColumn & row)  ' Category Cost
    Set shippingCost = Range(shippingCostColumn & row)      ' Shipping Cost
    
    ' Only collect data within the selected year
    ' Invoiced items not Outstanding
    If IsDate(cellDate) And Not IsEmpty(cellDate) And cellDate >= DateFrom And cellDate <= DateTo Then
        
        ' Culture
        If IsNumeric(cultureCost) And Not IsEmpty(cultureCost) Then
            invoicedItem.Add "Cultures"
            invoicedItem.Add cultureCost
            invoicedItem.Add cellDate
        End If
        
        ' Media
        If IsNumeric(mediaCost) And Not IsEmpty(mediaCost) Then
            'invoicedItem.Add mediaType
            invoicedItem.Add "Medium"
            invoicedItem.Add mediaCost
            invoicedItem.Add cellDate
        End If
        
        ' Concentrate
        If IsNumeric(concCost) And Not IsEmpty(concCost) Then
            'invoicedItem.Add mediaType
            invoicedItem.Add "Concentrate"
            invoicedItem.Add concCost
            invoicedItem.Add cellDate
        End If
        
        ' Category
        If IsNumeric(categoryTypeCost) And Not IsEmpty(categoryTypeCost) Then
            invoicedItem.Add CategoryType
            invoicedItem.Add categoryTypeCost
            invoicedItem.Add cellDate
            'categoryItem.Add CategoryType
        End If
        
        ' Shipping
        If IsNumeric(shippingCost) And Not IsEmpty(shippingCost) Then
            invoicedItem.Add "Shipping & Handling"
            invoicedItem.Add shippingCost
            invoicedItem.Add cellDate
        End If
        
    End If
    
Next row

' ******************************* Service Invoiced Breakdown Sheet *******************************

'** Back to the 'Service Revenue Breakdown' Sheet
Sheets(sheetName).Select

'MsgBox invoicedStartRow
'MsgBox OutstandingList.Count

'** Enter the Invoiced Items
' Remove the previous list of Invoiced items
LastRow = Range("A" & Rows.Count).End(xlUp).row
'Range("A" & invoicedStartRow & ":P" & lastRow).Delete shift:=xlUp


'** Enter the Invoiced Items
' Find the List of the Invoiced Items
Dim BreakdownListInvoiced As New Collection

i = invoicedStartRow
Do While Cells(i, 1).Value <> "Total ="
    'your code here
     serviceType = Cells(i, 1).Value
    'MsgBox mediaType
    BreakdownListInvoiced.Add serviceType
    i = i + 1
Loop


'**** Begin Counting and entering request of each invoiced items per month
Dim totalSumInvoiced As Integer

Dim invoicedMonth As New Collection
Dim invoicedYear As New Collection
DateNext = DateFrom


For Index = 1 To invoicedItem.Count Step 3
    
    itemType = invoicedItem(Index)
    itemCost = invoicedItem(Index + 1)
    ItemInvoiced = invoicedItem(Index + 2)
    
    'If itemType = "Concentrate" Then
    '    MsgBox itemCost
    'End If
         
Next Index
    
    
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


'***** FirstMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 2) = totalSum
    
Next EachBreakdown


'***** SecondMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 3) = totalSum
    
Next EachBreakdown


'***** ThirdMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 4) = totalSum
    
Next EachBreakdown


'***** FourthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 5) = totalSum
    
Next EachBreakdown


'***** FifthMonth_Request
' Loop through the list of Countries
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 6) = totalSum
    
Next EachBreakdown


'***** SixthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 7) = totalSum
    
Next EachBreakdown

'***** SeventhMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 8) = totalSum
    
Next EachBreakdown


'***** EigthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 9) = totalSum
    
Next EachBreakdown

'***** NinethMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 10) = totalSum
    
Next EachBreakdown

'***** TenthMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 11) = totalSum
    
Next EachBreakdown

'***** EleventhMonth_Request
' Loop through the list of items
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 12) = totalSum
    
Next EachBreakdown


'***** TwelevthMonth_Request
' Loop through the list of Countries
For EachBreakdown = 1 To BreakdownListInvoiced.Count

    'Set default values
    breakdownItem = BreakdownListInvoiced(EachBreakdown)
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
    Cells(col, 13) = totalSum
    
Next EachBreakdown




'****************** Create Chart ******************

'** Move to the User Sheet
Sheets(sheetName).Select

'** Delete Charts in the Sheet
If Worksheets(sheetName).ChartObjects.Count > 0 Then
    Worksheets(sheetName).ChartObjects.Delete
End If

'** Find the List of Items
fOrdered = orderedStartRow + BreakdownList.Count - 1 ' final Index
fInvoiced = invoicedStartRow + BreakdownListInvoiced.Count - 1 ' final Index

allServiceOrdered = fOrdered - 1
allServiceInvoiced = fInvoiced - 1
totalServiceOrdered = fOrdered
totalServiceInvoiced = fInvoiced

ChartTitleFontSize = 15
TextTitleFontSize = 9
'MsgBox fIndex


'******************** Ordered - All Services
'** Create the Chart
Set MyRangeASO = Sheets(sheetName).Range("A" & orderedStartRow & ":A" & allServiceOrdered & ",N" & orderedStartRow & ":O" & allServiceOrdered)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRangeASO
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Ordered - All Services  (" & FromYear & " - " & ToYear & ")"
    .ChartTitle.Font.Size = ChartTitleFontSize
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = TextTitleFontSize
    .DataLabels.ShowPercentage = True
    .DataLabels.ShowValue = True
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 160
    .Top = 160
    .Left = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 380  ' resize
End With


' Location of the Pie Chart
ActiveChart.Parent.Left = Range("B36").Left
ActiveChart.Parent.Top = Range("B36").Top
'ActiveChart.HasLegend = True



'******************** Ordered - Special Services
'** Create the Chart
Set MyRangeSSO = Sheets(sheetName).Range("A" & orderedStartRow + 1 & ":A" & allServiceOrdered & ",N" & orderedStartRow + 1 & ":O" & allServiceOrdered)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRangeSSO
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie
ActiveChart.ChartStyle = 10

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Ordered - Special Services  (" & FromYear & " - " & ToYear & ")"
    .ChartTitle.Font.Size = ChartTitleFontSize
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = TextTitleFontSize
    .DataLabels.ShowPercentage = True
    .DataLabels.ShowValue = True
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 160
    .Top = 160
    .Left = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 380  ' resize
End With


' Location of the Pie Chart
ActiveChart.Parent.Left = Range("B53").Left
ActiveChart.Parent.Top = Range("B53").Top
'ActiveChart.HasLegend = True


'******************** Invoiced - All Services
'** Create the Chart
Set MyRangeASI = Sheets(sheetName).Range("A" & invoicedStartRow & ":A" & allServiceInvoiced & ",N" & invoicedStartRow & ":O" & allServiceInvoiced)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRangeASI
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Invoiced - All Services  (" & FromYear & " - " & ToYear & ")"
    .ChartTitle.Font.Size = ChartTitleFontSize
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = TextTitleFontSize
    .DataLabels.ShowPercentage = True
    .DataLabels.ShowValue = True
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 160
    .Top = 160
    .Left = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 380  ' resize
End With


' Location of the Pie Chart
ActiveChart.Parent.Left = Range("I36").Left
ActiveChart.Parent.Top = Range("I36").Top
'ActiveChart.HasLegend = True


'******************** Invoiced - Special Services
'** Create the Chart
Set MyRangeSSI = Sheets(sheetName).Range("A" & invoicedStartRow + 1 & ":A" & allServiceInvoiced & ",N" & invoicedStartRow + 1 & ":O" & allServiceInvoiced)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRangeSSI
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Invoiced - Special Services  (" & FromYear & " - " & ToYear & ")"
    .ChartTitle.Font.Size = ChartTitleFontSize
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = TextTitleFontSize
    .DataLabels.ShowPercentage = True
    .DataLabels.ShowValue = True
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 160
    .Top = 160
    .Left = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 380  ' resize
End With

' Location of the Pie Chart
ActiveChart.Parent.Left = Range("I53").Left
ActiveChart.Parent.Top = Range("I53").Top
'ActiveChart.HasLegend = True



'******************** Ordered - Total Revenue
'** Create the Chart
Set MyRangeASO = Sheets(sheetName).Range("A" & orderedStartRow & ":A" & totalServiceOrdered & ",N" & orderedStartRow & ":O" & totalServiceOrdered)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRangeASO
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Ordered - Total Revenue  (" & FromYear & " - " & ToYear & ")"
    .ChartTitle.Font.Size = ChartTitleFontSize
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = TextTitleFontSize
    .DataLabels.ShowPercentage = True
    .DataLabels.ShowValue = True
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 160
    .Top = 160
    .Left = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 380  ' resize
End With


' Location of the Pie Chart
ActiveChart.Parent.Left = Range("B71").Left
ActiveChart.Parent.Top = Range("B71").Top
'ActiveChart.HasLegend = True



'******************** Invoiced - Total Revenue
'** Create the Chart
Set MyRangeSSI = Sheets(sheetName).Range("A" & invoicedStartRow & ":A" & totalServiceInvoiced & ",N" & invoicedStartRow & ":O" & totalServiceInvoiced)
ActiveSheet.Shapes.AddChart.Select
ActiveChart.SetSourceData Source:=MyRangeSSI
ActiveChart.PlotBy = xlColumns
ActiveChart.ChartType = xl3DPie

' Chart Layout
With PlotArea
    ActiveChart.ApplyLayout (1)
End With

' Chart title
With ActiveChart
    .ChartTitle.Text = "Invoiced - Total Revenue  (" & FromYear & " - " & ToYear & ")"
    .ChartTitle.Font.Size = ChartTitleFontSize
End With

' Data setup
With ActiveChart.SeriesCollection(1)
    .DataLabels.Font.Name = "Arial"
    .DataLabels.Font.Size = TextTitleFontSize
    .DataLabels.ShowPercentage = True
    .DataLabels.ShowValue = True
End With

' Size of the Pie
With ActiveChart.PlotArea
    .Height = 160
    .Width = 160
    .Top = 160
    .Left = 100
End With

' Location of the Pie Chart
With ActiveChart.Parent
     .Height = 220 ' resize
     .Width = 380  ' resize
End With


' Location of the Pie Chart
ActiveChart.Parent.Left = Range("I71").Left
ActiveChart.Parent.Top = Range("I71").Top
'ActiveChart.HasLegend = True
End Sub
