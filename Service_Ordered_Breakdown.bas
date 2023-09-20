Sub Service_Ordered_Breakdown()

'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Service Revenue Breakdown"

'Set the Columns in the 'Order'
serviceColumn = "T"
shippingColumn = "Y"


' Start of the Ordered Row in the 'Service Revenue Breakdown' page
orderedStartRow = 6

' Dates row
DateStartRow = "U15"
DateEndRow = "U16"

'****************************************************************************************************


'***************************************** Actual Code *********************************************
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




' Collect the entered From & To Date
'DateFrom = Range("T20").Value
'DateTo = Range("T21").Value

' Collect the entered From & To Date
'Year = Range("C2").Value

'YearFrom = Year(DateFrom)
'YearTo = Year(DateTo)
'YearTotal = YearFrom & "-" & YearTo

'Range("C2").Value = YearTotal


'** Move to the Order Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row


'*************** Beginning of the Service Revenue Breakdown Function ***************

' Collection of each requests
Dim service_List As New Collection
Dim shipping_List As New Collection


' Loop Through to collect data for the Service Revenue Breakdown
For row = No_Of_Rows To 3 Step -1

    ' Find the date of each data
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Find the affiliation
    Set serviceFee = Range(serviceColumn & row)
    Set shippingFee = Range(shippingColumn & row)
    'MsgBox serviceFee
    
    ' Only collect data within the selected year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        'MsgBox serviceFee
        
        ' Service
        If IsNumeric(serviceFee) And Not IsEmpty(serviceFee) Then
            service_List.Add cellDate 'First Value
            service_List.Add serviceFee 'Second Value
        End If
        
        ' Shipping
        If IsNumeric(shippingFee) And Not IsEmpty(shippingFee) Then
            shipping_List.Add cellDate 'First Value
            shipping_List.Add shippingFee 'Second Value
        End If
        
        
    End If
    
Next row


'** Back to 'Service Revenue Breakdown'
Sheets(sheetName).Select

'*** Service
'** Service Cost per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' serviceTotal refreshes to zero
    serviceTotal = 0

    ' Count through the service_List
    For i = 1 To service_List.Count Step 2

        'Items
        serviceDate = service_List(i)
        serviceCost = service_List(i + 1)
        'MsgBox serviceCost

        ' If the Month matches add
        If Month(DateNext) = Month(serviceDate) Then
            serviceTotal = serviceTotal + serviceCost
        End If
        
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(orderedStartRow, Index) = serviceTotal
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n


'*** Shipping
'** Shipping Cost per Month

' Find the Date to start from
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' serviceTotal refreshes to zero
    shippingTotal = 0

    ' Count through the service_List
    For i = 1 To shipping_List.Count Step 2

        'Items
        shippingDate = shipping_List(i)
        shippingCost = shipping_List(i + 1)
        'MsgBox shippingCost

        ' If the Month matches add
        If Month(DateNext) = Month(shippingDate) Then
            shippingTotal = shippingTotal + shippingCost
        End If
        
    
    Next i
    
    ' Locate the entry of the data
    Index = n + 2
    Cells(orderedStartRow + 1, Index) = shippingTotal
    
    ' Find the next month
    DateNext = DateAdd("m", 1, DateNext)
    'MsgBox Month(DateNext)

Next n


End Sub

