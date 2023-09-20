
Sub Strains_Ordered()


'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Strains Ordered"

'Set the Columns in the 'Order'
strainColumn = "K"

' Start of the Row in the 'Strains Ordered' page
startNumber = 4

' Initialize variables
num = 0

' Dates row
DateStartRow = "I13"
DateEndRow = "I14"

'****************************************************************************************************



'***************************************** Actual Code *********************************************


'** Move to the User Sheet
Sheets(sheetName).Select

' Collect the entered From & To Date
' Assigne variables to the date
Dim DateFrom As String
Dim DateTo As String

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

' Find & Enter the Month String
TotalString = FromString + " '" + Right(FromYear, 2) + " - " + ToString + " '" + Right(ToYear, 2)
Range("D8") = TotalString


' Collect the data from the ORDER List
'** Move to the Order Sheet
Sheets(fromsheetName).Select

'** Collect data for the selected Year
Dim Strain_Log As New Collection

'Dim FirstMonth_Request As String
' Find the Date to start from
DateNext = DateFrom
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row


' Count through each order data
For row = 3 To No_Of_Rows Step 1

    ' Find the date of each data
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    
    ' Only collect data within the selected year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        ' Find the strain
         Set strain = Range(strainColumn & row)

         Result = Split(strain, ", ")
        'MsgBox Result.Count
        For Each strainName In Result
            'MsgBox strainName
            Strain_Log.Add cellDate
            Strain_Log.Add strainName
        Next
        
    End If

Next row


'** Move to the Accession Log
Sheets(sheetName).Select

'** Find the List of the Strains
'Declare the Strain Collection
Dim StrainList As New Collection

' Find the Number of countries
lastRow = Range("A" & Rows.Count).End(xlUp).row

' Add the each country to the Country Collection
For n = startNumber To lastRow
    strain = Range("A" & n)
    StrainList.Add strain
    
    Range("D" & n) = ""
    Range("E" & n) = 0
    'Range("L" & n) = strain
Next n


'**************** Adding the values
'** Run through the list of collected ordered item
For n = 1 To Strain_Log.Count Step 2

    ' Find Strain Date & Item
    strainDate = Strain_Log(n)
    strainItem = Strain_Log(n + 1)
    
    '** Run through the list of item
    For Index = 1 To StrainList.Count Step 1
    
        ' Each item of the Strain list
        ItemList = StrainList(Index)
    
        ' Count if the Strain Item Number matches
        If CStr(strainItem) = CStr(ItemList) Then

            Range("D" & Index + 8) = strainDate
            Range("E" & Index + 8) = Range("E" & Index + 8) + 1
        End If
    
    Next Index
    
Next n

' Border Placement
Range("E9:" & "E" & lastRow + 8) _
        .Borders(xlEdgeRight) _
            .LineStyle = XlLineStyle.xlContinuous


'******* Find the Total Summary of the Ordered Integers (Accession Log)
Dim strainZero As Integer
Dim strainTens As Integer
Dim strainHundreds As Integer
Dim strainThousands As Integer
Dim StrainTenThousands As Integer

strainZero = 0
strainTens = 0
strainHundreds = 0
strainThousands = 0
StrainTenThousands = 0


' Add the Number of Ordered
For n = startNumber To lastRow
    strainNumber = Range("E" & n)
    
    If strainNumber = 0 Then
    
        strainZero = strainZero + 1
    
    ElseIf strainNumber >= 1 And strainNumber < 10 Then
    
        strainTens = strainTens + 1
    
    ElseIf strainNumber >= 10 And strainNumber < 100 Then
        
        strainHundreds = strainHundreds + 1
    
    ElseIf strainNumber >= 100 And strainNumber < 1000 Then
    
        strainThousands = strainThousands + 1
    
    Else
        StrainTenThousands = StrainTenThousands + 1
    
    End If

Next n



'** Enter the Summary Numbers
Range("E" & lastRow + 3) = strainZero
Range("E" & lastRow + 4) = strainTens
Range("E" & lastRow + 5) = strainHundreds
Range("E" & lastRow + 6) = strainThousands
Range("E" & lastRow + 7) = StrainTenThousands
Range("E" & lastRow + 8) = strainZero + strainTens + strainHundreds + strainThousands + StrainTenThousands
End Sub



