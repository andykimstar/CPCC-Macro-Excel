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
DateStartRow = "L13"
DateEndRow = "L14"
DateInception = Format(CDate("2022-12-31"), "yyyy-mm-dd")
SummaryOrdered = "O32"

'****************************************************************************************************


'************************************** Strains: Find Years *******************************************

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
TotalString = FromString + " '" + Right(FrsomYear, 2) + " - " + ToString + " '" + Right(ToYear, 2)
'Range("D8") = TotalString



'**************************************** Order Sheet: Data Collection ***********************************************

' Collect the data from the ORDER List
'** Move to the Order Sheet
Sheets(fromsheetName).Select

'** Count the number of rows
No_Of_Rows = Range("A" & Rows.Count).End(xlUp).row
Count = 0

'** Collect data for the selected Year
Dim StrainOrder As New Collection

' Loop Through to collect data for the fisical year
For row = 2 To No_Of_Rows Step 1
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Assigning variables
    Set num_Strain = Range(strainColumn & row)
    
    ' Enter only if its meets the condition of the fisical year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        '** mL of Concentrate
        If Not IsEmpty(num_Strain) And num_Strain <> 0 Then
            
            StrainArr = Split(num_Strain, ", ")
            For Each each_strain In StrainArr
                StrainOrder.Add cellDate
                StrainOrder.Add each_strain
            Next
        End If
        
    End If
    
Next row

'MsgBox (No_Of_Rows)
'MsgBox (StrainOrder.Count)

'***************************************** Move to the Strained Ordered *****************************************
Sheets(sheetName).Select

'** Find the List of the Strains
'Declare the Strain Collection
Dim StrainList As New Collection

' Find the Number of countries
LastRow = Range("A" & Rows.Count).End(xlUp).row

' Add the each country to the Country Collection
For n = startNumber To LastRow
    Strain = Range("A" & n)
    StrainList.Add Strain
    
    ' If it includes before 2022-12-31
    If DateFrom <= DateInception Then
        Range("D" & n) = Range("H" & n)   'Historical Dates
    Else
        Range("D" & n) = "-"  'New Dates
    End If
    
    Range("E" & n) = 0   'Reset
    Range("F" & n) = 0   'Reset
Next n

'**************** Adding the values
'** Run through the list of collected ordered item
For n = 1 To StrainOrder.Count Step 2

    ' Find Strain Date & Item
    strainDate = StrainOrder(n)
    strainItem = StrainOrder(n + 1)
    
    '** Run through the list of item
    For Index = 1 To StrainList.Count Step 1
    
        ' Each item of the Strain list
        ItemList = StrainList(Index)
    
        ' Count if the Strain Item Number matches
        If CStr(strainItem) = CStr(ItemList) Then

            Range("D" & Index + startNumber - 1) = strainDate
            Range("E" & Index + startNumber - 1) = Range("E" & Index + startNumber - 1) + 1
        End If
    
    Next Index
    
Next n


' Find the total number = new count + inception count
If DateFrom <= DateInception Then
    For n = startNumber To LastRow
        new_count = Range("E" & n)
        old_count = Range("G" & n)
        Range("F" & n) = new_count + old_count
    Next n
End If

' Border Placement
Range("E" & startNumber & ":" & "E" & LastRow) _
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
For n = startNumber To LastRow
    strainNumber = Range("E" & n)
    totalStrainNumber = Range("F" & n)

    ' If the Date is prior show Total Count
    If DateFrom <= DateInception Then
    
        If totalStrainNumber = 0 Then
        
            strainZero = strainZero + 1
        
        ElseIf totalStrainNumber >= 1 And totalStrainNumber < 10 Then
        
            strainTens = strainTens + 1
        
        ElseIf totalStrainNumber >= 10 And totalStrainNumber < 100 Then
            
            strainHundreds = strainHundreds + 1
        
        ElseIf totalStrainNumber >= 100 And totalStrainNumber < 1000 Then
        
            strainThousands = strainThousands + 1
        
        Else
            StrainTenThousands = StrainTenThousands + 1
        
        End If
    
    ' If the Date is after show New Count
    Else
    
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
        
    End If

Next n



'** Enter the Summary Numbers
Range("O32") = strainZero
Range("O33") = strainTens
Range("O34") = strainHundreds
Range("O35") = strainThousands
Range("O36") = StrainTenThousands
Range("O37") = strainZero + strainTens + strainHundreds + strainThousands + StrainTenThousands
End Sub
