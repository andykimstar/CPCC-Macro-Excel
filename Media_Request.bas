Sub Media_Request()

'***************************************** USER EDITS *********************************************

' Sheet Name
fromsheetName = "Orders"
sheetName = "Media Request2"

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


'******************************* Type of Media

' Collection of each requests
Dim year_Media As New Collection
Dim monthly_Media As New Collection
DateNext = DateFrom


' Count up the 12 month
For n = 1 To 12

    ' Count through each order data
    For i = 1 To order_Media.Count Step 2

        'Each CA Request Date
        mediaTypeDate = order_Media(i)
        mediaTypeRequest = order_Media(i + 1)
                
        ' Only collect data if its a matching month
        If Month(DateNext) = Month(mediaTypeDate) And Not IsEmpty(mediaTypeRequest) Then

            'Add the country data into the monthly list
            monthly_Media.Add mediaTypeRequest
            'MsgBox Media
        End If
    
    Next i
    
    'MsgBox monthly_Media.Count
    
    Set monthly_MediaList = New Collection ' Reset the Monthly Request List
    For ItemIndex = 1 To monthly_Media.Count Step 1
    
        Item = monthly_Media(ItemIndex)
        'MsgBox Item
        Result = Split(monthly_Media(ItemIndex), ", ")
        'MsgBox Result.Count
        For Each itemName In Result
            'MsgBox itemName
            monthly_MediaList.Add itemName
        Next
    
    Next ItemIndex
    
    'MsgBox monthly_MediaList.Count
    ' Add the Monthly Request List into the year list
    year_Media.Add monthly_MediaList
    'MsgBox monthly_Media.Count
    
    ' Set deafult values
    Set monthly_MediaList = New Collection ' Reset the Monthly Request List
    Set monthly_Media = New Collection ' Reset the Monthly Request List
    DateNext = DateAdd("m", 1, DateNext) ' Find the next month


Next n


'** Find the List of the Media

'Declare the Country Collection
Dim MediaList As New Collection

' ** Determine Whether the year exists
i = distance_Row
Do While Cells(i, 1).Value <> "Total"
    'your code here
    mediaType = Cells(i, 1).Value
    'MsgBox mediaType
    MediaList.Add mediaType
    i = i + 1
Loop


' Find the distance between the two table
distance = MediaList.Count + 3
startDistanceType = distance_Row + distance + 1
'MsgBox distance


'**** Begin Counting and entering request of each media per month
Dim Counter As Integer

'***** FirstMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
    'MsgBox media
    
     ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 2) = "=SUM(B" & startDistanceType & ":B" & col & ")"
    Else
    
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(1).Count
    
            ' Count if the media matches
            If media = year_Media(1)(i) Then
                Counter = Counter + 1
            End If
        Next i
        'MsgBox Counter
    
    End If
    
    ' Locate the entry of the data
    Cells(col, 2) = Counter
    
Next EachMedia


'***** SecondMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 3) = "=SUM(C" & startDistanceType & ":C" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(2).Count
    
            ' Count if the media matches
            If media = year_Media(2)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    'MsgBox media
    Cells(col, 3) = Counter
    
Next EachMedia


'***** ThirdMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 4) = "=SUM(D" & startDistanceType & ":D" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(3).Count
    
            ' Count if the media matches
            If media = year_Media(3)(i) Then
                Counter = Counter + 1
                
            End If
        Next i

    End If
    
    ' Locate the entry of the data
    Cells(col, 4) = Counter
    
Next EachMedia



'***** FourthMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 5) = "=SUM(E" & startDistanceType & ":E" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(4).Count
    
            ' Count if the media matches
            If media = year_Media(4)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 5) = Counter
    
Next EachMedia


'***** FifthMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 6) = "=SUM(F" & startDistanceType & ":F" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(5).Count
    
            ' Count if the media matches
            If media = year_Media(5)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 6) = Counter
    
Next EachMedia


'***** SixthMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 7) = "=SUM(G" & startDistanceType & ":G" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(6).Count
    
            ' Count if the media matches
            If media = year_Media(6)(i) Then
                Counter = Counter + 1
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 7) = Counter
    
Next EachMedia


'***** SeventhMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 8) = "=SUM(H" & startDistanceType & ":H" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(7).Count
    
            ' Count if the media matches
            If media = year_Media(7)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If

    ' Locate the entry of the data
    Cells(col, 8) = Counter
    
Next EachMedia


'***** EigthMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 9) = "=SUM(I" & startDistanceType & ":I" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(8).Count
    
            ' Count if the media matches
            If media = year_Media(8)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 9) = Counter
    
Next EachMedia


'***** NinethMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 10) = "=SUM(J" & startDistanceType & ":J" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(9).Count
    
            ' Count if the media matches
            If media = year_Media(9)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
     Cells(col, 10) = Counter
    
Next EachMedia


'***** TenthMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 11) = "=SUM(K" & startDistanceType & ":K" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(10).Count
    
            ' Count if the media matches
            If media = year_Media(10)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 11) = Counter
    
Next EachMedia


'***** EleventhMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 12) = "=SUM(L" & startDistanceType & ":L" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(11).Count
    
            ' Count if the media matches
            If media = year_Media(11)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 12) = Counter
    
Next EachMedia


'***** TwelvethMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count

    'Set default values
    media = MediaList(EachMedia)
    Counter = 0
    col = distance_Row + distance + EachMedia
    'MsgBox col
      
    ' If its the Last Row add the SUMMATION formula
    If EachMedia = MediaList.Count Then
          Cells(col + 1, 13) = "=SUM(M" & startDistanceType & ":M" & col & ")"
    Else
    
    ' If it is not than continue to add the values
        ' Loop through the list of media in each given month
        For i = 1 To year_Media(12).Count
    
            ' Count if the media matches
            If media = year_Media(12)(i) Then
                Counter = Counter + 1
                
            End If
        Next i
        
    End If
    
    ' Locate the entry of the data
    Cells(col, 13) = Counter
    
Next EachMedia




'***** TOTALMonth_Request
' Loop through the list of Media
For EachMedia = 1 To MediaList.Count + 1

    'Set default values
    col = distance_Row + distance + EachMedia
      
    ' If its the Last Row add the SUMMATION formula
    Cells(col, 14) = "=SUM(B" & col & ":M" & col & ")"
    
Next EachMedia


End Sub


