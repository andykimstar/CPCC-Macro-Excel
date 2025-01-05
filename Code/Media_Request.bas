Sub Media_Requests()


'***************************************** USER EDITS *********************************************
' Last Edit: 2025-01-02

' Sheet Name
fromsheetName = "Orders"
sheetName = "Media Requests"

' Dates rows
DateStartRow = "R15"
DateEndRow = "R16"

' Set the Columns in the 'Order'
'mediaNumColumn = "J"  'New
mediaColumn = "O"
LmediaColumn = "P"
conceColumn = "Q"  'New
mlConColumn = "R"

' Start Row
StartRow = 7
countColumn = "B"
'****************************************************************************************************


'************************************** Media Request: Find Years *******************************************

'Dim Year As String
Dim DateFrom As String
Dim DateTo As String

'** Move to the Media Request Sheet
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
No_Of_Rows = Range(countColumn & Rows.Count).End(xlUp).row
'MsgBox No_Of_Rows
Count = 0

'** Collect data for the selected Year
Dim numMediaList As New Collection
Dim numConcList As New Collection
Dim LMediaList As New Collection
Dim mlConcList As New Collection

' Loop Through to collect data for the fisical year
For row = No_Of_Rows To 2 Step -1
    Set Cell = Range("A" & row)
    cellDate = Format(Cell.Value, "yyyy-mm-dd")
    
    ' Assigning variables
    Set num_Media = Range(mediaColumn & row)
    Set l_Media = Range(LmediaColumn & row)
    Set num_Concentrate = Range(conceColumn & row)
    Set ml_Concentrate = Range(mlConColumn & row)
    
    ' Find the Media
    'Set Media = Range("M" & row)
    
    ' Enter only if its meets the condition of the fisical year
    If cellDate >= DateFrom And cellDate <= DateTo Then
    
        '** # of Media
        If Not IsNumeric(num_Media) And Not IsEmpty(num_Media) And num_Media <> "-" Then
        
            num_MediaArr = Split(num_Media, ", ")

            For Each each_media In num_MediaArr
                numMediaList.Add cellDate
                numMediaList.Add each_media
                'MsgBox (cellDate)
                'MsgBox (each_media)
            Next
        End If
        
        '** # of Concentrate
        If Not IsNumeric(num_Concentrate) And Not IsEmpty(num_Concentrate) And num_Concentrate <> "-" Then
        
            num_ConcArr = Split(num_Concentrate, ", ")
            For Each each_conc In num_ConcArr
                numConcList.Add cellDate
                numConcList.Add each_conc
            Next
        End If
         
        '** L of Media
        If Not IsEmpty(l_Media) And l_Media <> 0 And l_Media <> "-" Then
            
            L_MediaArr = Split(l_Media, ", ")
            For Each each_media_l In L_MediaArr
                LMediaList.Add cellDate
                LMediaList.Add each_media_l
                'MsgBox (cellDate)
                'MsgBox (each_media_l)
            Next
        End If
         
        '** mL of Concentrate
        If Not IsEmpty(ml_Concentrate) And ml_Concentrate <> 0 And ml_Concentrate <> "-" Then
            
            ml_ConcArr = Split(ml_Concentrate, ", ")
            For Each each_conc_ml In ml_ConcArr
                mlConcList.Add cellDate
                mlConcList.Add each_conc_ml
            Next
        End If
        
    End If
    
Next row


' ERROR Notification: Numbers dont add up between # of Media & L of Media
If numMediaList.Count <> LMediaList.Count Then

    MsgBox ("Number of Media (" & numMediaList.Count & ") & L of medium (" & LMediaList.Count & ") do not add up properly")
            
End If

' ERROR Notification: Numbers dont add up between # of Concentrate & mL of Concentrate
If numConcList.Count <> mlConcList.Count Then

    MsgBox ("Number of Concentrate (" & numConcList.Count & ") & mL of Concentrate (" & mlConcList.Count & ") do not add up properly")
            
End If

'***********
' Collection of each requests
Dim type_Media As New Collection
Dim order_Media As New Collection
Dim order_Media_Litre As Double

' Create a Collection of Media
For i = 2 To numMediaList.Count Step 2
    order_Media.Add numMediaList(i - 1)
    order_Media.Add numMediaList(i)
    order_Media.Add LMediaList(i)
Next

' Create a Collection of Concentrate
For i = 2 To numConcList.Count Step 2
    order_Media.Add numConcList(i - 1)
    order_Media.Add numConcList(i)
    order_Media_Litre = mlConcList(i) / 1000  ' ml division
    order_Media.Add order_Media_Litre
Next


'******************************* 12 Month Data Collection *******************************

' Collection of each requests
Dim year_MediaRequest As New Collection
Dim year_MediaLitre As New Collection
Dim monthly_MediaRequest As New Collection
Dim monthly_MediaLitre As New Collection
DateNext = DateFrom

' Count up the 12 month
For n = 1 To 12

    ' Count through each order data
    For i = 1 To order_Media.Count Step 3

        mediaDate = order_Media(i)
        mediaRequest = order_Media(i + 1) 'Request
        mediaLitre = order_Media(i + 2) 'Litre
        
        ' Only collect data if its a matching month
        If Month(DateNext) = Month(mediaDate) And Not IsEmpty(mediaRequest) Then

            'Add Media Request Monthly
            monthly_MediaRequest.Add mediaRequest
            
            'Add Media Litre Monthly
            monthly_MediaLitre.Add mediaRequest
            monthly_MediaLitre.Add mediaLitre
        End If
    
    Next i
    
    'MsgBox monthly_MediaType.Count
    'MsgBox monthly_MediaLitre.Count

    ' Create a 'Year List' of the Months
    year_MediaRequest.Add monthly_MediaRequest
    year_MediaLitre.Add monthly_MediaLitre
    
    'Reset the month collection
    Set monthly_MediaRequest = New Collection ' Reset the Monthly Request List
    Set monthly_MediaLitre = New Collection ' Reset the Monthly Litre List
    DateNext = DateAdd("m", 1, DateNext) ' Find the next month


Next n


' ******************************* Media Request Sheet *******************************

'** Find the List of the Media

'** Move to the Media Request Sheet
Sheets(sheetName).Select

'Declare the Media List Collection
Dim MediaList As New Collection

' ** Determine Whether the year exists
i = StartRow
Do While Cells(i, 1).Value <> "Total"
    'your code here
    mediaType = Cells(i, 1).Value
    'MsgBox mediaType
    MediaList.Add mediaType
    i = i + 1
Loop

' Find the distance between the two table
distance = StartRow + MediaList.Count + 4

'**** Begin Counting and entering request of each media per month
Dim countRequest As Integer
Dim countLitre As Double


' Loop Through Months
For k = 2 To 13

    ' Loop Through list of Media
    For EachMedia = 1 To MediaList.Count
    
        'Set default values
        media = MediaList(EachMedia)
        countRequest = 0
        countLitre = 0
        rowNumLitre = StartRow + EachMedia - 1
        rowNumRequest = distance + EachMedia - 1
        'MsgBox media
        
    
         ' Loop through the list of media in each given month
         For i = 1 To year_MediaLitre(k - 1).Count Step 2
     
             ' Count if the media matches
             If media = year_MediaLitre(k - 1)(i) Then
                 countRequest = countRequest + 1  ' Media Request
                 countLitre = countLitre + year_MediaLitre(k - 1)(i + 1)  ' Media Litre
             End If
         Next i
        
        ' Locate the entry of the data
        Cells(rowNumLitre, k) = countLitre
        Cells(rowNumRequest, k) = countRequest
        
    Next EachMedia
    
Next k



'SUMMATION formula
sumFormulaRowLitre = StartRow + MediaList.Count
sumFormulaRowReque = StartRow + MediaList.Count + 4 + MediaList.Count

' Loop Through Months
For k = 2 To 13

    'If its the Last Row add the SUMMATION formula
    Cells(sumFormulaRowLitre, k) = "=SUM(" & Cells(StartRow, k).Address(False, False) & ":" & Cells(StartRow + MediaList.Count - 1, k).Address(False, False) & ")"
    Cells(sumFormulaRowReque, k) = "=SUM(" & Cells(distance, k).Address(False, False) & ":" & Cells(distance + MediaList.Count - 1, k).Address(False, False) & ")"
    
Next k

End Sub
