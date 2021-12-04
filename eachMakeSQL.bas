Public Sub eachMakeSQL(ByVal sheetName As String, ByVal path As String, ByVal lot As Integer)
    Dim dataX, dataY, limitCount As Integer
    Dim SQL, tmp, tmp2, tmp3, tmpVal As String
    Dim headCollection As New Collection
    Dim dataCollection As New Collection
    Dim allDataCollection As New Collection
    Dim headSQL, dataSQL As String
    
    tmp = ""
    Worksheets(sheetName).Select
    SQL = "truncate table " + sheetName + ";" & vbCrLf
    
     '   do column loop
    
    'do row loop
    'head
    headX = 1
    headY = 1
    Do While Cells(headY, headX) <> ""
        headCollection.Add Cells(headY, headX)
        headX = headX + 1
    Loop
            
    For dataY = 2 To 99999
        'data
        limitCount = headCollection.Count
        'initialize
        Set dataCollection = New Collection
        isData = False
        For dataX = 1 To limitCount
            tmpVal = Cells(dataY, dataX)
            If tmpVal <> "" Then
                isData = True
            End If
            dataCollection.Add "'" + tmpVal + "'"
        Next
        
        If isData = False Then
            Exit For
        End If
        
        tmp = concatArr(dataCollection, ",")
        tmp2 = "(" + tmp + ")"
        allDataCollection.Add tmp2
        
        ' save for memory saving
        If (dataY - 1) Mod lot = 0 Then
            Call saveOnePeriodSQL(headCollection, allDataCollection, path)
            Set allDataCollection = New Collection
        End If
    Next
    
    Call saveOnePeriodSQL(headCollection, allDataCollection, path)
    
End Sub
Public Sub saveOnePeriodSQL(ByVal headCollection As Collection, ByVal allDataCollection As Collection, ByVal path As String)
    Dim lastSQL As String
    
    headSQL = concatArr(headCollection, ",")
    dataSQL = concatArr(allDataCollection, ",")
    
    lastSQL = "INSERT INTO " + sheetName + " (" + headSQL + ")" + " VALUES " + dataSQL
        
    Open path For Append As #1
        Print #1, lastSQL
    Close #1
    
End Sub
