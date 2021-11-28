Attribute VB_Name = "Module4"
Public Function eachMakeSQL(ByVal sheetName As String) As String
    Dim dataX, dataY, limitCount As Integer
    Dim sql, tmp, tmp2, tmp3, tmpVal As String
    Dim headCollection As New Collection
    Dim dataCollection As New Collection
    Dim allDataCollection As New Collection
    Dim headSQL, dataSQL As String

    tmp = ""
    Worksheets(sheetName).Select
    sql = "truncate table " + sheetName + ";" & vbCrLf
    
     '   do column loop
    
    'do row loop
    'head
    headX = 1
    headY = 1
    Do While Cells(headY, headX) <> ""
        headCollection.Add Cells(headY, headX)
        headX = headX + 1
    Loop
            
    For dataY = 2 To 999
        'data
        limitCount = headCollection.Count
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
        y = y + 1
    Next
    
    headSQL = concatArr(headCollection, ",")
    dataSQL = concatArr(allDataCollection, ",")
    
    tmp3 = "INSERT INTO " + sheetName + " (" + headSQL + ")" + " VALUES " + dataSQL
    
    eachMakeSQL = sql + tmp3
    
End Function
