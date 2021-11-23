Attribute VB_Name = "Module4"
Public Function eachMakeSQL(ByVal sheetName As String) As String
    Dim sql, tmp, tmp2, tmp3, tmpVal As String
    Dim headCollection As New Collection
    Dim dataCollection As New Collection
    Dim allDataCollection As New Collection
    Dim headSQL, dataSQL As String

    tmp = ""
    Worksheets(sheetName).Select
    sql = "truncate table " + sheetName + ";" & vbCrLf
    
    startX = 1
    startY = 1
     '   do column loop
    Do While Cells(startY, startX) <> ""
        
        rowX = 1
        'do row loop
        If startY = 1 Then
            'head
            Do While Cells(startY, rowX) <> ""
                headCollection.Add Cells(startY, rowX)
                rowX = rowX + 1
            Loop
        Else
            'data
            Set dataCollection = New Collection
            Do While Cells(startY, rowX) <> ""
                tmpVal = Cells(startY, rowX)
                If IsNumeric(tmpVal) = True Then
                    dataCollection.Add tmpVal
                Else
                    dataCollection.Add "'" + tmpVal + "'"
                End If
                rowX = rowX + 1
            Loop
             tmp = concatArr(dataCollection, ",")
             tmp2 = "(" + tmp + ")"
             allDataCollection.Add tmp2
        End If
        startY = startY + 1
    Loop
    
    headSQL = concatArr(headCollection, ",")
    dataSQL = concatArr(allDataCollection, ",")
    
    tmp3 = "INSERT INTO " + sheetName + " (" + headSQL + ")" + " VALUES " + dataSQL
    
    eachMakeSQL = tmp3
    
End Function
