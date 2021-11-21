Attribute VB_Name = "Module2"
Public Sub checkFlg()
     
    Dim i, startY As Integer
    Worksheets("sample").Select
    
    startX = 3
    startY = 2
    
    Do While Cells(startX, startY).Value = 1
        'arbitrary condition
        'Debug.Print Cells(startX, startY + 1)
         startX = startX + 1
    Loop
End Sub
Public Function concatArr(ByVal pkCollection As collection, ByVal delimter As String) As String
    Dim pkCode As String
    Dim comma As String
    Dim collectionCnt As Integer
    
    collectionCnt = 0
    For Each eachCollection In pkCollection
        collectionCount = collectionCount + i
        If collectionCount < pkCollection.Count Then
            comma = delimiter
        Else
            comma = ""
        End If
        
        pkCode = pkCode + eachCollection + comma
        
    Next eachCollection
    
    concatArr = pkCode
    
End Function
