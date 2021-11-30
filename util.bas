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
Public Function concatArr(ByVal pkCollection As Collection, ByVal delimter As String) As String
    Dim pkCode As String
    Dim comma As String
    Dim collectionCnt As Integer
    
    collectionCnt = 0
    For Each eachCollection In pkCollection
        collectionCount = collectionCount + 1
        If collectionCount < pkCollection.Count Then
            comma = delimter
        Else
            comma = ""
        End If
        
        If IsNumeric(eachCollection) = True Then
            pkCode = pkCode + Str(eachCollection) + comma
        Else
            pkCode = pkCode + eachCollection + comma
        End If
        
        
    Next eachCollection
    
    concatArr = pkCode
    
End Function
Public Function inArray(ByVal objCol As Collection, ByVal item As String) As Boolean
     
    isExist = False
    
    If objCol.Count > 0 Then
        
        For Each eachItem In objCol
            If eachItem = item Then
                isExist = True
            End If
        Next
    
    End If
    
    inArray = isExist
End Function
