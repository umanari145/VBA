Sub makeUpdateSQL()

Dim headCollection As New Collection
Dim updateKeyCollection As New Collection
Dim updateKeyNumCollection As New Collection
Dim dataCollection As New Collection
Dim sqlCollection As New Collection
Dim path, tmp, tableName As String

path = Cells(2, 7).Value
Open path For Output As #1
    Print #1,
Close #1

tableName = Cells(4, 8).Value
'tableName = ActiveSheet.Name
x = 1
y = 1

 'update Key
i = 0
Do While Cells(3, 8 + i) <> ""
    updateKeyCollection.Add Cells(3, 8 + i).Value
    i = i + 1
Loop


'head
Do While Cells(y, x) <> ""
    If inArray(updateKeyCollection, Cells(y, x).Value) = True Then
        updateKeyNumCollection.Add x
    Else
        headCollection.Add Cells(y, x).Value
    End If
    x = x + 1
Loop

x = 1
'data
totalNums = headCollection.Count + updateKeyNumCollection.Count
lineNum = countLineNum(2, totalNums)


For dataNum = 2 To lineNum
    Set updateDataCollection = New Collection
    Set updateHeadCollection = New Collection
    For x = 1 To totalNums
        If Cells(dataNum, x) <> "" Then
            tmp = Cells(1, x).Value & "='" & Cells(dataNum, x).Value & "'"
            If inArray(updateKeyNumCollection, x) = True Then
                'where
                updateHeadCollection.Add tmp
             Else
                 'data
                updateDataCollection.Add tmp
            End If
        End If
    Next

    whereSQL = concatArr(updateHeadCollection, " and ")
    dataSQL = concatArr(updateDataCollection, ",")
    eachSQL = "update " + tableName + " set " + dataSQL + " where " + whereSQL
    sqlCollection.Add eachSQL
Next

lastSQL = concatArr(sqlCollection, vbCrLf)
Open path For Append As #1
     Print #1, lastSQL
Close #1

End Sub

