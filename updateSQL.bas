Attribute VB_Name = "Module5"
Sub updateQL()

Dim headCollection As New Collection
Dim updateKeyCollection As New Collection
Dim updateKeyNumCollection As New Collection
Dim dataCollection As New Collection

x = 1
y = 1

updateKeyCollection.Add "belong_soshiki"
updateKeyCollection.Add "member_num"

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
totalNum = headCollection.Count + updateKeyNumCollection.Count
For dataNum = 2 To 3
    Set updateDataCollection = New Collection
    Set updateHeadCollection = New Collection
    For x = 1 To totalNum
        If inArray(updateKeyNumCollection, x) = True Then
            updateHeadCollection.Add Cells(dataNum, x).Value
        Else
            updateDataCollection.Add Cells(dataNum, x).Value
        End If
    Next
Next



'
'lastSQL = concatArr(totalSQLCollection, vbCrLf)
'Open path For Append As #1
'     Print #1, lastSQL
'Close #1

End Sub
