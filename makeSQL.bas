Attribute VB_Name = "Module3"
Sub makeSQL()
'
' Macro1 Macro
'
Dim tableName, path As String
Dim startX, startY As Integer
Dim tableSQL As String
Dim totalSQLCollection As New Collection

'clear SQL file
fileName = ""
path = "/Users/matsumotonorio/Desktop/sampleSQL.sql"
Open path For Output As #1
    Print #1,
Close #1

Worksheets("dataList").Select
startY = 2
startX = 1
tableNameX = 3
For startY = 1 To 10
    If Cells(startY, startX).Value = 1 Then
        sheetName = Cells(startY, tableNameX).Value
        tableSQL = eachMakeSQL(sheetName)
        totalSQLCollection.Add tableSQL
    End If
    Worksheets("dataList").Select
Next

lastSQL = concatArr(totalSQLCollection, vbCrLf)
Open path For Append As #1
     Print #1, lastSQL
Close #1

End Sub
