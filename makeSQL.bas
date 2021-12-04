Sub makeInsertSQL()
'
' Macro1 Macro
'
Dim tableName, path As String
Dim startX, startY, lot As Integer
Dim tableSQL As String
Dim totalSQLCollection As New Collection

Worksheets("dataList").Select
lot = Cells(4, 5).Value

'clear SQL file
fileName = ""
path = Cells(3, 5).Value
Open path For Output As #1
    Print #1,
Close #1

startY = 2
startX = 1
tableNameX = 3
For startY = 1 To 10
    If Cells(startY, startX).Value = 1 Then
        sheetName = Cells(startY, tableNameX).Value
        Call eachMakeSQL(sheetName, path, lot)
    End If
    Worksheets("dataList").Select
Next

End Sub
