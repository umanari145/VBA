Attribute VB_Name = "Module1"
Sub makeDDL()
'
' Macro1 Macro
'
Dim tableName, path As String
Dim startX, startY As Integer

'clear SQL file
Worksheets("tableList").Select
path = Cells(3, 5).Value
Open path For Output As #1
    Print #1,
Close #1

startX = 2
startY = 1
tableNameY = 2
Do While Cells(startX, startY).Value = 1
    sheetName = Cells(startX, tableNameY).Value
    Call eachSheetSQL(sheetName, path)
    startX = startX + 1
Loop

End Sub
Private Sub eachSheetSQL(ByVal sheetName As String, path As String)
    Dim i, j, k As Integer
    Dim indexY As Integer
    Dim sql, indexNo, tmp, tmp2 As String
    Dim pkCollection As New Collection
    Dim pkStr, pkCode, pkComma, pkCodeLine As String

    pkCode = ""
    Worksheets(sheetName).Select
    butsuriTableName = Cells(1, 2).Value
    sql = "drop table " + butsuriTableName + ";" & vbCrLf
    sql = sql + "create table " + butsuriTableName + "(" + vbCrLf

    indexX = 4
    indexY = 1
     '   do column loop
     
      
    Do While Cells(indexX, indexY) <> ""
        indexNo = Cells(indexX, indexY).Value
        
        'do row loop
        ' j=3  physics column
        ' j=4   table mold info
        ' j=5 not null & PK
        For j = 3 To 5
            tmp = Cells(indexX, j).Value
            If j = 5 Then
                If tmp = "Yes（PK）" Then
                    tmp2 = " NOT NULL ," & vbCrLf
                    pkCollection.Add Cells(indexX, 3)
                ElseIf tmp = "Yes" Then
                    tmp2 = " NOT NULL," & vbCrLf
                Else
                    tmp2 = "," & vbCrLf
                End If
            Else
                tmp2 = " " + tmp + " "
            End If
                sql = sql + tmp2
        Next
        indexX = indexX + 1
    Loop
    
    pkComma = concatArr(pkCollection, ",")
    pkCodeLine = concatArr(pkCollection, "_")
    
     ' make PK line
    pkStr = "constraint " + butsuriTableName + pkCodeLine + "_PKC " + " primar key (" + pkComma + ")" + vbCrLf
    sql = sql + pkStr + ");"
    
    Open path For Append As #1
        Print #1, sql
    Close #1
    
End Sub

