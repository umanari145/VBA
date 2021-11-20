Attribute VB_Name = "Module1"
Sub �e�[�u��DDL�쐬()
Attribute �e�[�u��DDL�쐬.VB_ProcData.VB_Invoke_Func = " �n14"
'
' Macro1 Macro
'
Dim tableName, path As String
Dim startX, startY As Integer

'sql����U��ɂ���
fileName = ""
path = "/Users/matsumotonorio/Desktop/sampleDDL.sql"
Open path For Output As #1
    Print #1,
Close #1

Worksheets("�e�[�u���ꗗ").Select
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
    Dim pkCollection As New collection
    Dim pkStr, pkCode, pkComma, pkCodeLine As String

    pkCode = ""
    Worksheets(sheetName).Select
    butsuriTableName = Cells(1, 2).Value
    sql = "drop table " + butsuriTableName + ";" & vbCrLf
    sql = sql + "create table " + butsuriTableName + "(" + vbCrLf

    indexX = 4
    indexY = 1
     '  �c��̃��[�v
    Do While Cells(indexX, indexY) <> ""
        indexNo = Cells(indexX, indexY).Value
        
        ' ��������ɂ���
        ' j=3  �����J������
        ' j=4  �^���
        ' j=5 not null & PK  ���
        For j = 3 To 5
            tmp = Cells(indexX, j).Value
            If j = 5 Then
                If tmp = "Yes�iPK�j" Then
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
    
     ' PK �쐬
    pkStr = "constraint " + butsuriTableName + pkCodeLine + "_PKC " + " primar key (" + pkComma + ")" + vbCrLf
    sql = sql + pkStr + ");"
    
    Open path For Append As #1
        Print #1, sql
    Close #1
    
End Sub
Function concatArr(ByVal pkCollection As collection, ByVal delimter As String) As String
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
