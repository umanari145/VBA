Attribute VB_Name = "Module2"
Sub 汎用メソッド()
     
    Dim i, startY As Integer
    Worksheets("サンプル").Select
    
    startX = 3
    startY = 2
    
    Do While Cells(startX, startY).Value = 1
        'ここに任意の条件文を入れる
        'Debug.Print Cells(startX, startY + 1)
         startX = startX + 1
    Loop
End Sub
