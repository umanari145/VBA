Attribute VB_Name = "Module2"
Sub �ėp���\�b�h()
     
    Dim i, startY As Integer
    Worksheets("�T���v��").Select
    
    startX = 3
    startY = 2
    
    Do While Cells(startX, startY).Value = 1
        '�����ɔC�ӂ̏�����������
        'Debug.Print Cells(startX, startY + 1)
         startX = startX + 1
    Loop
End Sub
