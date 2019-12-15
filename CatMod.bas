Attribute VB_Name = "CatMod"
Public Sub Init()
    
    Open ThisWorkbook.Path & "\masinas.txt" For Input As #1
    j = 0
    Do While Not EOF(1)
        Line Input #1, Line
        masina = Split(Line, "/")
        Catalog.ListBox1.AddItem
        For i = 0 To 6
            Catalog.ListBox1.List(j, i) = masina(i) 'row and column
        Next i
        j = j + 1
    Loop
    Close #1
End Sub

Public Sub Import()
'    Const Rows = Catalog.ListBox1.ListCount - 1
'    Dim visasmasinas(Rows, 6) As Variant
'
'    For k = 0 To Catalog.ListBox1.ListCount - 1
'        For j = 0 To 6
'            visasmasinas(k, j) = Catalog.ListBox1.List(k, j)
'        Next j
'    Next k
    j = Catalog.ListBox1.ListCount - 1
    Open ThisWorkbook.Path & "\vw.txt" For Input As #1
    Do While Not EOF(1)
    
        Line Input #1, Line
        CarMatch = False
        
        For k = 0 To Catalog.ListBox1.ListCount - 1
        
            If Line = Catalog.ListBox1.List(k, 0) & "/" & _
            Catalog.ListBox1.List(k, 1) & "/" & Catalog.ListBox1.List(k, 2) & "/" & _
            Catalog.ListBox1.List(k, 3) & "/" & Catalog.ListBox1.List(k, 4) & "/" & _
            Catalog.ListBox1.List(k, 5) & "/" & Catalog.ListBox1.List(k, 6) Then
            
                CarMatch = True 'Duplicate car found
                Exit For
            End If
        Next k
        
        If CarMatch = False Then 'No ducplicate car is found
            masina = Split(Line, "/")
            Catalog.ListBox1.AddItem
            For i = 0 To 6
                Catalog.ListBox1.List(Catalog.ListBox1.ListCount - 1, i) = masina(i)
'                MsgBox "car added" & masina(6)
            Next i
        End If
        
    Loop
    Close #1
End Sub
