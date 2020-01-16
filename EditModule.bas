Attribute VB_Name = "EditModule"
Public Sub Init()

Catalog.CatalogMessage.Caption = ""

'Gets selected item and stores it to array
 For x = 0 To Catalog.ListBox1.ListCount - 1
    If Catalog.ListBox1.Selected(x) = True Then
        Dim SelectedItem(6)
        For i = 0 To 6
            SelectedItem(i) = Catalog.ListBox1.List(x, i)
        Next i
    End If
 Next x

 For i = 1 To 7
    EditForm.Controls("TextBox" & i) = SelectedItem(i - 1)
 Next i
 
 'color
    With EditForm.TextBox4
        .AddItem "balta"
        .AddItem "bruna"
        .AddItem "dzeltna"
        .AddItem "gaisi zila"
        .AddItem "melna"
        .AddItem "orandza"
        .AddItem "peleka"
        .AddItem "sarkana"
        .AddItem "sudraba"
        .AddItem "tumsi sarkana"
        .AddItem "violeta"
        .AddItem "zala"
        .AddItem "zila"
    End With
 
    'gear
    With EditForm.TextBox5
        .AddItem "automats"
        .AddItem "manuala"
    End With

    'usage
    With EditForm.TextBox6
        .AddItem "lietota"
        .AddItem "jauna"
    End With
 
 Catalog.Hide
 EditForm.Show
End Sub

Public Sub Delete()
    'Gets selected item and stores it to string
    For x = 0 To Catalog.ListBox1.ListCount - 1
       If Catalog.ListBox1.Selected(x) = True Then
            CarToDelete = Catalog.ListBox1.List(x, 0) & "/" & _
            Catalog.ListBox1.List(x, 1) & "/" & Catalog.ListBox1.List(x, 2) & "/" & _
            Catalog.ListBox1.List(x, 3) & "/" & Catalog.ListBox1.List(x, 4) & "/" & _
            Catalog.ListBox1.List(x, 5) & "/" & Catalog.ListBox1.List(x, 6)
       End If
    Next x
    
    AppendDelete = CarToDelete & "/Deleted"
    
    'Stores all data into FileContent
    Dim TextFile As Integer
    Dim FileContent As String
    TextFile = FreeFile
    Open ThisWorkbook.Path & "\masinas.txt" For Input As TextFile
    FileContent = Input(LOF(TextFile), TextFile)
    Close TextFile
    
    'Replaces Line in FileContent and rewrites textfile
    FileContent = Replace(FileContent, CarToDelete, AppendDelete)
    TextFile = FreeFile
    Open ThisWorkbook.Path & "\masinas.txt" For Output As TextFile
    Print #TextFile, FileContent
    Close TextFile
    
    With Catalog
        .CatalogMessage.Caption "Car Deleted"
        .CartalogMessage.ForeColor = RGB(0, 0, 255)
    End With
    BackToCatalog
    
End Sub

Public Sub Edit()

    'First check that year is 4 char long
    If Len(EditForm.TextBox2) <> 4 Then
        EditForm.TextBox2.ForeColor = vbRed
        Exit Sub
    End If
    
    'Check that Model name is not empty
    If EditForm.TextBox1 = "" Then
        EditForm.TextBox1.ForeColor = vbRed
        Exit Sub
    End If
    
    'Check that Model name is not empty
    If EditForm.TextBox3 = "" Then
        EditForm.TextBox3.ForeColor = vbRed
        Exit Sub
    End If
    
    'Check that Model name is not empty
    If EditForm.TextBox7 = "" Then
        EditForm.TextBox7.ForeColor = vbRed
        Exit Sub
    End If
    
    'Gets selected item and stores it to string
    For x = 0 To Catalog.ListBox1.ListCount - 1
       If Catalog.ListBox1.Selected(x) = True Then
            CarToEditLine = Catalog.ListBox1.List(x, 0) & "/" & _
            Catalog.ListBox1.List(x, 1) & "/" & Catalog.ListBox1.List(x, 2) & "/" & _
            Catalog.ListBox1.List(x, 3) & "/" & Catalog.ListBox1.List(x, 4) & "/" & _
            Catalog.ListBox1.List(x, 5) & "/" & Catalog.ListBox1.List(x, 6)
       End If
    Next x
    
    'Gets data form TextBox input and makes a line with delimiters
    Dim EditedCar(6)
    For i = 1 To 7
        EditedCar(i - 1) = EditForm.Controls("TextBox" & i)
    Next i
    EditedCarLine = Join(EditedCar, "/")
    
    
    'Stores all data into FileContent
    Dim TextFile As Integer
    Dim FileContent As String
    TextFile = FreeFile
    Open ThisWorkbook.Path & "\masinas.txt" For Input As TextFile
    FileContent = Input(LOF(TextFile), TextFile)
    Close TextFile
    
    'Replaces Line in FileContent and rewrites textfile
    FileContent = Replace(FileContent, CarToEditLine, EditedCarLine)
    TextFile = FreeFile
    Open ThisWorkbook.Path & "\masinas.txt" For Output As TextFile
    Print #TextFile, FileContent
    Close TextFile
    
    With Catalog
        .CatalogMessage.Caption = "Car Edited"
        .CatalogMessage.ForeColor = RGB(0, 255, 0)
    End With
    BackToCatalog
    
End Sub

Public Sub Clear()
    EditForm.TextBox1.ForeColor = vbBlack
    EditForm.TextBox2.ForeColor = vbBlack
    EditForm.TextBox3.ForeColor = vbBlack
    EditForm.TextBox7.ForeColor = vbBlack
    EditForm.TextBox4.Clear
    EditForm.TextBox5.Clear
    EditForm.TextBox6.Clear
End Sub

Public Sub BackToCatalog()
    EditModule.Clear
    EditForm.Hide
    CatMod.Init
    Catalog.Show
End Sub
