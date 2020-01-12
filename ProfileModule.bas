Attribute VB_Name = "ProfileModule"
Public Sub Init()
    ProfileForm.InfoMessage = ""
    For i = 8 To 13
        ProfileForm.Controls("TextBox" & i) = GlobalUserData(i - 6)
    Next i
    
    Open ThisWorkbook.Path & "\masinas.txt" For Input As #1
    ItemsFound = ""
    Do While Not EOF(1)
        Line Input #1, Line
        CarAndCustomer = Split(Line, "/")
        
        ArrayLenght = UBound(CarAndCustomer) - LBound(CarAndCustomer) + 1
        If ArrayLenght <> 15 Then
            GoTo NextIteration
        Else
            If CarAndCustomer(7) <> GlobalUserData(0) Then
                GoTo NextIteration
            End If
        End If
        
        ItemsFound = ItemsFound & Line & "#"
     
NextIteration:
    Loop
    Close #1
    
    ProfileForm.Frame1.Visible = False
    
    If ItemsFound <> "" Then
        ProfileForm.Frame1.Visible = True
        ItemArray = Split(ItemsFound, "#")
        ItemArrayLength = UBound(ItemArray) - LBound(ItemArray)
        j = 0
        ProfileForm.BillTextBox = 0
        
        For i = 0 To ItemArrayLength - 1
            CarAndCustomerLine = ItemArray(i)
            CarAndCustomerLineArray = Split(CarAndCustomerLine, "/")
            ProfileForm.ListBox1.AddItem
            For k = 0 To 6
                ProfileForm.ListBox1.List(j, k) = CarAndCustomerLineArray(k) 'row and column
            Next k
            ProfileForm.BillTextBox = ProfileForm.BillTextBox + CInt(CarAndCustomerLineArray(6)) * 1
            j = j + 1
        Next i
    End If
End Sub

Public Sub EditInfo()
    CurrentCustomerInfo = GlobalUserData(0) & "/" & GlobalUserData(1) & "/" & GlobalUserData(2) & "/" & GlobalUserData(3) & "/" & GlobalUserData(4) & "/" & GlobalUserData(5) & "/" & GlobalUserData(6) & "/" & GlobalUserData(7)
    NewCustomerInfo = GlobalUserData(0) & "/" & GlobalUserData(1) & "/" & ProfileForm.TextBox8 & "/" & ProfileForm.TextBox9 & "/" & ProfileForm.TextBox10 & "/" & ProfileForm.TextBox11 & "/" & ProfileForm.TextBox12 & "/" & ProfileForm.TextBox13
    
    Dim TextFile As Integer
    Dim FileContent As String
    TextFile = FreeFile
    Open ThisWorkbook.Path & "\login.txt" For Input As TextFile
    FileContent = Input(LOF(TextFile), TextFile)
    Close TextFile
    
    FileContent = Replace(FileContent, CurrentCustomerInfo, NewCustomerInfo)
    TextFile = FreeFile
    Open ThisWorkbook.Path & "\login.txt" For Output As TextFile
    Print #TextFile, FileContent
    Close TextFile
    
    NewUserData = Split(NewCustomerInfo, "/")
    GlobalUserData = NewUserData
    ProfileForm.InfoMessage.ForeColor = RGB(0, 255, 0)
    ProfileForm.InfoMessage = "Data Changed"
    ProfileForm.InfoMessage.Visible = True
End Sub

Public Sub BackToCatalog()
    ProfileForm.Hide
    Catalog.CatalogMessage.Caption = ""
    CatMod.Init
    Catalog.Show
End Sub
