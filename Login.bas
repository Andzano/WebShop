Attribute VB_Name = "Login"
'Global variable that can be accesses across all Modules
Public GlobalUserData As Variant

Public Sub Login()
    Verify
End Sub

Private Sub Verify()

    Dim Password As String
    
    Username = LoginForm.UsernameBox
    Password = LoginForm.PasswordBox
    Call ReadFile(Username, Password)
    
End Sub

Public Sub ReadFile(Username, Password As String)
    Dim EncryptedPassword As String
    EncryptedPassword = Encrypt.encription(Password, False, "abcdef")
    User = False
    
    Open ThisWorkbook.Path & "\login.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, nameandpass
        UserInfo = Split(nameandpass, "/")
        
        If nameandpass <> "" Then
            If UserInfo(0) = Username Then
                If UserInfo(1) = EncryptedPassword Then
                    GlobalUserData = UserInfo
                    User = True
                End If
            End If
        End If
    Loop
    Close #1
    
    If User = False Then
        LoginForm.Label2.ForeColor = RGB(255, 0, 0)
        LoginForm.Label2.Caption = "Incorrect UserName/Password"
        LoginForm.Label2.Visible = True
    Else
        LoginForm.Hide
        With Catalog
            .BuyButton.Visible = True
            .EditButton.Visible = False
            .CommandButton1.Visible = True 'Profile Button
            .CatalogMessage.Visible = True
            .CatalogMessage.ForeColor = RGB(0, 255, 0)
            .CatalogMessage.Caption = "Welcome, " & GlobalUserData(2) & "!"
        End With
            CatMod.Init
            Catalog.Show
    End If
    
    If Username = "root" And Password = "root" Then
        LoginForm.Hide
        Dim RootArray(1)
        RootArray(0) = "root"
        GlobalUserData = RootArray
        With Catalog
            .CommandButton1.Visible = False 'Profile Button
            .ImportButton.Visible = True
            .BuyButton.Visible = False
            .EditButton.Visible = True
            .CatalogMessage.Visible = True
            .CatalogMessage.ForeColor = RGB(0, 0, 255)
            .CatalogMessage.Caption = "Loged in as Admin"
        End With
        CatMod.Init
        Catalog.Show
    End If
End Sub

