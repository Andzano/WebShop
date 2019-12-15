Attribute VB_Name = "Login"
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
        userInfo = Split(nameandpass, "/")
        
        If userInfo(0) = Username Then
            If userInfo(1) = EncryptedPassword Then
                User = userInfo(0)
                User = True
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
        Catalog.Show
    End If
    
    If Username = "root" And Password = "root" Then
        LoginForm.Hide
        Catalog.Import.Visible = True
        Catalog.Show
    End If
End Sub

