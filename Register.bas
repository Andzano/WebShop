Attribute VB_Name = "Register"
Public Sub Register()
    LoginForm.Hide
    RegisterForm.Message.Visible = False
    RegisterForm.Show
End Sub

Public Sub SetRegister()

    Dim Password As String
        
    Username = RegisterForm.Username
    Password = RegisterForm.Password
    RPassword = RegisterForm.RPassword
    Name = RegisterForm.PersonName
    Surname = RegisterForm.Surname
    PersonalCode = RegisterForm.PerosonalCode
    City = RegisterForm.City
    Address = RegisterForm.Address
    Email = RegisterForm.Email
    
    
    If Password = RPassword Then
        Call RegisterUser(Username, Password, Name, Surname, PersonalCode, City, Address, Email)
        RegisterForm.Hide
        LoginForm.Label2.ForeColor = RGB(0, 255, 0)
        LoginForm.Label2.Caption = "User Registered"
        LoginForm.Label2.Visible = True
        LoginForm.Show
    End If
    
    If Password <> RPassword Then
        RegisterForm.Message.ForeColor = RGB(255, 0, 0)
        RegisterForm.Message.Caption = "Passwords must match"
        RegisterForm.Message.Visible = True
    End If

    
End Sub

Private Sub RegisterUser(Username, Password As String, Name, Surname, PersonalCode, City, Address, Email)

    Dim EncryptedPassword As String

    EncryptedPassword = Encrypt.encription(Password, False, "abcdef")
    
    Open ThisWorkbook.Path & "\login.txt" For Append As #1
    
    userInfo = Username & "/" & EncryptedPassword & "/" & Name & "/" & Surname & "/" & PersonalCode & "/" & City & "/" & Address & "/" & Email
    
    Print #1, userInfo
    Close #1
End Sub
