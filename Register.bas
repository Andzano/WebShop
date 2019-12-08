Attribute VB_Name = "Register"
Public Sub Register()
'    MsgBox ("Register")
' TODO izveidot funkciju kas parbauda vai lietotaajs neeksistee
    LoginForm.Hide
    RegisterForm.Show
End Sub

Public Sub SetRegister()
'    MsgBox ("Register")

    Dim Password As String
        
    Username = RegisterForm.Username
    Password = RegisterForm.Password
    Name = RegisterForm.PersonName
    Surname = RegisterForm.Surname
    PersonalCode = RegisterForm.PerosonalCode
    City = RegisterForm.City
    Address = RegisterForm.Address
    Email = RegisterForm.Email
    
'    MsgBox Password & "BEfore calling registeruser"
    
    Call RegisterUser(Username, Password, Name, Surname, PersonalCode, City, Address, Email)
    
    RegisterForm.Hide
    LoginForm.Show
    
End Sub

Private Sub RegisterUser(Username, Password As String, Name, Surname, PersonalCode, City, Address, Email)

    Dim EncryptedPassword As String

    EncryptedPassword = Encrypt.encription(Password, False, "abcdef")
'    MsgBox EncryptedPassword & "after encryption"
    
    Open ThisWorkbook.Path & "\login.txt" For Append As #1
    
    userInfo = Username & " " & EncryptedPassword & " " & Name & " " & Surname & " " & PersonalCode & " " & City & " " & Address & " " & Email
    
    Print #1, userInfo
    Close #1
End Sub
