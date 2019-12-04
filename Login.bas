Attribute VB_Name = "Login"
Public Sub Login()
    MsgBox ("Login")
    Verify
End Sub

Private Sub Verify()
    UserName = LoginForm.UsernameBox
    Password = LoginForm.PasswordBox
    Call ReadFile(UserName, Password)
    
End Sub

Public Sub ReadFile(UserName, Password)
'C:\Users\User\Documents\GitHub\WebShop
    Dim FilePath As String
    FilePath = ThisWorkbook.Path & "\login.txt"
    Open FilePath For Input As #1
    Do While Not EOF(1)
        Line Input #1, nameandpass
        UserInfo = Split(nameandpass, " ")
        If UserInfo(0) = UserName And UserInfo(1) = Password Then
            MsgBox ("You have logged in")
        Else
            MsgBox ("Incorect Username or Password")
        End If
    Loop
    Close #1
End Sub

