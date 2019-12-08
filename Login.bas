Attribute VB_Name = "Login"
Public Sub Login()
'    MsgBox ("Login")
    Verify
End Sub

Private Sub Verify()

    Dim Password As String
    
    Username = LoginForm.UsernameBox
    Password = LoginForm.PasswordBox
    Call ReadFile(Username, Password)
    
End Sub

Public Sub ReadFile(Username, Password As String)

    Dim EncryptPassword As String

    EncryptPassword = Encrypt.encription("triofro", False, "abcdef")
'    MsgBox EncryptPassword & "after encryption"
    
    User = ""
'    Dim FilePath As String
'    FilePath = ThisWorkbook.Path & "\login.txt"
    Open ThisWorkbook.Path & "\login.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, nameandpass
        userInfo = Split(nameandpass, " ")
        
'        MsgBox userInfo(0)
'        MsgBox Username
        
'        MsgBox userInfo(1)
'        MsgBox EncryptedPassword
'        Dim DatabasePassword As String
'        DatabasePassword = userInfo(1)
        If userInfo(0) = Username Then
            If userInfo(1) = EncryptedPassword Then
                User = userInfo(0)
            End If
'            User = userInfo(0)
'            MsgBox ("Login successful")
        End If
    Loop
'   TODO what happend when no user is found
    If User = "" Then
'        MsgBox ("User Not Found")
    End If
    Close #1
End Sub

