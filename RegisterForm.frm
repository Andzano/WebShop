VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterForm 
   Caption         =   "Register"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "RegisterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
    LoginForm.Label2.Visible = False
    LoginForm.UsernameBox.Value = ""
    LoginForm.PasswordBox.Value = ""
    RegisterForm.Hide
    LoginForm.Show
End Sub

Private Sub Password_Change()

End Sub

Private Sub SetRegister_Click()
    Register.SetRegister
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
