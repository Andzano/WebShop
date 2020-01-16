VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Excel Shop"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    'Calls Login Sub in Login Modules
    Login.Login
End Sub

Private Sub CommandButton2_Click()
    Register.Register
End Sub

Private Sub OptionButton1_Click()
    Lang.English
End Sub

Private Sub OptionButton2_Click()
    Lang.Latvian
End Sub

Private Sub UserForm_Terminate()
ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
