VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Catalog 
   Caption         =   "Catalog"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13230
   OleObjectBlob   =   "Catalog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Catalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Import_Click()
    CatMod.Import
End Sub

Private Sub Logout_Click()
    LoginForm.Label2.Visible = False
    LoginForm.UsernameBox.Value = ""
    LoginForm.PasswordBox.Value = ""
    Catalog.Import.Visible = False
    Catalog.Hide
    LoginForm.Show
End Sub

Private Sub UserForm_Initialize()
    'Add data to list
    CatMod.Init
End Sub
