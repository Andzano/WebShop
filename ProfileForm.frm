VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProfileForm 
   Caption         =   "Profile"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10485
   OleObjectBlob   =   "ProfileForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProfileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackButton_Click()
    ProfileModule.BackToCatalog
End Sub

Private Sub CommandButton1_Click()
    ProfileModule.EditInfo
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
