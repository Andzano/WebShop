VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditForm 
   Caption         =   "Edit"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "EditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    EditModule.Delete
End Sub

Private Sub CommandButton2_Click()
    EditModule.Edit
End Sub

Private Sub CommandButton3_Click()
    EditModule.BackToCatalog
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
