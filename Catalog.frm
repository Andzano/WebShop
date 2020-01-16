VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Catalog 
   Caption         =   "Catalog"
   ClientHeight    =   9660.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13485
   OleObjectBlob   =   "Catalog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Catalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BuyButton_Click()
    BuyModule.Init
End Sub

Private Sub CommandButton1_Click()
    Catalog.Hide
    ProfileModule.Init
    ProfileForm.Show
End Sub

Private Sub EditButton_Click()
    'Check if there is something selected, else show message
    For x = 0 To Catalog.ListBox1.ListCount - 1
    If Catalog.ListBox1.Selected(x) = True Then
        EditModule.Init
    Else
        Catalog.CatalogMessage.Visible = True
        Catalog.CatalogMessage.ForeColor = RGB(255, 0, 0)
        Catalog.CatalogMessage.Caption = "Please Select an item"
    End If
 Next x
    
End Sub

Private Sub FilterButton_Click()
    CatMod.Filter
End Sub

Private Sub ImportButton_Click()
    CatMod.Import
End Sub

Private Sub Logout_Click()
    LoginForm.Label2.Visible = False
    LoginForm.UsernameBox.Value = ""
    LoginForm.PasswordBox.Value = ""
    Catalog.ImportButton.Visible = False
    Catalog.ComboBox1 = ""
    Catalog.ComboBox3 = ""
    Catalog.ComboBox4 = ""
    Catalog.ComboBox5 = ""
    Catalog.TextBox1 = ""
    Catalog.TextBox2 = ""
    Catalog.TextBox3 = ""
    Catalog.TextBox4 = ""
    
    Dim EmptyArray(1)
    GlobalUserData = EmptyArray
    
    Catalog.Hide
    LoginForm.Show
End Sub
'
Private Sub UserForm_Initialize()
    'Add data to list
'    CatMod.Init
End Sub

Private Sub UserForm_Terminate()
    ThisWorkbook.Close savechanges:=False
    Application.Quit
End Sub
