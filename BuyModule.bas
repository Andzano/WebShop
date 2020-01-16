Attribute VB_Name = "BuyModule"
 
Public Sub Init()

Catalog.CatalogMessage.Caption = ""

'Gets selected item and stores it to array
 For x = 0 To Catalog.ListBox1.ListCount - 1
    If Catalog.ListBox1.Selected(x) = True Then
        Dim SelectedItem(6)
        For i = 0 To 6
            SelectedItem(i) = Catalog.ListBox1.List(x, i)
        Next i
    End If
 Next x
 
 For i = 1 To 7
    BuyForm.Controls("TextBox" & i) = SelectedItem(i - 1)
 Next i
 
 
 For i = 8 To 13
    BuyForm.Controls("TextBox" & i) = GlobalUserData(i - 6)
 Next i
 
 Catalog.Hide
 BuyForm.Show
End Sub

Public Sub Purchase()
Dim PurchasedCar(6)
Dim CustomerInfo(5)

For i = 1 To 7
    PurchasedCar(i - 1) = BuyForm.Controls("TextBox" & i)
Next i
SelectedCar = Join(PurchasedCar, "/")

For i = 8 To 13
    CustomerInfo(i - 8) = BuyForm.Controls("TextBox" & i)
Next i
SelectedCustomer = Join(CustomerInfo, "/")
CarAndCustomer = SelectedCar & "/" & SelectedCustomer

Dim TextFile As Integer
Dim FileContent As String
TextFile = FreeFile
Open ThisWorkbook.Path & "\masinas.txt" For Input As TextFile
FileContent = Input(LOF(TextFile), TextFile)
Close TextFile

FileContent = Replace(FileContent, SelectedCar, CarAndCustomer)
TextFile = FreeFile
Open ThisWorkbook.Path & "\masinas.txt" For Output As TextFile
Print #TextFile, FileContent
Close TextFile

Catalog.CatalogMessage.Caption = "Purchase Made"
BackToCatalog

End Sub

Public Sub BackToCatalog()
    BuyForm.Hide
    CatMod.Init
    Catalog.Show
End Sub
    
