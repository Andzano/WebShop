Attribute VB_Name = "Lang"

Public Sub Latvian()
    With LoginForm
        .Frame1.Caption = "Valoda"
        .OptionButton1.Caption = "Anglu"
        .OptionButton2.Caption = "Latviesu"
        .Caption = "Excel Veikals"
        .Label1.Caption = "Excel Veikals"
        .UsernameLabel.Caption = "Lietotajvards"
        .PasswordLabel.Caption = "Parole"
        .CommandButton1.Caption = "Ieiet"
        .CommandButton2.Caption = "Registreties"
    End With
    
    With RegisterForm
        .Caption = "Registracija"
        .UsernameLabel.Caption = "Lietotajvards"
        .PasswordLabel.Caption = "Parole"
        .RPasswordLabel.Caption = "Atkartot paroli"
        .Label1.Caption = "Vards"
        .Label2.Caption = "Uzvards"
        .Label3.Caption = "Personas Kods"
        .Label4.Caption = "Pilseta"
        .Label5.Caption = "Adrese"
        .Label6.Caption = "E-pasts"
        .Cancel.Caption = "Atcelt"
        .SetRegister.Caption = "Registreties"
    End With
    
    With Catalog
        .Caption = "Katalogs"
        .ImportButton.Caption = "Importet"
        .Logout.Caption = "Iziet"
        .CommandButton1.Caption = "Profils"
        .FilterButton.Caption = "Meklet"
        .Label1.Caption = "Modelis:"
        .Label3.Caption = "Krasa:"
        .Label4.Caption = "Atrumu karba:"
        .Label5.Caption = "Lietojums:"
        .Label6.Caption = "Cena:"
        .Label8.Caption = "Gads:"
        .BuyButton.Caption = "Pirkt"
        .EditButton.Caption = "Labot"
    End With
    
    With EditForm
        .Caption = "Labot"
        .Frame1.Caption = "Mashinas Detalas"
        .Label1.Caption = "Modelis"
        .Label2.Caption = "Gads"
        .Label3.Caption = "Marka"
        .Label4.Caption = "Krasa"
        .Label5.Caption = "Atrumu karba"
        .Label6.Caption = "Lietojums"
        .Label7.Caption = "Cena"
        .CommandButton1.Caption = "Dzest"
        .CommandButton2.Caption = "Labot"
        .CommandButton3.Caption = "Uz Katalogu"
    End With
    
    With BuyForm
        .Caption = "Pirkt"
        .Frame1.Caption = "Mashinas Detalas"
        .Label1.Caption = "Modelis"
        .Label2.Caption = "Gads"
        .Label3.Caption = "Marka"
        .Label4.Caption = "Krasa"
        .Label5.Caption = "Atrumu karba"
        .Label6.Caption = "Lietojums"
        .Label7.Caption = "Cena"
        .Frame2.Caption = "Pirceja Info"
        .Label8.Caption = "Vards"
        .Label9.Caption = "Uzvards"
        .Label10.Caption = "Personas Kods"
        .Label11.Caption = "Pilseta"
        .Label12.Caption = "Adrese"
        .Label13.Caption = "Epasts"
        .BackButton.Caption = "Uz Katalogu"
        .ConfirmButton.Caption = "Veikt Pirkumu"
    End With
    
    With ProfileForm
        .Caption = "Profile"
        .Frame1.Caption = "Bought Cars"
        .BillLabel.Caption = "Total:"
        .Frame2.Caption = "Pirceja Info"
        .Label8.Caption = "Vards"
        .Label9.Caption = "Uzvards"
        .Label10.Caption = "Personas Kods"
        .Label11.Caption = "Pilseta"
        .Label12.Caption = "Adrese"
        .Label13.Caption = "Epasts"
        .InfoMessage.Caption = "Dati laboti"
        .CommandButton1.Caption = "Labot Profilu"
        .BackButton.Caption = "Uz Katalogu"
    End With
End Sub

Public Sub English()
    With LoginForm
        .Frame1.Caption = "Language"
        .OptionButton1.Caption = "English"
        .OptionButton2.Caption = "Latvian"
        .Caption = "Excel Shop"
        .Label1.Caption = "Excel Shop"
        .UsernameLabel.Caption = "Username"
        .PasswordLabel.Caption = "Password"
        .CommandButton1.Caption = "Login"
        .CommandButton2.Caption = "Register"
    End With
    
    With RegisterForm
        .Caption = "Registration"
        .UsernameLabel.Caption = "Username"
        .PasswordLabel.Caption = "Password"
        .RPasswordLabel.Caption = "Repeat Password"
        .Label1.Caption = "Name"
        .Label2.Caption = "Surname"
        .Label3.Caption = "Personal code"
        .Label4.Caption = "City"
        .Label5.Caption = "Address"
        .Label6.Caption = "E-mail"
        .Cancel.Caption = "Cancel"
        .SetRegister.Caption = "Register"
    End With
    
    With Catalog
        .Caption = "Catalog"
        .ImportButton.Caption = "Import"
        .Logout.Caption = "Logout"
        .CommandButton1.Caption = "Profile"
        .FilterButton.Caption = "Meklet"
        .Label1.Caption = "Model:"
        .Label3.Caption = "Color:"
        .Label4.Caption = "Gear:"
        .Label5.Caption = "Usage:"
        .Label6.Caption = "Price:"
        .Label8.Caption = "Year:"
        .BuyButton.Caption = "Buy"
        .EditButton.Caption = "Edit"
    End With
    
    With EditForm
        .Caption = "Edit"
        .Frame1.Caption = "Car Details"
        .Label1.Caption = "Model"
        .Label2.Caption = "Year"
        .Label3.Caption = "Mark"
        .Label4.Caption = "Color"
        .Label5.Caption = "Gear"
        .Label6.Caption = "Usage"
        .Label7.Caption = "Price"
        .CommandButton1.Caption = "Delete"
        .CommandButton2.Caption = "Edit"
        .CommandButton3.Caption = "Back To Catalog"
    End With
    
    With BuyForm
        .Caption = "Buy"
        .Frame1.Caption = "Car Details"
        .Label1.Caption = "Model"
        .Label2.Caption = "Year"
        .Label3.Caption = "Mark"
        .Label4.Caption = "Color"
        .Label5.Caption = "Gear"
        .Label6.Caption = "Usage"
        .Label7.Caption = "Price"
        .Frame2.Caption = "Customer Details"
        .Label8.Caption = "Name"
        .Label9.Caption = "Surname"
        .Label10.Caption = "Personal Code"
        .Label11.Caption = "City"
        .Label12.Caption = "Address"
        .Label13.Caption = "Email"
        .BackButton.Caption = "Back To Catalog"
        .ConfirmButton.Caption = "Confirm Purchase"
    End With
    
    With ProfileForm
        .Caption = "Profile"
        .Frame1.Caption = "Bought Cars"
        .BillLabel.Caption = "Total:"
        .Frame2.Caption = "Customer Details"
        .Label8.Caption = "Name"
        .Label9.Caption = "Surname"
        .Label10.Caption = "Personal Code"
        .Label11.Caption = "City"
        .Label12.Caption = "Address"
        .Label13.Caption = "Email"
        .InfoMessage.Caption = "Data Changed"
        .CommandButton1.Caption = "Edit Profile"
        .BackButton.Caption = "Back To Catalog"
    End With
End Sub
