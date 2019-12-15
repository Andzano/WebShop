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
        .Import.Caption = "Importet"
        .Logout.Caption = "Iziet"
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
        .Import.Caption = "Import"
        .Logout.Caption = "Logout"
    End With
End Sub
