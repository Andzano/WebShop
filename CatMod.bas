Attribute VB_Name = "CatMod"
Public Sub Init()

    'Puts all of data in masinas text file into catalog, clean slate
    RefreshCatalog

    'Opon Catalog from opening set Combobox values
    'combobox4 - atruma karbas
    With Catalog.ComboBox4
        .AddItem "automats"
        .AddItem "manuala"
    End With

    'lietojums
    With Catalog.ComboBox5
        .AddItem "lietota"
        .AddItem "jauna"
    End With

    'krasa
    With Catalog.ComboBox3
        .AddItem "balta"
        .AddItem "bruna"
        .AddItem "dzeltna"
        .AddItem "gaisi zila"
        .AddItem "melna"
        .AddItem "orandza"
        .AddItem "peleka"
        .AddItem "sarkana"
        .AddItem "sudraba"
        .AddItem "tumsi sarkana"
        .AddItem "violeta"
        .AddItem "zala"
        .AddItem "zila"
    End With

    'Adds Items to Car Model Combobox programicaly, using data in the list
    Catalog.ComboBox1.AddItem Catalog.ListBox1.List(0, 0)
    For k = 0 To Catalog.ListBox1.ListCount - 1
    Contains = False
        For i = 0 To Catalog.ComboBox1.ListCount - 1
            If Catalog.ComboBox1.List(i) = Catalog.ListBox1.List(k, 0) Then
                Contains = True
            End If
        Next i
        If Contains = False Then
            Catalog.ComboBox1.AddItem Catalog.ListBox1.List(k, 0)
        End If
    Next k
    
End Sub
'Puts all of data in masinas text file into catalog, clean slate
Public Sub RefreshCatalog()
    Catalog.ListBox1.Clear
    Call CatMod.ImportCatalog("", "", "", "", "", "", "", "")
End Sub

Public Sub ImportCatalog(Model, CarColor, Gear, Usage, PriceFrom, PriceTo, YearFrom, YearTo)
    Model = Model
    CarColor = CarColor
    Gear = Gear
    Usage = Usage
    PriceFrom = PriceFrom
    PriceTo = PriceTo
    YearFrom = YearFrom
    YearTo = YearTo

    Open ThisWorkbook.Path & "\masinas.txt" For Input As #1
    j = 0
    Do While Not EOF(1)
        Line Input #1, Line
        masina = Split(Line, "/")
        
        ArrayLenght = UBound(masina) - LBound(masina) + 1
        If ArrayLenght <> 7 Then
            GoTo NextIteration
        Else
            If Len(masina(1)) <> 4 Then
                GoTo NextIteration
            End If
        End If

        If Model = "" And CarColor = "" And Gear = "" And Usage = "" And PriceFrom = "" And PriceTo = "" And YearFrom = "" And YearTo = "" Then
            'If no filter is set then lead all data, like in case of just opening store catalog
            Catalog.ListBox1.AddItem
            For i = 0 To 6
                Catalog.ListBox1.List(j, i) = masina(i) 'row and column
            Next i
            j = j + 1
        Else
            'If Filter is set, then set filter Conditions accordingly to fields that ar filled
            Dim Condition(7)
            If Model <> "" Then
                Condition(0) = Model
            Else
                Condition(0) = masina(0)
            End If
            If CarColor <> "" Then
                Condition(1) = CarColor
            Else
                Condition(1) = masina(3)
            End If
            If Gear <> "" Then
                Condition(2) = Gear
            Else
                Condition(2) = masina(4)
            End If
            If Usage <> "" Then
                Condition(3) = Usage
            Else
                Condition(3) = masina(5)
            End If
            If PriceFrom <> "" Then
                Condition(4) = PriceFrom
            Else
                Condition(4) = masina(6)
            End If
            If PriceTo <> "" Then
                Condition(5) = PriceTo
            Else
                Condition(5) = masina(6)
            End If
            
            If YearFrom <> "" Then
                Condition(6) = YearFrom
            Else
                Condition(6) = masina(1)
            End If
            If YearTo <> "" Then
                Condition(7) = YearTo
            Else
                Condition(7) = masina(1)
            End If
            
            If masina(0) = Condition(0) And masina(3) = Condition(1) And masina(4) = Condition(2) And masina(5) = Condition(3) _
            And CLng(masina(6)) >= CLng(Condition(4)) And CLng(masina(6)) <= CLng(Condition(5)) And CLng(masina(1)) >= CLng(Condition(6)) And CLng(masina(1)) <= CLng(Condition(7)) Then
                Catalog.ListBox1.AddItem
                For n = 0 To 6
                    Catalog.ListBox1.List(k, n) = masina(n) 'row and column
                Next n
                k = k + 1
            End If
        End If
NextIteration:
    Loop
    Close #1
    
    'After Data has been put into list, save the list into text and refresh the list
    
    
    
'    Catalog.ListBox1.Clear
    
End Sub

'Public Sub ValidateLine()
'    'checks if line has 6 slashes "/" and also that there is correct
'    'year. if invalid, go to next line
'    CheckTheLine = Split(Line, "/")
'    ArrayLenght = UBound(CheckTheLine) - LBound(CheckTheLine) + 1
'    If ArrayLenght <> 7 Then
'        GoTo NextIteration
'    Else
'        If Len(CheckTheLine(1)) <> 4 Then
'            GoTo NextIteration
'        End If
'    End If
'End Sub

Public Sub Import()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    Dim filename As String
    If fd.Show = -1 Then
        j = Catalog.ListBox1.ListCount - 1
        filename = fd.SelectedItems(1)
        Open filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, Line
            CarMatch = False

            'checks if line has 6 slashes "/" and also that there is correct
            'year. if invalid, go to next line
            CheckTheLine = Split(Line, "/")
            ArrayLenght = UBound(CheckTheLine) - LBound(CheckTheLine) + 1
            If ArrayLenght <> 7 Then
                GoTo NextIteration
            Else
                If Len(CheckTheLine(1)) <> 4 Then
                    GoTo NextIteration
                End If
            End If


            'Check for duplicate lines, if found will skip line
            For k = 0 To Catalog.ListBox1.ListCount - 1
                If Line = Catalog.ListBox1.List(k, 0) & "/" & _
                Catalog.ListBox1.List(k, 1) & "/" & Catalog.ListBox1.List(k, 2) & "/" & _
                Catalog.ListBox1.List(k, 3) & "/" & Catalog.ListBox1.List(k, 4) & "/" & _
                Catalog.ListBox1.List(k, 5) & "/" & Catalog.ListBox1.List(k, 6) Then GoTo NextIteration
            Next k

            'if line is not duplicate, then add car to list
            If CarMatch = False Then 'No ducplicate car is found
            
                'Appends the text to existing list
'                Open ThisWorkbook.Path & "\masinas.txt" For Append As #1
'                Write #1, masina
'                Close #1
                
                'Adds text to Catalog
                masina = Split(Line, "/")
                Catalog.ListBox1.AddItem
                For i = 0 To 6
                    Catalog.ListBox1.List(Catalog.ListBox1.ListCount - 1, i) = masina(i)
                Next i
            End If

NextIteration:
        Loop
        Close #1
    End If
    
    'Function, that refreshes data in text file according to ListBox
    RefreshMasinasTextFile
    
End Sub

Public Sub RefreshMasinasTextFile()
    Open ThisWorkbook.Path & "\masinas.txt" For Output As #1
    For k = 0 To Catalog.ListBox1.ListCount - 1
        Line = Catalog.ListBox1.List(k, 0) & "/" & _
                Catalog.ListBox1.List(k, 1) & "/" & Catalog.ListBox1.List(k, 2) & "/" & _
                Catalog.ListBox1.List(k, 3) & "/" & Catalog.ListBox1.List(k, 4) & "/" & _
                Catalog.ListBox1.List(k, 5) & "/" & Catalog.ListBox1.List(k, 6)
        Print #1, Line
    Next k
    Close #1
End Sub

Public Sub Filter()
    Model = Catalog.ComboBox1
    CarColor = Catalog.ComboBox3
    Gear = Catalog.ComboBox4
    Usage = Catalog.ComboBox5
    PriceFrom = Catalog.TextBox1
    PriceTo = Catalog.TextBox2
    YearFrom = Catalog.TextBox3
    YearTo = Catalog.TextBox4

    'Clear the Main list box
    Catalog.ListBox1.Clear
    Call CatMod.ImportCatalog(Model, CarColor, Gear, Usage, PriceFrom, PriceTo, YearFrom, YearTo)
End Sub

Public Sub PreserveData(UserData)
    UserData = UserData
    BuyModule.PreserveData (UserData)
End Sub
