Attribute VB_Name = "Encrypt"
    Public Function encription(original As String, operator As Boolean, clef As String)
        If Len(clef) = 6 Then
            Dim i, j, k As Long
            Dim a1, a2, b1, b2, temp, intdecrypt(0 To 35) As Byte
            Dim str, result, strdecrypt(0 To 5, 0 To 5) As String
     
            For i = 0 To 5
                intdecrypt(i) = Asc(Mid(LCase(clef), i + 1, 1))
            Next i
     
            temp = 97
     
            For i = 6 To 35
                Do Until intdecrypt(i) <> 0
                    If Not (intdecrypt(0) = temp Or intdecrypt(1) = temp Or intdecrypt(2) = temp Or intdecrypt(3) = temp Or intdecrypt(4) = temp Or intdecrypt(5) = temp) Then intdecrypt(i) = temp
                    temp = temp + 1
                    If temp = 123 Then temp = 32
                    If temp = 35 Then temp = 39
                    If temp = 42 Then temp = 44
                    If temp = 47 Then temp = 63
                Loop
            Next i
     
            For i = 0 To 5
                For j = 0 To 5
                    strdecrypt(i, j) = Chr(intdecrypt((6 * i) + j))
                Next j
            Next i
     
            For i = 1 To Len(original)
                Select Case Asc(Mid(original, i, 1))
                Case 32 To 34, 39 To 41, 44 To 46, 63
                    str = str & Mid(original, i, 1)
                Case 65 To 90
                    str = str & Chr(Asc(Mid(original, i, 1)) + 32)
                Case 97 To 122
                    str = str & Mid(original, i, 1)
                End Select
            Next i
     
            If Len(str) Mod 2 = 1 Then str = str & "q"
     
            For i = 1 To Len(str) Step 2
                For j = 0 To 5
                    For k = 0 To 5
                        If Mid(str, i, 1) = strdecrypt(j, k) Then
                        a1 = j
                        a2 = k
                        End If
                    Next k
                Next j
                For j = 0 To 5
                    For k = 0 To 5
                        If Mid(str, i + 1, 1) = strdecrypt(j, k) Then
                        b1 = j
                        b2 = k
                        End If
                    Next k
                Next j
                    If operator = False Then
                        result = result & strdecrypt(a2, b1) & strdecrypt(b2, a1)
                    Else
                        result = result & strdecrypt(b2, a1) & strdecrypt(a2, b1)
                    End If
            Next i
     
            If Not result = "" Then
                If Mid(result, Len(result), 1) = "q" Then result = Left(result, Len(result) - 1)
            End If
     
            encription = result
        Else
            encription = "clef invalide"
        End If
    End Function
    
Public Sub test()
    Dim Encrypt As String
    Dim Decrypt As String

    Encrypt = encription("triofro", False, "abcdef")

    MsgBox Encrypt & "O"

    Decrypt = encription(Encrypt, True, "abcdef")

    MsgBox Decrypt & "/"
End Sub
