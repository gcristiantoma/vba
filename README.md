# vba
Sub step_test()

nrRiga = 1
nrows = 4263
x = 0
k = 1
For i = 1 To nrows + 500 Step 500

    txt = ""
    While k <> i
    If Range("c" & k).Value <> "" Then
        'Debug.Print Range("c" & k).Value
            If txt = "" Then
                txt = Range("c" & k).Value
                Else: txt = txt + ";" + Range("c" & k).Value
            End If
    Else: GoTo 0
    End If
0
        k = k + 1
    Wend
    Debug.Print txt
    Debug.Print "-------------------------------------------"
    'Range("d" & nrRiga).Value = txt
    nrRiga = nrRiga + 1
    txt = ""
Next i
End Sub
